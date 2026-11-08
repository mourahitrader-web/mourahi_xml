// XML Cache Manager - Cache IndexedDB pour les fichiers XML déjà parsés

class XmlCache {
    constructor(databaseName = 'BigSi_XML_Cache', version = 1) {
        this.databaseName = databaseName;
        this.version = version;
        this.db = null;
        this.initPromise = null;
    }

    // Initialiser la base de données
    async init() {
        if (this.db) return this.db;
        if (this.initPromise) return this.initPromise;

        this.initPromise = new Promise((resolve, reject) => {
            const request = indexedDB.open(this.databaseName, this.version);

            request.onerror = () => {
                console.warn('Impossible d\'initialiser le cache XML:', request.error);
                this.db = null;
                resolve(null);
            };

            request.onsuccess = () => {
                this.db = request.result;
                resolve(this.db);
            };

            request.onupgradeneeded = event => {
                const db = event.target.result;
                if (!db.objectStoreNames.contains('xmlFiles')) {
                    db.createObjectStore('xmlFiles', { 
                        keyPath: 'fileKey',
                        autoIncrement: false 
                    });
                }
                if (!db.objectStoreNames.contains('fileMetadata')) {
                    db.createObjectStore('fileMetadata', { 
                        keyPath: 'fileName',
                        autoIncrement: false 
                    });
                }
            };
        });

        return this.initPromise;
    }

    // Générer une clé unique pour un fichier
    generateFileKey(fileName, fileSize, lastModified) {
        // Utiliser un hash simple basé sur les métadonnées du fichier
        return `${fileName}_${fileSize}_${lastModified}`;
    }

    // Sauvegarder un fichier parsé dans le cache
    async saveToCache(fileName, fileSize, lastModified, xmlString, tableMap) {
        const db = await this.init();
        if (!db) return false;

        try {
            const fileKey = this.generateFileKey(fileName, fileSize, lastModified);
            
            // Sauvegarder les métadonnées
            const metadata = {
                fileName,
                fileSize,
                lastModified,
                cachedAt: Date.now(),
                tableNames: Object.keys(tableMap),
                rowCount: Object.values(tableMap).reduce((sum, rows) => sum + rows.length, 0)
            };

            // Sauvegarder dans IndexedDB
            const transaction = db.transaction(['xmlFiles', 'fileMetadata'], 'readwrite');
            const filesStore = transaction.objectStore('xmlFiles');
            const metadataStore = transaction.objectStore('fileMetadata');

            // Stocker le tableMap (sérialisé)
            filesStore.put({
                fileKey,
                fileName,
                tableMap,
                cachedAt: Date.now()
            });

            // Stocker les métadonnées
            metadataStore.put(metadata);

            await new Promise((resolve, reject) => {
                transaction.oncomplete = resolve;
                transaction.onerror = reject;
            });

            return true;
        } catch (error) {
            console.warn('Erreur lors de la sauvegarde dans le cache:', error);
            return false;
        }
    }

    // Charger un fichier depuis le cache
    async loadFromCache(fileName, fileSize, lastModified) {
        const db = await this.init();
        if (!db) return null;

        try {
            const fileKey = this.generateFileKey(fileName, fileSize, lastModified);
            const transaction = db.transaction(['xmlFiles'], 'readonly');
            const store = transaction.objectStore('xmlFiles');
            const request = store.get(fileKey);

            const result = await new Promise((resolve, reject) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = reject;
            });

            return result ? result.tableMap : null;
        } catch (error) {
            console.warn('Erreur lors du chargement depuis le cache:', error);
            return null;
        }
    }

    // Vérifier si un fichier est dans le cache
    async isInCache(fileName, fileSize, lastModified) {
        const db = await this.init();
        if (!db) return false;

        try {
            const fileKey = this.generateFileKey(fileName, fileSize, lastModified);
            const transaction = db.transaction(['xmlFiles'], 'readonly');
            const store = transaction.objectStore('xmlFiles');
            const request = store.get(fileKey);

            const result = await new Promise((resolve, reject) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = reject;
            });

            return !!result;
        } catch (error) {
            return false;
        }
    }

    // Lister tous les fichiers en cache
    async listCachedFiles() {
        const db = await this.init();
        if (!db) return [];

        try {
            const transaction = db.transaction(['fileMetadata'], 'readonly');
            const store = transaction.objectStore('fileMetadata');
            const request = store.getAll();

            const results = await new Promise((resolve, reject) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = reject;
            });

            return results;
        } catch (error) {
            return [];
        }
    }

    // Effacer un fichier du cache
    async removeFromCache(fileName, fileSize, lastModified) {
        const db = await this.init();
        if (!db) return false;

        try {
            const fileKey = this.generateFileKey(fileName, fileSize, lastModified);
            const transaction = db.transaction(['xmlFiles', 'fileMetadata'], 'readwrite');
            const filesStore = transaction.objectStore('xmlFiles');
            const metadataStore = transaction.objectStore('fileMetadata');

            filesStore.delete(fileKey);
            metadataStore.delete(fileName);

            await new Promise((resolve, reject) => {
                transaction.oncomplete = resolve;
                transaction.onerror = reject;
            });

            return true;
        } catch (error) {
            console.warn('Erreur lors de la suppression du cache:', error);
            return false;
        }
    }

    // Effacer tout le cache
    async clearCache() {
        const db = await this.init();
        if (!db) return false;

        try {
            const transaction = db.transaction(['xmlFiles', 'fileMetadata'], 'readwrite');
            const filesStore = transaction.objectStore('xmlFiles');
            const metadataStore = transaction.objectStore('fileMetadata');

            filesStore.clear();
            metadataStore.clear();

            await new Promise((resolve, reject) => {
                transaction.oncomplete = resolve;
                transaction.onerror = reject;
            });

            return true;
        } catch (error) {
            console.warn('Erreur lors du vidage du cache:', error);
            return false;
        }
    }

    // Nettoyer les anciens fichiers du cache (plus vieux que X jours)
    async cleanupOldFiles(maxAgeDays = 30) {
        const db = await this.init();
        if (!db) return 0;

        try {
            const maxAgeMs = maxAgeDays * 24 * 60 * 60 * 1000;
            const now = Date.now();
            const transaction = db.transaction(['xmlFiles'], 'readwrite');
            const store = transaction.objectStore('xmlFiles');
            const request = store.getAll();

            const allFiles = await new Promise((resolve, reject) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = reject;
            });

            let deletedCount = 0;
            for (const file of allFiles) {
                if (file.cachedAt && (now - file.cachedAt) > maxAgeMs) {
                    store.delete(file.fileKey);
                    deletedCount++;
                }
            }

            await new Promise((resolve, reject) => {
                transaction.oncomplete = resolve;
                transaction.onerror = reject;
            });

            return deletedCount;
        } catch (error) {
            console.warn('Erreur lors du nettoyage du cache:', error);
            return 0;
        }
    }
}

// Instance globale du cache
const xmlCache = new XmlCache();
