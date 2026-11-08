// XML Parser Manager - Gestion optimisée du parsing XML avec Web Workers et cache

class XmlParserManager {
    constructor() {
        this.worker = null;
        this.workerUrl = null;
        this.pendingPromises = new Map();
        this.workerBusy = false;
        this.workerQueue = [];
        this.cache = null;
    }

    // Initialiser le gestionnaire
    init() {
        this.cache = xmlCache;
        this._createWorker();
    }

    // Créer un Web Worker
    _createWorker() {
        if (typeof Worker === 'undefined' || typeof Blob === 'undefined' || !window.URL) {
            this.worker = null;
            return;
        }

        try {
            // Utiliser le worker dédié
            this.workerUrl = 'js/xml-parser-worker.js';
            this.worker = new Worker(this.workerUrl);
            
            this.worker.onmessage = event => {
                const { tableMap, error, fileName, chunkIndex, totalChunks } = event.data;
                const promiseId = event.data.promiseId || fileName;
                
                const promise = this.pendingPromises.get(promiseId);
                if (promise) {
                    this.pendingPromises.delete(promiseId);
                    
                    if (error) {
                        promise.reject(new Error(error));
                    } else {
                        promise.resolve({ tableMap, fileName });
                    }
                }
                
                this.workerBusy = false;
                this._processNextInQueue();
            };

            this.worker.onerror = event => {
                console.error('Erreur du Web Worker:', event.message);
                // Recréer le worker
                this._cleanupWorker();
                this._createWorker();
            };
        } catch (error) {
            console.warn('Impossible de créer le Web Worker:', error);
            this.worker = null;
        }
    }

    // Nettoyer le worker
    _cleanupWorker() {
        if (this.worker) {
            this.worker.terminate();
            if (this.workerUrl) {
                URL.revokeObjectURL(this.workerUrl);
            }
            this.worker = null;
            this.workerUrl = null;
        }
    }

    // Traiter la prochaine tâche dans la file d'attente
    _processNextInQueue() {
        if (this.workerBusy || this.workerQueue.length === 0) return;
        
        const nextTask = this.workerQueue.shift();
        this.workerBusy = true;
        this.worker.postMessage(nextTask);
    }

    // Parser un fichier XML avec optimisations
    async parseXmlFile(file, options = {}) {
        const {
            useCache = true,
            forceReparse = false,
            showProgress = false,
            onProgress = null
        } = options;

        const fileName = file.name;
        const fileSize = file.size;
        const lastModified = file.lastModified;

        // Vérifier le cache d'abord
        if (useCache && !forceReparse) {
            const cachedTableMap = await this.cache.loadFromCache(fileName, fileSize, lastModified);
            if (cachedTableMap) {
                if (onProgress) onProgress(100);
                return {
                    tableMap: cachedTableMap,
                    fromCache: true,
                    fileName
                };
            }
        }

        // Lire le fichier
        const xmlString = await this._readFile(file, onProgress);

        // Parser avec le worker
        const result = await this._parseXmlString(xmlString, fileName);

        // Sauvegarder dans le cache
        if (useCache && result.tableMap) {
            await this.cache.saveToCache(fileName, fileSize, lastModified, xmlString, result.tableMap);
        }

        return {
            tableMap: result.tableMap,
            fromCache: false,
            fileName
        };
    }

    // Parser une chaîne XML
    async parseXmlString(xmlString, fileName = 'unnamed.xml') {
        return this._parseXmlString(xmlString, fileName);
    }

    // Parser une chaîne XML avec le worker
    async _parseXmlString(xmlString, fileName) {
        // Si pas de worker disponible, parser sur le thread principal
        if (!this.worker) {
            return {
                tableMap: this._parseXmlOnMainThread(xmlString),
                fileName
            };
        }

        // Si le worker est occupé, mettre en file d'attente
        if (this.workerBusy) {
            return new Promise((resolve, reject) => {
                const promiseId = `${fileName}_${Date.now()}`;
                this.pendingPromises.set(promiseId, { resolve, reject });
                this.workerQueue.push({ xmlString, fileName, promiseId });
            });
        }

        // Sinon, envoyer directement au worker
        this.workerBusy = true;
        return new Promise((resolve, reject) => {
            const promiseId = `${fileName}_${Date.now()}`;
            this.pendingPromises.set(promiseId, { resolve, reject });
            this.worker.postMessage({ xmlString, fileName, promiseId });
        });
    }

    // Parser sur le thread principal (fallback)
    _parseXmlOnMainThread(xmlString) {
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlString, 'text/xml');
            
            if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
                throw new Error('Erreur de syntaxe XML détectée dans le fichier.');
            }
            
            const tableMap = {};
            const root = xmlDoc.documentElement;
            if (!root) return tableMap;
            
            const children = root.children;
            
            // Utiliser une boucle for optimisée
            for (let i = 0; i < children.length; i++) {
                const node = children[i];
                const tableName = this._normalizeTagName(node.tagName);
                if (!tableName) continue;
                
                const rows = tableMap[tableName] || (tableMap[tableName] = []);
                const row = this._extractRowData(node);
                rows.push(row);
            }
            
            return tableMap;
        } catch (error) {
            throw new Error(error.message || String(error));
        }
    }

    // Normaliser un nom de balise
    _normalizeTagName(tagName) {
        if (tagName === null || tagName === undefined) return '';
        return String(tagName).replace(/^.*:/, '').trim();
    }

    // Extraire les données d'une ligne (version optimisée)
    _extractRowData(node) {
        const row = {};
        
        // Traiter les attributs
        const attrs = node.attributes || [];
        for (let i = 0; i < attrs.length; i++) {
            const attr = attrs[i];
            const name = this._normalizeTagName(attr.name);
            if (name) {
                row[name] = this._decodeXml(String(attr.value || '').trim());
            }
        }
        
        // Traiter les enfants de manière itérative
        const stack = [{ node, prefix: '' }];
        while (stack.length > 0) {
            const { node: current, prefix } = stack.pop();
            const fieldName = this._normalizeTagName(current.tagName);
            if (!fieldName) continue;
            
            if (current.children.length === 0) {
                const fullName = prefix + fieldName;
                const value = this._decodeXml((current.textContent || '').trim());
                row[fullName] = value;
            } else {
                // Ajouter les enfants à la pile (dans l'ordre inverse)
                for (let i = current.children.length - 1; i >= 0; i--) {
                    const child = current.children[i];
                    const childPrefix = prefix + fieldName + '.';
                    stack.push({ node: child, prefix: childPrefix });
                }
            }
        }
        
        return row;
    }

    // Décoder les entités XML
    _decodeXml(value) {
        return value
            .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
            .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(parseInt(code, 16)))
            .replace(/&#([0-9]+);/g, (_, code) => String.fromCodePoint(parseInt(code, 10)))
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
            .replace(/&amp;/g, '&');
    }

    // Lire un fichier
    _readFile(file, onProgress) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onprogress = progressEvent => {
                if (progressEvent.lengthComputable && onProgress) {
                    const progress = (progressEvent.loaded / progressEvent.total) * 50;
                    onProgress(progress);
                }
            };
            
            reader.onload = () => {
                if (onProgress) onProgress(50);
                resolve(reader.result);
            };
            
            reader.onerror = () => {
                reject(reader.error || new Error('Impossible de lire le fichier.'));
            };
            
            reader.onabort = () => {
                reject(new Error('Lecture du fichier annulée.'));
            };
            
            reader.readAsText(file, 'UTF-8');
        });
    }

    // Parser plusieurs fichiers en parallèle (avec limite de workers)
    async parseMultipleFiles(files, options = {}) {
        const {
            maxConcurrent = navigator.hardwareConcurrency || 4,
            useCache = true,
            onProgress = null
        } = options;

        const results = [];
        const errors = [];
        const activeParsers = [];

        // Créer plusieurs workers pour le traitement parallèle
        const workerPool = [];
        for (let i = 0; i < Math.min(maxConcurrent, 4); i++) {
            workerPool.push(this._createWorkerInstance(i));
        }

        async function parseFile(file, workerIndex) {
            const fileName = file.name;
            const fileSize = file.size;
            const lastModified = file.lastModified;

            try {
                // Vérifier le cache
                if (useCache) {
                    const cachedTableMap = await xmlCache.loadFromCache(fileName, fileSize, lastModified);
                    if (cachedTableMap) {
                        return {
                            tableMap: cachedTableMap,
                            fromCache: true,
                            fileName
                        };
                    }
                }

                // Lire le fichier
                const xmlString = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result);
                    reader.onerror = reject;
                    reader.readAsText(file, 'UTF-8');
                });

                // Parser avec le worker
                const worker = workerPool[workerIndex];
                const result = await new Promise((resolve, reject) => {
                    const promiseId = `${fileName}_${Date.now()}_${workerIndex}`;
                    worker.pendingPromises.set(promiseId, { resolve, reject });
                    worker.instance.postMessage({ xmlString, fileName, promiseId });
                });

                // Sauvegarder dans le cache
                if (useCache && result.tableMap) {
                    await xmlCache.saveToCache(fileName, fileSize, lastModified, xmlString, result.tableMap);
                }

                return {
                    tableMap: result.tableMap,
                    fromCache: false,
                    fileName
                };
            } catch (error) {
                return {
                    error: error.message || String(error),
                    fileName
                };
            }
        }

        // Traiter les fichiers par lots
        for (let i = 0; i < files.length; i += maxConcurrent) {
            const batch = files.slice(i, i + maxConcurrent);
            const batchPromises = batch.map((file, batchIndex) => 
                parseFile(file, batchIndex % workerPool.length)
            );
            
            const batchResults = await Promise.all(batchPromises);
            
            for (const result of batchResults) {
                if (result.error) {
                    errors.push(result);
                } else {
                    results.push(result);
                }
            }
            
            if (onProgress) {
                onProgress({
                    processed: i + batch.length,
                    total: files.length,
                    currentFile: batch.length > 0 ? batch[0].name : null
                });
            }
        }

        // Nettoyer les workers
        workerPool.forEach(worker => {
            worker.instance.terminate();
        });

        return { results, errors };
    }

    // Créer une instance de worker
    _createWorkerInstance(index) {
        const worker = new Worker('js/xml-parser-worker.js');
        const pendingPromises = new Map();
        
        worker.onmessage = event => {
            const { tableMap, error, fileName, promiseId } = event.data;
            const promise = pendingPromises.get(promiseId);
            if (promise) {
                pendingPromises.delete(promiseId);
                if (error) {
                    promise.reject(new Error(error));
                } else {
                    promise.resolve({ tableMap, fileName });
                }
            }
        };
        
        worker.onerror = event => {
            console.error(`Erreur du worker ${index}:`, event.message);
        };
        
        return {
            instance: worker,
            pendingPromises
        };
    }

    // Fermer le gestionnaire
    destroy() {
        this._cleanupWorker();
        this.pendingPromises.clear();
        this.workerQueue = [];
    }
}

// Instance globale du gestionnaire
const xmlParserManager = new XmlParserManager();
xmlParserManager.init();
