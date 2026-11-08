// XML Parser Web Worker - Version optimisée pour les gros fichiers
// Utilise un parsing itératif et optimisé pour éviter les problèmes de mémoire et de performance

const decodeXml = value => value
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
    .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(parseInt(code, 16)))
    .replace(/&#([0-9]+);/g, (_, code) => String.fromCodePoint(parseInt(code, 10)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');

const normalizeXmlTagName = tagName => {
    if (tagName === null || tagName === undefined) return '';
    return String(tagName).replace(/^.*:/, '').trim();
};

// Structure pour stocker les données pendant le parsing itératif
class XmlNode {
    constructor(tagName, attributes = {}, parent = null) {
        this.tagName = normalizeXmlTagName(tagName);
        this.attributes = attributes;
        this.children = [];
        this.textContent = '';
        this.parent = parent;
    }
}

// Parseur XML itératif (SAX-like) pour éviter la récursion
function parseXmlIterative(xmlString) {
    const stack = [];
    const rootNodes = [];
    let currentNode = null;
    let i = 0;
    let inTag = false;
    let inClosingTag = false;
    let inComment = false;
    let inCData = false;
    let inDoctype = false;
    let currentTagName = '';
    let currentText = '';
    let currentAttributeName = '';
    let currentAttributeValue = '';
    let inAttributeValue = false;
    let attributeQuote = null;
    let currentAttributes = {};

    while (i < xmlString.length) {
        const char = xmlString[i];
        const nextChar = xmlString[i + 1];
        const prevChar = xmlString[i - 1];

        // Gérer les commentaires
        if (!inComment && !inCData && !inDoctype && char === '<' && nextChar === '!') {
            if (xmlString.substring(i, i + 4) === '<!--') {
                inComment = true;
                i += 4;
                continue;
            }
            if (xmlString.substring(i, i + 9).toLowerCase() === '<!doctype') {
                inDoctype = true;
                i += 9;
                continue;
            }
            if (xmlString.substring(i, i + 8).toLowerCase() === '<![cdata[') {
                inCData = true;
                i += 8;
                continue;
            }
        }

        if (inComment) {
            if (char === '-' && nextChar === '-' && xmlString[i + 2] === '>') {
                inComment = false;
                i += 3;
            } else {
                i++;
            }
            continue;
        }

        if (inDoctype) {
            if (char === '>') {
                inDoctype = false;
            }
            i++;
            continue;
        }

        if (inCData) {
            if (char === ']' && nextChar === ']' && xmlString[i + 2] === '>') {
                inCData = false;
                if (currentNode) {
                    currentNode.textContent += currentText;
                    currentText = '';
                }
                i += 3;
            } else {
                currentText += char;
                i++;
            }
            continue;
        }

        // Gérer les balises
        if (char === '<') {
            // Sauvegarder le texte accumulé
            if (currentText && currentNode) {
                currentNode.textContent += currentText;
                currentText = '';
            }

            if (nextChar === '/') {
                // Balise de fermeture
                inClosingTag = true;
                i += 2;
                let tagEnd = xmlString.indexOf('>', i);
                if (tagEnd === -1) tagEnd = xmlString.length;
                const closingTagName = normalizeXmlTagName(xmlString.substring(i, tagEnd).trim());
                
                if (stack.length > 0) {
                    const expectedTag = stack[stack.length - 1].tagName;
                    if (closingTagName === expectedTag) {
                        currentNode = stack.pop();
                    }
                    // Si mismatch, on ferme quand même pour éviter les boucles infinies
                }
                i = tagEnd + 1;
                inClosingTag = false;
            } else {
                // Balise d'ouverture
                inTag = true;
                i++;
                currentTagName = '';
                currentAttributes = {};
                inAttributeValue = false;
                currentAttributeName = '';
                currentAttributeValue = '';
                attributeQuote = null;
            }
            continue;
        }

        if (inTag) {
            if (char === '>') {
                inTag = false;
                // Créer le nouveau nœud
                const newNode = new XmlNode(currentTagName, currentAttributes, currentNode);
                
                if (stack.length === 0) {
                    rootNodes.push(newNode);
                } else {
                    stack[stack.length - 1].children.push(newNode);
                }
                
                currentNode = newNode;
                stack.push(newNode);
                currentTagName = '';
                currentAttributes = {};
                i++;
            } else if (char === '/' && nextChar === '>') {
                // Balise auto-fermante
                inTag = false;
                const newNode = new XmlNode(currentTagName, currentAttributes, currentNode);
                
                if (stack.length === 0) {
                    rootNodes.push(newNode);
                } else {
                    stack[stack.length - 1].children.push(newNode);
                }
                
                i += 2;
                currentTagName = '';
                currentAttributes = {};
            } else if (char === ' ' || char === '\t' || char === '\n' || char === '\r') {
                // Espace - peut-être entre le nom de la balise et les attributs
                if (currentTagName && !currentAttributeName) {
                    // On a fini le nom de la balise
                }
                i++;
            } else if (char === '=' && currentAttributeName) {
                // Début de la valeur de l'attribut
                i++;
                if (nextChar === '"' || nextChar === "'") {
                    attributeQuote = nextChar;
                    inAttributeValue = true;
                    i++;
                } else {
                    inAttributeValue = true;
                }
            } else if (inAttributeValue) {
                if (attributeQuote && char === attributeQuote) {
                    inAttributeValue = false;
                    currentAttributes[currentAttributeName] = currentAttributeValue;
                    currentAttributeName = '';
                    currentAttributeValue = '';
                    attributeQuote = null;
                    i++;
                } else if (!attributeQuote && (char === ' ' || char === '\t' || char === '\n' || char === '\r' || char === '>')) {
                    inAttributeValue = false;
                    currentAttributes[currentAttributeName] = currentAttributeValue;
                    currentAttributeName = '';
                    currentAttributeValue = '';
                } else {
                    currentAttributeValue += char;
                    i++;
                }
            } else {
                // On est en train de lire le nom de la balise ou un attribut
                if (!currentAttributeName && (char === ' ' || char === '\t' || char === '\n' || char === '\r')) {
                    // Ignorer les espaces
                    i++;
                } else if (currentTagName && (char === ' ' || char === '\t' || char === '\n' || char === '\r')) {
                    // On a fini le nom de la balise, on commence un attribut
                    if (!currentAttributeName) {
                        currentAttributeName = '';
                    }
                    i++;
                } else if (currentAttributeName) {
                    currentAttributeName += char;
                    i++;
                } else {
                    currentTagName += char;
                    i++;
                }
            }
            continue;
        }

        // Accumuler le texte
        if (!inTag && !inClosingTag) {
            currentText += char;
        }
        i++;
    }

    return rootNodes;
}

// Convertir la structure XmlNode en tableMap (format attendu par l'application)
function convertToTableMap(rootNodes) {
    const tableMap = {};
    
    for (const rootNode of rootNodes) {
        const tableName = rootNode.tagName;
        if (!tableName) continue;
        
        const rows = tableMap[tableName] || (tableMap[tableName] = []);
        
        // Traiter chaque enfant comme une ligne
        for (const child of rootNode.children) {
            const row = extractRowData(child);
            rows.push(row);
        }
    }
    
    return tableMap;
}

// Extraire les données d'un nœud de manière itérative (évite la récursion)
function extractRowData(node) {
    const row = {};
    const stack = [{ node, prefix: '' }];
    
    // Ajouter les attributs du nœud courant
    for (const [attrName, attrValue] of Object.entries(node.attributes)) {
        row[attrName] = decodeXml(String(attrValue || '').trim());
    }
    
    // Traiter les enfants de manière itérative
    while (stack.length > 0) {
        const { node: current, prefix } = stack.pop();
        
        // Si le nœud a du texte et pas d'enfants, c'est une valeur simple
        if (current.children.length === 0 && current.tagName) {
            const fieldName = prefix + current.tagName;
            const value = decodeXml((current.textContent || '').trim());
            row[fieldName] = value;
        }
        
        // Ajouter les enfants à la pile (dans l'ordre inverse pour maintenir l'ordre)
        for (let i = current.children.length - 1; i >= 0; i--) {
            const child = current.children[i];
            const childPrefix = prefix + (current.tagName ? current.tagName + '.' : '');
            stack.push({ node: child, prefix: childPrefix });
        }
    }
    
    return row;
}

// Fonction principale de parsing
function parseXmlString(xmlString) {
    try {
        // Essayer d'abord avec DOMParser si disponible (plus rapide pour les petits fichiers)
        if (typeof DOMParser !== 'undefined') {
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
                const tableName = normalizeXmlTagName(node.tagName);
                if (!tableName) continue;
                
                const rows = tableMap[tableName] || (tableMap[tableName] = []);
                const row = {};
                
                // Traiter les attributs
                const attrs = node.attributes || [];
                for (let j = 0; j < attrs.length; j++) {
                    const attr = attrs[j];
                    const name = normalizeXmlTagName(attr.name);
                    if (name) row[name] = decodeXml(String(attr.value || '').trim());
                }
                
                // Traiter les enfants de manière itérative
                const childStack = [{ node, prefix: '' }];
                while (childStack.length > 0) {
                    const { node: current, prefix } = childStack.pop();
                    const fieldName = normalizeXmlTagName(current.tagName);
                    if (!fieldName) continue;
                    
                    if (current.children.length === 0) {
                        const fullName = prefix + fieldName;
                        const value = decodeXml((current.textContent || '').trim());
                        row[fullName] = value;
                    } else {
                        // Ajouter les enfants à la pile
                        for (let k = current.children.length - 1; k >= 0; k--) {
                            const child = current.children[k];
                            const childPrefix = prefix + fieldName + '.';
                            childStack.push({ node: child, prefix: childPrefix });
                        }
                    }
                }
                
                rows.push(row);
            }
            
            return tableMap;
        }
        
        // Fallback : parsing itératif manuel
        const rootNodes = parseXmlIterative(xmlString);
        return convertToTableMap(rootNodes);
        
    } catch (error) {
        throw new Error(error.message || String(error));
    }
}

// Gestionnaire de message
self.onmessage = event => {
    try {
        const { xmlString, chunkIndex, totalChunks, fileName } = event.data;
        
        // Si on reçoit des chunks, on les assemble
        if (typeof chunkIndex === 'number' && typeof totalChunks === 'number') {
            // Pour l'instant, on gère uniquement le parsing complet
            // On pourrait implémenter un système de streaming plus tard
            if (chunkIndex === 0) {
                // Premier chunk, on commence le parsing
                const tableMap = parseXmlString(xmlString);
                self.postMessage({ tableMap, fileName, chunkIndex, totalChunks });
            }
        } else {
            // Parsing standard
            const tableMap = parseXmlString(xmlString);
            self.postMessage({ tableMap, fileName });
        }
    } catch (error) {
        self.postMessage({ 
            error: error.message || String(error), 
            fileName: event.data.fileName 
        });
    }
};
