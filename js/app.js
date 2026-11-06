        function replaceChildrenCompat(element) {
            const children = Array.prototype.slice.call(arguments, 1);
            while (element.firstChild) element.removeChild(element.firstChild);
            children.forEach(child => element.appendChild(child));
        }

        function parseXmlTableMapOnMainThread(xmlString) {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlString, 'text/xml');
            if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
                throw new Error('Erreur de syntaxe XML détectée dans le fichier.');
            }
            const tableMap = {};
            const children = xmlDoc.documentElement ? xmlDoc.documentElement.children : [];
            for (let i = 0; i < children.length; i++) {
                const node = children[i];
                const tableName = normalizeXmlTagName(node.tagName);
                const rows = tableMap[tableName] || (tableMap[tableName] = []);
                rows.push(extractXmlRowData(node));
            }
            return tableMap;
        }

        // État Global de l'application
        let appState = {
            fileName: "",
            tables: {},          // { tableName: [ {col1: val1, ...}, ... ] }
            tableOrder: [],      // Ordre d'affichage des tables dans le volet de navigation
            tableSources: {},    // Fichier XML d'origine pour chaque table
            xmlSources: {},      // Couleur associée à chaque fichier XML importé
            canvasTables: [],    // [ { name: "z_com", x: 50, y: 50, width: 220, height: 200 }, ... ]
            joins: [],           // [ { id, fromTable, fromField, toTable, toField, type: "INNER"|"LEFT"|"RIGHT" }, ... ]
            queryGridRows: [],   // Lignes de la grille de critères style Access (8 lignes)
            activeJoinId: null   // ID de la liaison actuellement cliquée
        };
        const QUERY_STORAGE_KEY = 'xml2data.savedQueries';
        const QUERY_DRAFT_KEY = 'xml2data.queryDraft';
        const QUERY_SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1lTZqnON6cHFABqoQuNamg7l2ecwqi7A6QsHO1VnErWw/export?format=csv';
        let savedQueries = [];
        let sheetQueries = [];
        let sheetQueriesLoaded = false;
        let sheetQueriesLoading = false;
        let activeSheetQuery = null;
        let activeQueryId = null;
        let isNewQueryDraft = false;
        let contextMenuQueryId = null;
        let contextMenuQuerySource = 'saved';
        let pendingRenameQueryId = null;
        let pendingSaveAsNewFromSheetName = null;
        let activeSqlText = null;
        let savedQueryDefinition = '';
        let skipUnloadWarning = false;
        let pendingCloseAfterSave = false;
        let pendingNewQueryAfterSave = false;
        let lastQueryResults = [];
        let lastQueryResultColumns = [];
        let openTableTabs = [];
        let activeWorkspaceTab = 'query-design';
        let activeQueryWorker = null;
        let queryExecutionTimer = null;
        const QUERY_EXECUTION_TIMEOUT_MS = 30000;
        const QUERY_MAX_INTERMEDIATE_ROWS = 1000000;

        window.addEventListener('DOMContentLoaded', () => {
            initQueryGrid();
            loadSavedQueries();
            savedQueryDefinition = getQueryDefinition();
            restoreQueryDraft();
            updateQueryDesignTabLabel();
            updateSqlButtonState();
            updateActionButtonsState();
            const selectAllCheckbox = document.getElementById('selectAllTablesCheckbox');
            if (selectAllCheckbox) {
                selectAllCheckbox.addEventListener('change', toggleSelectAllTables);
            }
            const addXmlButton = document.getElementById('addXmlButton');
            const xmlFileInput = document.getElementById('xmlFileInput');
            if (addXmlButton && xmlFileInput) {
                addXmlButton.addEventListener('click', () => xmlFileInput.click());
            }
            document.addEventListener('click', event => {
                if (!event.target.closest('#savedQueryContextMenu')) closeSavedQueryContextMenu();
            });
            document.addEventListener('keydown', event => {
                if (event.key === 'Escape') closeSavedQueryContextMenu();
            });
            document.addEventListener('input', updateSaveButtonState);
            document.addEventListener('change', updateSaveButtonState);
            document.addEventListener('click', updateSaveButtonState);
            document.addEventListener('dblclick', updateSaveButtonState);
            document.addEventListener('mouseup', updateSaveButtonState);
            document.addEventListener('dragend', updateSaveButtonState);
            updateSaveButtonState();
        });

        function updateActionButtonsState() {
            const executeButton = document.getElementById('executeQueryButton');
            const exportButton = document.getElementById('exportRibbonButton');
            const hasLoadedTables = Object.keys(appState.tables).length > 0;
            const hasQueryResultToExport = lastQueryResults.length > 0 || lastQueryResultColumns.length > 0;
            const hasTableTabToExport = activeWorkspaceTab !== 'query-design' && !!appState.tables[activeWorkspaceTab];

            if (executeButton) {
                executeButton.disabled = !hasLoadedTables;
                executeButton.title = hasLoadedTables ? 'Exécuter la requête' : 'Importez une table pour exécuter la requête';
            }

            if (exportButton) {
                const canExport = hasLoadedTables || hasQueryResultToExport || hasTableTabToExport;
                exportButton.disabled = !canExport;
                exportButton.title = canExport ? 'Exporter vers Excel' : 'Aucune table ni résultat à exporter';
            }
        }

        // ===== Gestion des requetes enregistrees et de l'etat brouillon =====
        function getQueryDefinition() {
            return JSON.stringify({
                canvasTables: appState.canvasTables,
                joins: appState.joins,
                queryGridRows: appState.queryGridRows,
                sqlText: activeSqlText
            });
        }

        function updateSaveButtonState() {
            const button = document.getElementById('saveQueryButton');
            if (!button) return;
            const hasChanges = getQueryDefinition() !== savedQueryDefinition;
            button.disabled = !hasChanges;
            button.title = hasChanges ? 'Enregistrer les modifications de la requête' : 'Aucune modification à enregistrer';
            const sqlSaveButton = document.getElementById('saveSqlButton');
            if (sqlSaveButton) sqlSaveButton.disabled = !hasChanges;
        }

        function updateSqlButtonState() {
            const button = document.getElementById('showSqlButton');
            if (!button) return;
            const hasSelectedQuery = isNewQueryDraft || savedQueries.some(query => query.id === activeQueryId) || Boolean(activeSheetQuery);
            button.disabled = !hasSelectedQuery;
            button.title = hasSelectedQuery ? 'Afficher et modifier la requête en SQL' : 'Sélectionnez une requête enregistrée';
        }

        function loadSavedQueries() {
            try {
                const storedQueries = JSON.parse(localStorage.getItem(QUERY_STORAGE_KEY) || '[]');
                savedQueries = Array.isArray(storedQueries) ? storedQueries : [];
            } catch (error) {
                savedQueries = [];
            }
            renderSavedQueries();
        }

        function persistSavedQueries() {
            try {
                localStorage.setItem(QUERY_STORAGE_KEY, JSON.stringify(savedQueries));
                return true;
            } catch (error) {
                alert("Impossible d'enregistrer les requêtes dans le stockage du navigateur.");
                return false;
            }
        }

        function renderSavedQueries() {
            const list = document.getElementById('savedQueryList');
            const sheetList = document.getElementById('sheetQueryList');
            if (!list || !sheetList) return;
            updateNavigationTabCounts();
            list.innerHTML = '';
            sheetList.innerHTML = '';
            const groupedQueries = new Map();
            sheetQueries.forEach((query, index) => {
                const groupName = query.group || 'Sans groupe';
                if (!groupedQueries.has(groupName)) groupedQueries.set(groupName, []);
                groupedQueries.get(groupName).push({ query, index });
            });
            groupedQueries.forEach((queries, groupName) => {
                const group = document.createElement('section');
                group.className = 'mb-1 border-b border-[#e1dfdd]';
                const groupButton = document.createElement('button');
                groupButton.type = 'button';
                groupButton.className = 'flex w-full items-center justify-between border-b border-[#c8c6c4] bg-[#e1dfdd] px-2 py-1.5 text-left font-semibold text-[#323130] hover:bg-[#d6d3d1]';
                groupButton.setAttribute('aria-expanded', 'false');
                const groupLabel = document.createElement('span');
                groupLabel.innerHTML = `<i class="fa-solid fa-chevron-right mr-2 text-[10px] text-gray-500"></i>${groupName}`;
                const groupCount = document.createElement('span');
                groupCount.className = 'text-[10px] font-normal text-gray-500';
                groupCount.textContent = String(queries.length);
                groupButton.append(groupLabel, groupCount);
                const groupList = document.createElement('div');
                groupList.className = 'hidden pb-1';
                groupButton.addEventListener('click', () => {
                    const expanded = groupButton.getAttribute('aria-expanded') === 'true';
                    groupButton.setAttribute('aria-expanded', String(!expanded));
                    groupList.classList.toggle('hidden', expanded);
                    groupLabel.querySelector('i').className = `fa-solid ${expanded ? 'fa-chevron-right' : 'fa-chevron-down'} mr-2 text-[10px] text-gray-500`;
                });
                queries.forEach(({ query, index }) => {
                    const row = document.createElement('div');
                    row.className = `flex items-center justify-between gap-1 rounded p-1.5 pl-5 group ${activeSheetQuery && activeSheetQuery.name === query.name ? 'bg-red-50 text-red-800' : 'text-gray-800 hover:bg-gray-100'}`;
                    row.title = query.name;

                    const nameButton = document.createElement('button');
                    nameButton.type = 'button';
                    nameButton.className = 'flex min-w-0 items-center gap-2 text-left';
                    nameButton.title = `Charger ${query.name}`;
                    nameButton.innerHTML = '<i class="fa-solid fa-table-list text-blue-700"></i>';
                    const name = document.createElement('span');
                    name.className = 'truncate';
                    name.textContent = query.name;
                    nameButton.appendChild(name);
                    nameButton.addEventListener('click', () => loadSheetQuery(index));
                    nameButton.addEventListener('contextmenu', event => {
                        event.preventDefault();
                        event.stopPropagation();
                        openSavedQueryContextMenu(index, 'shared', event.clientX, event.clientY);
                    });
                    row.appendChild(nameButton);
                    groupList.appendChild(row);
                });
                group.append(groupButton, groupList);
                sheetList.appendChild(group);
            });

            if (sheetQueries.length === 0) {
                const emptySheetList = document.createElement('div');
                emptySheetList.className = 'text-gray-400 text-center py-2 italic';
                emptySheetList.textContent = sheetQueriesLoaded ? 'Aucune requête dans la feuille.' : 'Aucune requête partagée chargée.';
                sheetList.appendChild(emptySheetList);
            }

            if (savedQueries.length === 0) {
                list.innerHTML = '<div class="text-gray-400 text-center py-3 italic">Aucune requête enregistrée.</div>';
            } else {
                savedQueries.forEach(query => {
                    const row = document.createElement('div');
                    row.className = `flex items-center justify-between p-1.5 rounded group ${query.id === activeQueryId ? 'bg-red-50 text-red-800' : 'hover:bg-gray-100 text-gray-800'}`;
                    const openButton = document.createElement('button');
                    openButton.type = 'button';
                    openButton.draggable = true;
                    openButton.className = 'flex min-w-0 items-center gap-2 text-left';
                    openButton.title = `Ouvrir la requête ${query.name}`;
                    openButton.innerHTML = '<i class="fa-solid fa-table-list text-red-700"></i>';
                    const name = document.createElement('span');
                    name.className = 'truncate';
                    name.textContent = query.name;
                    openButton.appendChild(name);
                    openButton.addEventListener('click', () => loadSavedQuery(query.id));
                    openButton.addEventListener('contextmenu', event => {
                        event.preventDefault();
                        event.stopPropagation();
                        openSavedQueryContextMenu(query.id, 'saved', event.clientX, event.clientY);
                    });
                    openButton.addEventListener('dblclick', event => {
                        event.stopPropagation();
                        loadSavedQuery(query.id);
                        showQueryDesignTab();
                    });
                    openButton.addEventListener('dragstart', event => {
                        event.dataTransfer.setData('application/x-query-id', query.id);
                        event.dataTransfer.effectAllowed = 'copy';
                    });

                    const deleteButton = document.createElement('button');
                    deleteButton.type = 'button';
                    deleteButton.className = 'px-1 text-gray-500 hover:text-red-700';
                    deleteButton.title = `Supprimer la requête ${query.name}`;
                    deleteButton.setAttribute('aria-label', `Supprimer ${query.name}`);
                    deleteButton.innerHTML = '<i class="fa-solid fa-trash-can"></i>';
                    deleteButton.addEventListener('click', () => deleteSavedQuery(query.id));
                    row.append(openButton, deleteButton);
                    list.appendChild(row);
                });
            }
        }

        async function loadSheetQueries() {
            const status = document.getElementById('sheetQueryStatus');
            const refreshButton = document.querySelector('#navQueriesPanel button[onclick="loadSheetQueries()"]');
            if (sheetQueriesLoading) return;
            sheetQueriesLoading = true;
            if (refreshButton) refreshButton.disabled = true;
            status.classList.remove('text-red-700');
            status.textContent = 'Chargement des requêtes depuis Google Sheets...';
            try {
                const response = await fetch(QUERY_SHEET_CSV_URL, { cache: 'no-store' });
                if (!response.ok) throw new Error(`Réponse HTTP ${response.status}`);
                const csvText = await response.text();
                const records = parseCsvRecords(csvText);
                if (records.length === 0) throw new Error('La feuille ne contient aucune ligne.');

                const headers = records[0].map(header => header.trim().toLowerCase());
                const nameIndex = headers.indexOf('nom_requete');
                const groupIndex = headers.indexOf('groupe');
                const descriptionIndex = headers.indexOf('description');
                const sqlIndex = headers.indexOf('requete');
                const labelIndex = groupIndex >= 0 ? groupIndex : descriptionIndex;
                if ([nameIndex, labelIndex, sqlIndex].some(index => index < 0)) {
                    throw new Error('Colonnes attendues introuvables : nom_requete, groupe, requete.');
                }

                sheetQueries = records.slice(1)
                    .filter(record => record.some(value => value.trim()))
                    .map(record => ({
                        name: (record[nameIndex] || '').trim(),
                        group: (record[labelIndex] || '').trim(),
                        sql: (record[sqlIndex] || '').trim()
                    }))
                    .filter(query => query.name && query.sql);
                sheetQueriesLoaded = true;
                status.textContent = `${sheetQueries.length} requête(s) chargée(s). Survolez un nom pour afficher sa description.`;
                renderSavedQueries();
            } catch (error) {
                console.error('Échec du chargement des requêtes Google Sheets :', error);
                status.classList.add('text-red-700');
                const isLocalFile = window.location.protocol === 'file:';
                status.textContent = isLocalFile
                    ? 'Impossible de charger les requêtes partagées depuis un fichier local. Ouvrez l’application via un serveur HTTP/HTTPS, puis actualisez.'
                    : `Impossible de charger les requêtes partagées : ${error.message}`;
            } finally {
                sheetQueriesLoading = false;
                if (refreshButton) refreshButton.disabled = false;
            }
        }

        function parseCsvRecords(text) {
            const records = [];
            let record = [];
            let value = '';
            let quoted = false;
            for (let index = 0; index < text.length; index++) {
                const char = text[index];
                if (quoted) {
                    if (char === '"' && text[index + 1] === '"') {
                        value += '"';
                        index++;
                    } else if (char === '"') {
                        quoted = false;
                    } else {
                        value += char;
                    }
                } else if (char === '"' && value.length === 0) {
                    quoted = true;
                } else if (char === ',') {
                    record.push(value);
                    value = '';
                } else if (char === '\n' || char === '\r') {
                    if (char === '\r' && text[index + 1] === '\n') index++;
                    record.push(value);
                    records.push(record);
                    record = [];
                    value = '';
                } else {
                    value += char;
                }
            }
            if (quoted) throw new Error('Le CSV contient un champ entre guillemets non terminé.');
            if (value.length > 0 || record.length > 0) {
                record.push(value);
                records.push(record);
            }
            return records;
        }

        function splitSqlTopLevel(text, delimiter) {
            const parts = [];
            let start = 0;
            let depth = 0;
            let inBracket = false;
            let quote = '';
            for (let index = 0; index < text.length; index++) {
                const char = text[index];
                if (inBracket) {
                    if (char === ']' && text[index + 1] === ']') index++;
                    else if (char === ']') inBracket = false;
                    continue;
                }
                if (quote) {
                    if (char === quote && text[index + 1] === quote) index++;
                    else if (char === quote) quote = '';
                    continue;
                }
                if (char === '[') inBracket = true;
                else if (char === '"' || char === "'") quote = char;
                else if (char === '(') depth++;
                else if (char === ')') depth--;
                else if (depth === 0 && text.slice(index, index + delimiter.length).toUpperCase() === delimiter.toUpperCase()) {
                    const before = index === 0 ? ' ' : text[index - 1];
                    const after = text[index + delimiter.length] || ' ';
                    const isWord = /^[A-Z]+$/i.test(delimiter);
                    if (!isWord || (!/[\w]/.test(before) && !/[\w]/.test(after))) {
                        parts.push(text.slice(start, index).trim());
                        index += delimiter.length - 1;
                        start = index + 1;
                    }
                }
            }
            parts.push(text.slice(start).trim());
            return parts;
        }

        function findSqlKeywordIndex(text, keyword, start = 0) {
            let depth = 0;
            let inBracket = false;
            let quote = '';
            const pattern = new RegExp(`^${keyword.replace(/\s+/g, '\\s+')}\\b`, 'i');
            for (let index = start; index < text.length; index++) {
                const char = text[index];
                if (inBracket) {
                    if (char === ']' && text[index + 1] === ']') index++;
                    else if (char === ']') inBracket = false;
                    continue;
                }
                if (quote) {
                    if (char === quote && text[index + 1] === quote) index++;
                    else if (char === quote) quote = '';
                    continue;
                }
                if (char === '[') inBracket = true;
                else if (char === '"' || char === "'") quote = char;
                else if (char === '(') depth++;
                else if (char === ')') depth--;
                else if (depth === 0 && (index === 0 || !/[\w]/.test(text[index - 1]))) {
                    if (pattern.test(text.slice(index))) return index;
                }
            }
            return -1;
        }

        function readSqlIdentifier(text, start = 0) {
            let index = start;
            while (/\s/.test(text[index] || '')) index++;
            if (text[index] === '[') {
                let value = '';
                index++;
                while (index < text.length) {
                    if (text[index] === ']' && text[index + 1] === ']') {
                        value += ']';
                        index += 2;
                    } else if (text[index] === ']') {
                        return { value, end: index + 1 };
                    } else {
                        value += text[index++];
                    }
                }
                throw new Error('Un identifiant entre crochets n’est pas terminé.');
            }
            const match = text.slice(index).match(/^[A-Za-z_][\w$]*/);
            if (!match) throw new Error('Identifiant SQL attendu.');
            return { value: match[0], end: index + match[0].length };
        }

        function readSqlReference(text) {
            let position = 0;
            const first = readSqlIdentifier(text, position);
            position = first.end;
            while (/\s/.test(text[position] || '')) position++;
            if (text[position] === '.') {
                const second = readSqlIdentifier(text, position + 1);
                position = second.end;
                return { table: first.value, field: second.value, end: position };
            }
            return { table: '', field: first.value, end: position };
        }

        function stripOuterSqlParentheses(text) {
            let value = text.trim();
            while (value.startsWith('(') && value.endsWith(')')) {
                let depth = 0;
                let closesAtEnd = false;
                let inBracket = false;
                let quote = '';
                for (let index = 0; index < value.length; index++) {
                    const char = value[index];
                    if (inBracket) {
                        if (char === ']' && value[index + 1] === ']') index++;
                        else if (char === ']') inBracket = false;
                        continue;
                    }
                    if (quote) {
                        if (char === quote && value[index + 1] === quote) index++;
                        else if (char === quote) quote = '';
                        continue;
                    }
                    if (char === '[') inBracket = true;
                    else if (char === '"' || char === "'") quote = char;
                    else if (char === '(') depth++;
                    else if (char === ')' && --depth === 0) {
                        closesAtEnd = index === value.length - 1;
                        break;
                    }
                }
                if (!closesAtEnd) break;
                value = value.slice(1, -1).trim();
            }
            return value;
        }

        function removeFromGroupingParentheses(text) {
            let result = '';
            let inBracket = false;
            let quote = '';
            for (let index = 0; index < text.length; index++) {
                const character = text[index];
                if (inBracket) {
                    result += character;
                    if (character === ']' && text[index + 1] === ']') result += text[++index];
                    else if (character === ']') inBracket = false;
                } else if (quote) {
                    result += character;
                    if (character === quote && text[index + 1] === quote) result += text[++index];
                    else if (character === quote) quote = '';
                } else if (character === '[') {
                    inBracket = true;
                    result += character;
                } else if (character === '"' || character === "'") {
                    quote = character;
                    result += character;
                } else if (character !== '(' && character !== ')') {
                    result += character;
                }
            }
            return result;
        }

        function parseSheetQuerySql(sqlText) {
            const sql = String(sqlText || '').trim().replace(/;\s*$/, '');
            if (!/^SELECT\b/i.test(sql)) throw new Error('Seules les requêtes SELECT sont prises en charge.');

            const fromIndex = findSqlKeywordIndex(sql, 'FROM', 6);
            if (fromIndex < 0) throw new Error('Clause FROM introuvable.');
            const selectText = sql.slice(6, fromIndex).trim().replace(/^DISTINCT\s+/i, '');
            const whereIndex = findSqlKeywordIndex(sql, 'WHERE', fromIndex + 4);
            const orderIndex = findSqlKeywordIndex(sql, 'ORDER BY', fromIndex + 4);
            const endFrom = [whereIndex, orderIndex].filter(index => index >= 0).reduce((min, index) => Math.min(min, index), sql.length);
            // Access encloses chained joins in parentheses. The designer only
            // supports table joins, so grouping parentheses can be ignored here.
            const fromText = removeFromGroupingParentheses(sql.slice(fromIndex + 4, endFrom).trim());
            const whereText = whereIndex < 0 ? '' : sql.slice(whereIndex + 5, orderIndex >= 0 ? orderIndex : sql.length).trim();
            const orderText = orderIndex < 0 ? '' : sql.slice(orderIndex + 8).trim();

            const aliases = new Map();
            const tables = [];
            const joins = [];
            let position = 0;
            const readTable = () => {
                const identifier = readSqlIdentifier(fromText, position);
                position = identifier.end;
                const tableName = identifier.value;
                while (/\s/.test(fromText[position] || '')) position++;
                let alias = tableName;
                const asMatch = fromText.slice(position).match(/^AS\s+/i);
                if (asMatch) {
                    position += asMatch[0].length;
                    const aliasIdentifier = readSqlIdentifier(fromText, position);
                    alias = aliasIdentifier.value;
                    position = aliasIdentifier.end;
                } else {
                    const aliasMatch = fromText.slice(position).match(/^([A-Za-z_][\w$]*)/);
                    if (aliasMatch && !/^(?:LEFT|RIGHT|INNER|FULL|CROSS|JOIN|ON|WHERE|ORDER)$/i.test(aliasMatch[1])) {
                        alias = aliasMatch[1];
                        position += aliasMatch[1].length;
                    }
                }
                aliases.set(alias.toLowerCase(), tableName);
                if (!tables.includes(tableName)) tables.push(tableName);
                return tableName;
            };

            const primaryTable = readTable();
            while (position < fromText.length) {
                while (/\s/.test(fromText[position] || '')) position++;
                if (position >= fromText.length) break;
                const joinMatch = fromText.slice(position).match(/^(?:(LEFT|RIGHT|INNER|FULL|CROSS)(?:\s+OUTER)?\s+)?JOIN\b/i);
                if (!joinMatch) throw new Error(`Syntaxe FROM/JOIN non prise en charge près de « ${fromText.slice(position, position + 30)} ».`);
                const joinType = (joinMatch[1] || 'INNER').toUpperCase();
                if (joinType === 'RIGHT' || joinType === 'FULL' || joinType === 'CROSS') {
                    throw new Error(`${joinType} JOIN n’est pas pris en charge par le moteur de conception.`);
                }
                position += joinMatch[0].length;
                const targetTable = readTable();
                while (/\s/.test(fromText[position] || '')) position++;
                const onMatch = fromText.slice(position).match(/^ON\b/i);
                if (!onMatch) throw new Error('Clause ON manquante après JOIN.');
                position += onMatch[0].length;
                const nextJoinKeyword = findSqlKeywordIndex(fromText, 'JOIN', position);
                let nextJoin = nextJoinKeyword;
                if (nextJoinKeyword >= 0) {
                    const joinPrefix = fromText.slice(position, nextJoinKeyword).match(/\b(?:LEFT|RIGHT|INNER|FULL|CROSS)(?:\s+OUTER)?\s*$/i);
                    if (joinPrefix) nextJoin -= joinPrefix[0].length;
                }
                const onText = fromText.slice(position, nextJoin < 0 ? fromText.length : nextJoin).trim();
                position = nextJoin < 0 ? fromText.length : nextJoin;
                const equality = onText.match(/^(.+?)\s*=\s*(.+)$/);
                if (!equality) throw new Error('La clause ON doit comparer deux champs avec « = ».');
                const left = readSqlReference(equality[1].trim());
                const right = readSqlReference(equality[2].trim());
                if (!left.table || !right.table) throw new Error('Les champs de la clause ON doivent être qualifiés par leur table.');
                const leftTable = aliases.get(left.table.toLowerCase());
                const rightTable = aliases.get(right.table.toLowerCase());
                if (leftTable === targetTable && tables.includes(rightTable)) {
                    joins.push({ id: `sheet-${joins.length}`, fromTable: rightTable, fromField: right.field, toTable: targetTable, toField: left.field, type: joinType });
                } else if (rightTable === targetTable && tables.includes(leftTable)) {
                    joins.push({ id: `sheet-${joins.length}`, fromTable: leftTable, fromField: left.field, toTable: targetTable, toField: right.field, type: joinType });
                } else {
                    throw new Error('Chaque jointure doit relier une table déjà déclarée à la nouvelle table.');
                }
            }

            const resolveReference = reference => {
                let tableName = reference.table ? aliases.get(reference.table.toLowerCase()) : '';
                if (reference.table && !tableName) throw new Error(`Table ou alias inconnu : ${reference.table}.`);
                if (!tableName) {
                    if (tables.length !== 1) throw new Error(`Le champ ${reference.field} doit être qualifié par sa table.`);
                    tableName = tables[0];
                }
                return { table: tableName, field: reference.field };
            };
            const parseProjection = item => {
                const asParts = splitSqlTopLevel(item, 'AS');
                if (asParts.length > 2) throw new Error(`Alias de champ invalide : ${item}.`);
                const source = asParts[0].trim();
                const aggregateMatch = source.match(/^(SUM|AVG|COUNT|MIN|MAX)\s*\((.*)\)$/i);
                const reference = readSqlReference(aggregateMatch ? aggregateMatch[2].trim() : source);
                if (!aggregateMatch && source.slice(reference.end).trim()) throw new Error(`Expression de champ non prise en charge : ${item}.`);
                const field = resolveReference(reference);
                let alias = '';
                if (asParts.length === 2) {
                    const aliasRef = readSqlIdentifier(asParts[1]);
                    if (asParts[1].slice(aliasRef.end).trim()) throw new Error(`Alias invalide : ${asParts[1]}.`);
                    alias = aliasRef.value;
                }
                return {
                    field: field.field,
                    table: field.table,
                    alias,
                    expression: source,
                    total: aggregateMatch ? aggregateMatch[1].toUpperCase() : '',
                    show: true,
                    sort: '',
                    criteria: '',
                    or: '',
                    width: 180
                };
            };
            if (!selectText || selectText === '*') throw new Error('Les requêtes avec SELECT * doivent être converties en sélection explicite de champs.');
            const queryRows = splitSqlTopLevel(selectText, ',').map(parseProjection);

            const parseCondition = conditionText => {
                const condition = stripOuterSqlParentheses(conditionText);
                const reference = readSqlReference(condition);
                const rest = condition.slice(reference.end).trim();
                const operatorMatch = rest.match(/^(IS\s+NOT\s+NULL|IS\s+NULL|NOT\s+LIKE|NOT\s+IN|NOT\s*=|LIKE|IN|>=|<=|<>|!=|=|>|<)\s*(.*)$/i);
                if (!operatorMatch) throw new Error(`Critère SQL non pris en charge : ${condition}.`);
                const field = resolveReference(reference);
                const operator = operatorMatch[1].replace(/\s+/g, ' ').toUpperCase();
                let operand = operatorMatch[2].trim();
                if (operator === 'IS NULL' || operator === 'IS NOT NULL') operand = operator;
                else if (operator === 'NOT =' || operator === '!=') operand = `<> ${operand}`;
                else operand = `${operator} ${operand}`;
                return { ...field, criterion: operand };
            };

            const conditions = whereText ? splitSqlTopLevel(whereText, 'AND') : [];
            conditions.forEach(condition => {
                const alternatives = splitSqlTopLevel(stripOuterSqlParentheses(condition), 'OR').map(parseCondition);
                if (alternatives.some(item => item.table !== alternatives[0].table || item.field !== alternatives[0].field)) {
                    throw new Error('Les conditions OR ne sont prises en charge que pour un même champ.');
                }
                const field = alternatives[0];
                let row = queryRows.find(item => item.table === field.table && item.field === field.field && !item.criteria && !item.or);
                if (!row) {
                    row = { ...field, alias: '', expression: `${field.table}.${field.field}`, total: '', show: false, sort: '', criteria: '', or: '', width: 180 };
                    queryRows.push(row);
                }
                row.criteria = alternatives.map(item => item.criterion).join(' OR ');
            });

            if (orderText) {
                splitSqlTopLevel(orderText, ',').forEach(orderItem => {
                    const match = orderItem.match(/^(.+?)(?:\s+(ASC|DESC))?$/i);
                    const referenceText = match[1].trim();
                    let row = queryRows.find(item => item.alias && item.alias.toLowerCase() === referenceText.replace(/^\[|\]$/g, '').toLowerCase());
                    if (!row) {
                        const reference = readSqlReference(referenceText);
                        row = queryRows.find(item => {
                            const field = resolveReference(reference);
                            return item.table === field.table && item.field === field.field;
                        });
                    }
                    if (!row) throw new Error(`Le tri doit porter sur un champ sélectionné : ${referenceText}.`);
                    row.sort = (match[2] || 'ASC').toUpperCase();
                });
            }

            while (queryRows.length < 8) queryRows.push({ field: '', table: '', alias: '', expression: '', total: '', show: false, sort: '', criteria: '', or: '', width: 180 });
            const canvasTables = tables.map((name, index) => ({
                name,
                x: 40 + index * 260,
                y: 40,
                width: 220,
                height: 220
            }));
            return { primaryTable, tables, canvasTables, joins, queryGridRows: queryRows };
        }

        function isAdvancedAccessSql(sqlText) {
            const sql = String(sqlText || '');
            return /\b(?:GROUP\s+BY|HAVING|UNION(?:\s+ALL)?|TRANSFORM|PIVOT|IIF|SWITCH|CHOOSE|DCOUNT|DSUM|DLOOKUP|SELECT\s*\(|FROM\s*\()/i.test(sql);
        }

        function applyAdvancedSqlOnly(sqlText, queryName) {
            activeSqlText = sqlText;
            isNewQueryDraft = !activeQueryId && !activeSheetQuery;
            updateSqlButtonState();
            updateSaveButtonState();
            document.getElementById('statusBarMode').textContent = queryName
                ? `Requête avancée « ${queryName} » chargée en mode SQL`
                : 'Requête avancée chargée en mode SQL';
        }

        function loadSheetQuery(index) {
            const query = sheetQueries[index];
            if (!query) return;
            if (isAdvancedAccessSql(query.sql)) {
                if (getQueryDefinition() !== savedQueryDefinition && !confirm('Les modifications non enregistrées seront perdues. Continuer ?')) return;
                activeQueryId = null;
                activeSheetQuery = query;
                activeSqlText = query.sql;
                isNewQueryDraft = false;
                updateQueryDesignTabLabel();
                updateSqlButtonState();
                updateSaveButtonState();
                document.getElementById('statusBarMode').textContent = `Requête avancée « ${query.name} » chargée en mode SQL`;
                return;
            }
            let parsed;
            try {
                parsed = parseSheetQuerySql(query.sql);
            } catch (error) {
                const status = document.getElementById('sheetQueryStatus');
                status.classList.add('text-red-700');
                status.textContent = `« ${query.name} » n’a pas été chargée : ${error.message}`;
                console.error(`Requête « ${query.name} » non prise en charge :`, error);
                return;
            }
            if (getQueryDefinition() !== savedQueryDefinition && !confirm('Les modifications non enregistrées seront perdues. Continuer ?')) return;

            activeQueryId = null;
            activeSheetQuery = query;
            isNewQueryDraft = false;
            activeSqlText = null;
            appState.canvasTables = parsed.canvasTables;
            appState.joins = parsed.joins;
            appState.queryGridRows = parsed.queryGridRows;
            renderQueryGrid();
            refreshNavigationPane();
            renderCanvasTables();
            renderSvgJoins();
            updateQueryDesignTabLabel();
            updateSqlButtonState();
            savedQueryDefinition = getQueryDefinition();
            updateSaveButtonState();
            renderSavedQueries();
            showQueryDesignTab();

            const missingTables = parsed.tables.filter(tableName => !appState.tables[tableName]);
            const status = document.getElementById('sheetQueryStatus');
            status.classList.remove('text-red-700');
            status.textContent = missingTables.length
                ? `« ${query.name} » chargée. Tables à importer avant exécution : ${missingTables.join(', ')}.`
                : `« ${query.name} » chargée dans la conception.`;
            document.getElementById('statusBarMode').textContent = `Requête « ${query.name} » chargée`;
        }

        function saveCurrentQuery() {
            let query = savedQueries.find(item => item.id === activeQueryId);
            if (!query) {
                openQueryNameModal();
                return false;
            }

            return persistQuery(query);
        }

        function openQueryNameModal() {
            if (activeSheetQuery && pendingSaveAsNewFromSheetName === null) {
                const sourceName = activeSheetQuery.name;
                if (!confirm(`La requête chargée « ${sourceName} » ne peut pas être modifiée. Enregistrer les changements comme une nouvelle requête sous un nouveau nom ?`)) {
                    pendingCloseAfterSave = false;
                    pendingNewQueryAfterSave = false;
                    return;
                }
                pendingSaveAsNewFromSheetName = sourceName;
            }
            const modal = document.getElementById('queryNameModal');
            modal.classList.remove('hidden');
            const isSheetCopy = pendingSaveAsNewFromSheetName !== null;
            document.getElementById('queryNameModalTitle').textContent = isSheetCopy
                ? 'Enregistrer comme nouvelle requête'
                : 'Enregistrer la requête';
            document.getElementById('queryNameSubmitButton').textContent = isSheetCopy ? 'Créer' : 'Enregistrer';
            pendingRenameQueryId = null;
            const input = document.getElementById('queryNameInput');
            input.value = '';
            input.placeholder = isSheetCopy ? `Nom différent de « ${pendingSaveAsNewFromSheetName} »` : '';
            input.focus();
        }

        function openSavedQueryContextMenu(queryId, source, x, y) {
            const queryExists = source === 'shared'
                ? sheetQueries[queryId] !== undefined
                : savedQueries.some(query => query.id === queryId);
            if (!queryExists) return;
            contextMenuQueryId = queryId;
            contextMenuQuerySource = source;
            const menu = document.getElementById('savedQueryContextMenu');
            menu.querySelector('[data-query-action="rename"]').disabled = source === 'shared';
            menu.querySelector('[data-query-action="delete"]').disabled = source === 'shared';
            menu.querySelectorAll('[data-query-action="rename"], [data-query-action="delete"]').forEach(button => {
                button.title = source === 'shared' ? 'Les requêtes partagées sont en lecture seule. Dupliquez-la pour la modifier.' : '';
                button.classList.toggle('cursor-not-allowed', source === 'shared');
                button.classList.toggle('opacity-50', source === 'shared');
            });
            menu.classList.remove('hidden');
            const left = Math.max(4, Math.min(x, window.innerWidth - menu.offsetWidth - 4));
            const top = Math.max(4, Math.min(y, window.innerHeight - menu.offsetHeight - 4));
            menu.style.left = `${left}px`;
            menu.style.top = `${top}px`;
            menu.querySelector('[data-query-action="duplicate"]').focus();
        }

        function handleSavedQueryContextAction(action) {
            if (contextMenuQueryId === null) return;
            const queryId = contextMenuQueryId;
            const source = contextMenuQuerySource;
            closeSavedQueryContextMenu();
            if (action === 'duplicate') {
                if (source === 'shared') duplicateSheetQuery(queryId);
                else duplicateSavedQuery(queryId);
            } else if (action === 'rename' && source === 'saved') {
                openRenameQueryDialog(queryId);
            } else if (action === 'delete' && source === 'saved') {
                deleteSavedQuery(queryId);
            }
        }

        function closeSavedQueryContextMenu() {
            const menu = document.getElementById('savedQueryContextMenu');
            if (menu) menu.classList.add('hidden');
            contextMenuQueryId = null;
            contextMenuQuerySource = 'saved';
        }

        function duplicateSheetQuery(index) {
            const source = sheetQueries[index];
            if (!source) return;
            let parsed;
            try {
                parsed = parseSheetQuerySql(source.sql);
            } catch (error) {
                document.getElementById('sheetQueryStatus').textContent = `Impossible de dupliquer « ${source.name} » : ${error.message}`;
                return;
            }
            const baseName = `${source.name} (copie)`;
            let name = baseName;
            let suffix = 2;
            while (savedQueries.some(query => query.name.toLowerCase() === name.toLowerCase())) {
                name = `${baseName} ${suffix++}`;
            }
            const now = new Date().toISOString();
            const duplicate = {
                id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
                name,
                canvasTables: parsed.canvasTables,
                joins: parsed.joins,
                queryGridRows: parsed.queryGridRows,
                sqlText: source.sql,
                createdAt: now,
                updatedAt: now
            };
            savedQueries.push(duplicate);
            if (!persistSavedQueries()) {
                savedQueries = savedQueries.filter(query => query.id !== duplicate.id);
                return;
            }
            renderSavedQueries();
            document.getElementById('statusBarMode').textContent = `Requête « ${name} » dupliquée dans Mes requêtes`;
            return duplicate;
        }

        function duplicateSavedQuery(queryId) {
            const source = savedQueries.find(query => query.id === queryId);
            if (!source) return;
            const baseName = `${source.name} (copie)`;
            let name = baseName;
            let suffix = 2;
            while (savedQueries.some(query => query.name.toLowerCase() === name.toLowerCase())) {
                name = `${baseName} ${suffix++}`;
            }

            const duplicate = JSON.parse(JSON.stringify(source));
            duplicate.id = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
            duplicate.name = name;
            duplicate.createdAt = new Date().toISOString();
            duplicate.updatedAt = duplicate.createdAt;
            savedQueries.push(duplicate);
            if (!persistSavedQueries()) {
                savedQueries = savedQueries.filter(query => query.id !== duplicate.id);
                return;
            }
            renderSavedQueries();
            document.getElementById('statusBarMode').textContent = `Requête « ${name} » dupliquée`;
        }

        function openRenameQueryDialog(queryId) {
            const query = savedQueries.find(item => item.id === queryId);
            if (!query) return;
            pendingRenameQueryId = queryId;
            document.getElementById('queryNameModalTitle').textContent = 'Renommer la requête';
            document.getElementById('queryNameSubmitButton').textContent = 'Renommer';
            document.getElementById('queryNameInput').value = query.name;
            document.getElementById('queryNameInput').placeholder = '';
            document.getElementById('queryNameModal').classList.remove('hidden');
            document.getElementById('queryNameInput').focus();
            document.getElementById('queryNameInput').select();
        }

        function closeQueryNameModal() {
            document.getElementById('queryNameModal').classList.add('hidden');
            document.getElementById('queryNameInput').placeholder = '';
            pendingRenameQueryId = null;
            pendingSaveAsNewFromSheetName = null;
            pendingCloseAfterSave = false;
            pendingNewQueryAfterSave = false;
        }

        function saveQueryFromDialog(event) {
            event.preventDefault();
            const name = document.getElementById('queryNameInput').value.trim();
            if (!name) return;
            if (pendingSaveAsNewFromSheetName !== null) {
                if (name.toLowerCase() === pendingSaveAsNewFromSheetName.toLowerCase()) {
                    alert('La nouvelle requête doit avoir un nom différent de la requête chargée.');
                    return;
                }
                if (savedQueries.some(item => item.name.toLowerCase() === name.toLowerCase())) {
                    alert(`Une requête nommée « ${name} » existe déjà. Choisissez un autre nom.`);
                    return;
                }
                const query = { id: `${Date.now()}-${Math.random().toString(36).slice(2)}`, name };
                savedQueries.push(query);
                if (!persistQuery(query)) {
                    savedQueries = savedQueries.filter(item => item.id !== query.id);
                    return;
                }
                pendingSaveAsNewFromSheetName = null;
                document.getElementById('queryNameModal').classList.add('hidden');
                document.getElementById('queryNameInput').placeholder = '';
                if (pendingCloseAfterSave) {
                    pendingCloseAfterSave = false;
                    window.close();
                } else if (pendingNewQueryAfterSave) {
                    pendingNewQueryAfterSave = false;
                    resetToNewQuery();
                }
                return;
            }
            if (pendingRenameQueryId) {
                const query = savedQueries.find(item => item.id === pendingRenameQueryId);
                if (!query) {
                    closeQueryNameModal();
                    return;
                }
                const conflict = savedQueries.find(item =>
                    item.id !== query.id && item.name.toLowerCase() === name.toLowerCase()
                );
                if (conflict) {
                    alert(`Une requête nommée « ${conflict.name} » existe déjà.`);
                    return;
                }
                const previousName = query.name;
                query.name = name;
                if (!persistSavedQueries()) {
                    query.name = previousName;
                    return;
                }
                if (activeQueryId === query.id) updateQueryDesignTabLabel();
                renderSavedQueries();
                document.getElementById('statusBarMode').textContent = `Requête renommée « ${name} »`;
                closeQueryNameModal();
                return;
            }
            let query = savedQueries.find(item => item.name.toLowerCase() === name.toLowerCase());
            if (query && !confirm(`La requête « ${query.name} » existe déjà. La remplacer ?`)) return;
            if (!query) {
                query = { id: `${Date.now()}-${Math.random().toString(36).slice(2)}`, name };
                savedQueries.push(query);
            }
            if (persistQuery(query)) {
                document.getElementById('queryNameModal').classList.add('hidden');
                if (pendingCloseAfterSave) {
                    pendingCloseAfterSave = false;
                    window.close();
                } else if (pendingNewQueryAfterSave) {
                    pendingNewQueryAfterSave = false;
                    resetToNewQuery();
                }
            }
        }

        function closeNewQueryConfirmation() {
            document.getElementById('newQueryConfirmModal').classList.add('hidden');
        }

        function resolveNewQueryConfirmation(decision) {
            closeNewQueryConfirmation();
            if (decision === 'cancel') return;
            if (decision === 'no') {
                resetToNewQuery();
                return;
            }

            if (savedQueries.some(query => query.id === activeQueryId)) {
                if (!saveCurrentQuery()) return;
                resetToNewQuery();
            } else {
                pendingNewQueryAfterSave = true;
                openQueryNameModal();
            }
        }

        function startNewQuery() {
            if (getQueryDefinition() !== savedQueryDefinition) {
                const modal = document.getElementById('newQueryConfirmModal');
                modal.classList.remove('hidden');
                return;
            }
            resetToNewQuery();
        }

        function resetToNewQuery() {
            appState.canvasTables = [];
            appState.joins = [];
            activeQueryId = null;
            activeSheetQuery = null;
            isNewQueryDraft = true;
            activeSqlText = null;
            lastQueryResults = [];
            lastQueryResultColumns = [];
            updateQueryDesignTabLabel();
            updateSqlButtonState();
            initQueryGrid();
            savedQueryDefinition = getQueryDefinition();
            updateSaveButtonState();
            refreshNavigationPane();
            renderCanvasTables();
            renderSvgJoins();
            updateActionButtonsState();
            document.getElementById('statusBarMode').textContent = 'Nouvelle requête';
        }

        function persistQuery(query) {
            Object.assign(query, JSON.parse(getQueryDefinition()), { updatedAt: new Date().toISOString() });
            if (!persistSavedQueries()) return false;
            activeQueryId = query.id;
            activeSheetQuery = null;
            isNewQueryDraft = false;
            updateSqlButtonState();
            savedQueryDefinition = getQueryDefinition();
            updateSaveButtonState();
            localStorage.removeItem(QUERY_DRAFT_KEY);
            renderSavedQueries();
            updateQueryDesignTabLabel();
            document.getElementById('statusBarMode').textContent = `Requête « ${query.name} » enregistrée`;
            return true;
        }

        function loadSavedQuery(queryId) {
            const query = savedQueries.find(item => item.id === queryId);
            if (!query) return;
            if (getQueryDefinition() !== savedQueryDefinition && !confirm('Les modifications non enregistrées seront perdues. Continuer ?')) return;

            activeSheetQuery = null;
            appState.canvasTables = query.canvasTables || [];
            appState.joins = query.joins || [];
            appState.queryGridRows = query.queryGridRows || [];
            activeSqlText = typeof query.sqlText === 'string' ? query.sqlText : null;
            if (appState.queryGridRows.length === 0) initQueryGrid();
            else renderQueryGrid();
            activeQueryId = query.id;
            isNewQueryDraft = false;
            updateQueryDesignTabLabel();
            updateSqlButtonState();
            savedQueryDefinition = getQueryDefinition();
            updateSaveButtonState();
            refreshNavigationPane();
            renderCanvasTables();
            renderSvgJoins();
            renderSavedQueries();
            showQueryDesignTab();
        }

        function deleteSavedQuery(queryId) {
            const query = savedQueries.find(item => item.id === queryId);
            if (!query || !confirm(`Supprimer la requête « ${query.name} » ?`)) return;
            const previousQueries = savedQueries;
            savedQueries = savedQueries.filter(item => item.id !== queryId);
            if (!persistSavedQueries()) {
                savedQueries = previousQueries;
                return;
            }
            if (activeQueryId === queryId) activeQueryId = null;
            updateQueryDesignTabLabel();
            updateSqlButtonState();
            updateSaveButtonState();
            renderSavedQueries();
            document.getElementById('statusBarMode').textContent = `Requête « ${query.name} » supprimée`;
        }

        function restoreQueryDraft() {
            let draft;
            try {
                draft = JSON.parse(localStorage.getItem(QUERY_DRAFT_KEY) || 'null');
            } catch (error) {
                localStorage.removeItem(QUERY_DRAFT_KEY);
                return;
            }
            if (!draft || !confirm('Une requête non enregistrée a été récupérée. Restaurer les modifications ?')) {
                localStorage.removeItem(QUERY_DRAFT_KEY);
                return;
            }
            appState.canvasTables = draft.canvasTables || [];
            appState.joins = draft.joins || [];
            appState.queryGridRows = draft.queryGridRows || appState.queryGridRows;
            activeQueryId = draft.activeQueryId || null;
            activeSheetQuery = null;
            isNewQueryDraft = !activeQueryId;
            activeSqlText = typeof draft.sqlText === 'string' ? draft.sqlText : null;
            updateQueryDesignTabLabel();
            updateSqlButtonState();
            renderQueryGrid();
            refreshNavigationPane();
            renderCanvasTables();
            renderSvgJoins();
        }

        window.addEventListener('beforeunload', event => {
            if (skipUnloadWarning || getQueryDefinition() === savedQueryDefinition) return;
            try {
                localStorage.setItem(QUERY_DRAFT_KEY, JSON.stringify({
                    canvasTables: appState.canvasTables,
                    joins: appState.joins,
                    queryGridRows: appState.queryGridRows,
                    activeQueryId,
                    sqlText: activeSqlText
                }));
            } catch (error) {
                // L'avertissement de fermeture reste actif même si le stockage est indisponible.
            }
            event.preventDefault();
            event.returnValue = '';
        });

        // ===== Importation et traitement des fichiers XML =====
        // Sélectionner un fichier XML local
        async function handleFileSelect(event) {
            const files = Array.from(event.target.files || []);
            event.target.value = '';
            if (files.length === 0) return;

            const addedTableNames = [];
            const existingTableNames = [];
            const errors = [];
            showXmlLoading(`Préparation de ${files.length} fichier(s) XML...`, 0);

            for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
                const file = files[fileIndex];
                const fileProgress = progress => {
                    const totalProgress = ((fileIndex + progress) / files.length) * 100;
                    showXmlLoading(`Lecture ${fileIndex + 1}/${files.length} : ${file.name}`, totalProgress);
                };

                try {
                    const xmlString = await readXmlFile(file, fileProgress);
                    showXmlLoading(`Analyse ${fileIndex + 1}/${files.length} : ${file.name}`, null);
                    const tableMap = await parseXmlInWorker(xmlString);
                    const result = loadXMLTableMap(tableMap, file.name, false, true);
                    addedTableNames.push(...result.addedTableNames);
                    existingTableNames.push(...result.existingTableNames);
                } catch (error) {
                    errors.push(`${file.name} : ${error.message || error}`);
                }
            }

            hideXmlLoading();
            const details = [`${files.length} fichier(s) traité(s) :`];
            files.forEach(file => details.push(`- ${file.name}`));
            if (addedTableNames.length > 0) {
                details.push(`${addedTableNames.length} table(s) ajoutée(s)`);
            }
            if (existingTableNames.length > 0) {
                details.push(`${existingTableNames.length} table(s) déjà présente(s), ignorée(s)`);
            }
            if (addedTableNames.length === 0 && existingTableNames.length === 0 && errors.length === 0) {
                details.push('Aucune table trouvée dans les fichiers.');
            }
            if (errors.length > 0) {
                details.push(`Erreurs :`);
                errors.forEach(error => details.push(`- ${error}`));
            }
            showXmlImportNotification(
                details.join('\n'),
                errors.length > 0 ? 'Import terminé avec des erreurs' : 'Import XML réussi'
            );
        }

        function readXmlFile(file, onProgress) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onprogress = progressEvent => {
                    if (progressEvent.lengthComputable) onProgress(progressEvent.loaded / progressEvent.total);
                };
                reader.onload = () => resolve(reader.result);
                reader.onerror = () => reject(reader.error || new Error('Impossible de lire ce fichier XML.'));
                reader.onabort = () => reject(new Error('Lecture du fichier annulée.'));
                reader.readAsText(file, 'UTF-8');
            });
        }

        function showXmlLoading(message, progress) {
            document.getElementById('xmlLoadingOverlay').classList.remove('hidden');
            document.getElementById('xmlLoadingMessage').textContent = message;
            const progressBar = document.getElementById('xmlLoadingProgress');
            const progressText = document.getElementById('xmlLoadingPercent');
            if (progress === null) {
                progressBar.classList.add('xml-loading-indeterminate');
                progressBar.style.width = '';
                progressText.textContent = 'Analyse';
                return;
            }
            progressBar.classList.remove('xml-loading-indeterminate');
            progressBar.style.width = `${Math.max(0, Math.min(100, progress))}%`;
            progressText.textContent = `${Math.round(progress)}%`;
        }

        function hideXmlLoading() {
            document.getElementById('xmlLoadingOverlay').classList.add('hidden');
        }

        function showExportLoading() {
            const overlay = document.getElementById('exportLoadingOverlay');
            if (!overlay) return;
            overlay.classList.remove('hidden');
            overlay.classList.add('flex');
        }

        function hideExportLoading() {
            const overlay = document.getElementById('exportLoadingOverlay');
            if (!overlay) return;
            overlay.classList.add('hidden');
            overlay.classList.remove('flex');
        }

        function normalizeXmlTagName(tagName) {
            if (tagName === null || tagName === undefined) return '';
            return String(tagName).replace(/^.*:/, '').trim();
        }

        function extractXmlRowData(node) {
            const row = {};

            Array.from(node.attributes || []).forEach(attribute => {
                const name = normalizeXmlTagName(attribute.name);
                if (name) row[name] = String(attribute.value || '').trim();
            });

            Array.from(node.children || []).forEach(field => {
                const fieldName = normalizeXmlTagName(field.tagName);
                if (!fieldName) return;

                if (field.children.length === 0) {
                    row[fieldName] = (field.textContent || '').trim();
                    return;
                }

                const nestedValues = extractXmlRowData(field);
                Object.entries(nestedValues).forEach(([nestedKey, nestedValue]) => {
                    row[fieldName + '.' + nestedKey] = nestedValue;
                });
            });

            return row;
        }

        function parseXmlInWorker(xmlString) {
            if (typeof Worker === 'undefined' || typeof Blob === 'undefined' || !window.URL || typeof window.URL.createObjectURL !== 'function') {
                return Promise.resolve(parseXmlTableMapOnMainThread(xmlString));
            }
            return new Promise((resolve, reject) => {
                const workerSource = String.raw`
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

                    const extractXmlRowData = node => {
                        const row = {};
                        for (let i = 0; i < (node.attributes || []).length; i++) {
                            const attribute = node.attributes[i];
                            const name = normalizeXmlTagName(attribute.name);
                            if (name) row[name] = String(attribute.value || '').trim();
                        }
                        for (let i = 0; i < node.children.length; i++) {
                            const field = node.children[i];
                            const fieldName = normalizeXmlTagName(field.tagName);
                            if (!fieldName) continue;
                            if (field.children.length === 0) {
                                row[fieldName] = decodeXml((field.textContent || '').trim());
                                continue;
                            }
                            const nestedValues = extractXmlRowData(field);
                            for (const [nestedKey, nestedValue] of Object.entries(nestedValues)) {
                                row[fieldName + '.' + nestedKey] = nestedValue;
                            }
                        }
                        return row;
                    };

                    const parseXmlWithoutDomParser = xmlString => {
                        const rootMatch = xmlString.match(/<([A-Za-z_][\w:.-]*)(?:\s[^>]*)?>([\s\S]*)<\/\1\s*>/);
                        if (!rootMatch) throw new Error('Erreur de syntaxe XML détectée dans le fichier.');

                        const tableMap = {};
                        const rowPattern = /<([A-Za-z_][\w:.-]*)(?:\s[^>]*)?>([\s\S]*?)<\/\1\s*>|<([A-Za-z_][\w:.-]*)\s*\/>/g;
                        let rowMatch;
                        while ((rowMatch = rowPattern.exec(rootMatch[2])) !== null) {
                            const tableName = normalizeXmlTagName(rowMatch[1] || rowMatch[3]);
                            const row = {};
                            const fieldPattern = /<([A-Za-z_][\w:.-]*)(?:\s[^>]*)?>([\s\S]*?)<\/\1\s*>|<([A-Za-z_][\w:.-]*)\s*\/>/g;
                            let fieldMatch;
                            while ((fieldMatch = fieldPattern.exec(rowMatch[2] || '')) !== null) {
                                const fieldName = normalizeXmlTagName(fieldMatch[1] || fieldMatch[3]);
                                if (fieldName) row[fieldName] = decodeXml((fieldMatch[2] || '').trim());
                            }
                            (tableMap[tableName] || (tableMap[tableName] = [])).push(row);
                        }
                        return tableMap;
                    };

                    self.onmessage = event => {
                        try {
                            let tableMap;
                            if (typeof DOMParser !== 'undefined') {
                                const parser = new DOMParser();
                                const xmlDoc = parser.parseFromString(event.data, 'text/xml');
                                if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
                                    throw new Error('Erreur de syntaxe XML détectée dans le fichier.');
                                }
                                tableMap = {};
                                const children = xmlDoc.documentElement.children;
                                for (let i = 0; i < children.length; i++) {
                                    const node = children[i];
                                    const tableName = normalizeXmlTagName(node.tagName);
                                    const rows = tableMap[tableName] || (tableMap[tableName] = []);
                                    rows.push(extractXmlRowData(node));
                                }
                            } else {
                                tableMap = parseXmlWithoutDomParser(event.data);
                            }
                            self.postMessage({ tableMap });
                        } catch (error) {
                            self.postMessage({ error: error.message || String(error) });
                        }
                    };
                `;
                const workerUrl = URL.createObjectURL(new Blob([workerSource], { type: 'application/javascript' }));
                const worker = new Worker(workerUrl);
                const finish = () => {
                    worker.terminate();
                    URL.revokeObjectURL(workerUrl);
                };
                worker.onmessage = event => {
                    finish();
                    if (event.data.error) reject(new Error(event.data.error));
                    else resolve(event.data.tableMap);
                };
                worker.onerror = event => {
                    finish();
                    reject(new Error(event.message || 'Impossible d’analyser le fichier XML.'));
                };
                worker.postMessage(xmlString);
            });
        }

        // Parser le XML avec DOMParser et peupler l'application
        function loadXMLFromString(xmlString, fileName, showNotification = true, throwOnError = false) {
            try {
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(xmlString, "text/xml");

                // Vérifier les erreurs de parsing
                const parseError = xmlDoc.getElementsByTagName("parsererror");
                if (parseError.length > 0) {
                    throw new Error('Erreur de syntaxe XML détectée dans le fichier.');
                }

                appState.fileName = fileName;
                document.getElementById('currentProjectTitle').textContent = fileName;

                // Parcourir tous les éléments enfants de la racine
                const root = xmlDoc.documentElement;
                const children = root.children;

                // Regrouper par nom de balise (table)
                const tableMap = {};
                for (let i = 0; i < children.length; i++) {
                    const node = children[i];
                    const tagName = normalizeXmlTagName(node.tagName);
                    if (!tableMap[tagName]) {
                        tableMap[tagName] = [];
                    }

                    tableMap[tagName].push(extractXmlRowData(node));
                }

                return loadXMLTableMap(tableMap, fileName, showNotification, throwOnError);

            } catch (err) {
                console.error(err);
                if (throwOnError) throw err;
                alert("Erreur lors du traitement du fichier XML : " + err.message);
                return null;
            }
        }

        function loadXMLTableMap(tableMap, fileName, showNotification = true, throwOnError = false) {
            try {
                appState.fileName = fileName;
                document.getElementById('currentProjectTitle').textContent = fileName;

                const addedTableNames = [];
                const existingTableNames = [];
                const xmlSourceName = fileName || 'XML sans nom';
                const xmlSourceColors = ['#2563eb', '#16a34a', '#c2410c', '#9333ea', '#0891b2', '#be123c', '#4d7c0f', '#b45309'];
                Object.entries(tableMap).forEach(([tableName, rows]) => {
                    const existingTable = Object.keys(appState.tables).find(name => name.toLowerCase() === tableName.toLowerCase());
                    if (existingTable) {
                        existingTableNames.push(tableName);
                        return;
                    }
                    if (!appState.xmlSources[xmlSourceName]) {
                        const sourceIndex = Object.keys(appState.xmlSources).length % xmlSourceColors.length;
                        appState.xmlSources[xmlSourceName] = xmlSourceColors[sourceIndex];
                    }
                    appState.tables[tableName] = rows;
                    if (!appState.tableOrder.includes(tableName)) appState.tableOrder.push(tableName);
                    appState.tableSources[tableName] = xmlSourceName;
                    addedTableNames.push(tableName);
                });

                initQueryGrid();

                refreshNavigationPane();
                renderCanvasTables();
                renderSvgJoins();
                updateActionButtonsState();
                updateStatusBar();

                if (showNotification) {
                    const details = [];
                    if (addedTableNames.length > 0) {
                        details.push(`${addedTableNames.length} table(s) ajoutée(s)`);
                    }
                    if (existingTableNames.length > 0) {
                        details.push(`${existingTableNames.length} table(s) déjà présente(s), ignorée(s)`);
                    }
                    if (addedTableNames.length === 0 && existingTableNames.length === 0) {
                        details.push('Aucune table trouvée dans le fichier.');
                    }
                    details.unshift(`Fichier : ${fileName}`);
                    showXmlImportNotification(details.join(' · '));
                }
                return { addedTableNames, existingTableNames };

            } catch (err) {
                console.error(err);
                if (throwOnError) throw err;
                alert("Erreur lors du traitement du fichier XML : " + err.message);
                return null;
            }
        }

        function showXmlImportNotification(details, title = 'Import XML réussi') {
            document.getElementById('xmlImportNotificationTitle').textContent = title;
            document.getElementById('xmlImportNotificationDetails').textContent = details;
            document.getElementById('xmlImportNotification').classList.remove('hidden');
            document.getElementById('xmlImportNotification').classList.add('flex');
        }

        function hideXmlImportNotification() {
            document.getElementById('xmlImportNotification').classList.add('hidden');
            document.getElementById('xmlImportNotification').classList.remove('flex');
        }

        // ===== Volet de navigation et selection des tables =====
        // Rafraîchir le volet de navigation (gauche)
        let tableSelectionMode = false;
        let selectedTableNames = new Set();

        function updateNavigationTabCounts() {
            const tableCount = document.getElementById('tableTabCount');
            const queryCount = document.getElementById('queryTabCount');
            if (tableCount) tableCount.textContent = Object.keys(appState.tables).length;
            if (queryCount) queryCount.textContent = savedQueries.length + sheetQueries.length;
        }

        function getTableSourceColor(tableName) {
            const sourceName = appState.tableSources[tableName] || '';
            return appState.xmlSources[sourceName] || '#6b7280';
        }

        function isKeyFieldName(fieldName) {
            const name = String(fieldName || '').trim();
            return /^(?:id_|cd_|[A-Za-z]{2}_)/i.test(name);
        }

        function refreshNavigationPane() {
            const listContainer = document.getElementById('navTableList');
            updateNavigationTabCounts();
            listContainer.innerHTML = '';

            const tableNames = [
                ...appState.tableOrder.filter(name => Object.prototype.hasOwnProperty.call(appState.tables, name)),
                ...Object.keys(appState.tables).filter(name => !appState.tableOrder.includes(name))
            ];
            appState.tableOrder = tableNames;
            if (tableNames.length === 0) {
                listContainer.innerHTML = '<div class="text-gray-400 text-center py-4 italic">Aucune table chargée.</div>';
                updateTableSelectionToolbar();
                return;
            }

            tableNames.forEach(tName => {
                const count = appState.tables[tName].length;
                const isOnCanvas = appState.canvasTables.some(ct => ct.name === tName);
                const sourceName = appState.tableSources[tName] || '';
                const sourceColor = appState.xmlSources[sourceName] || '#6b7280';

                const div = document.createElement('div');
                div.className = `flex items-center justify-between p-1.5 rounded cursor-grab group ${selectedTableNames.has(tName) ? 'bg-blue-50' : 'hover:bg-gray-100'}`;
                div.dataset.tableName = tName;
                div.title = tableSelectionMode ? 'Cochez les tables à supprimer' : 'Glisser vers le schéma ou sur une autre table pour réordonner. Maintenir pour sélectionner.';
                div.addEventListener('mousedown', event => startNavigationTableDrag(event, tName));
                div.innerHTML = `
                    <div class="table-list-content flex items-center space-x-2 truncate">
                        <i class="fa-solid fa-table xml-source-table-icon"></i>
                        <span data-table-label class="font-medium text-gray-800 truncate"></span>
                        <span class="text-[10px] bg-gray-200 px-1.5 py-0.2 rounded text-gray-600">${count} lignes</span>
                    </div>
                    <div class="flex items-center space-x-1">
                        ${!isOnCanvas ? `<i class="fa-solid fa-grip-vertical text-gray-400" title="Glisser vers le schéma"></i>` : `<i class="fa-solid fa-check text-emerald-600 text-[10px]" title="Présente sur le schéma"></i>`}
                    </div>
                `;
                if (tableSelectionMode) {
                    const checkbox = document.createElement('input');
                    checkbox.type = 'checkbox';
                    checkbox.checked = selectedTableNames.has(tName);
                    checkbox.className = 'mr-1 accent-red-700';
                    checkbox.setAttribute('aria-label', `Sélectionner ${tName}`);
                    checkbox.addEventListener('change', () => {
                        if (checkbox.checked) selectedTableNames.add(tName);
                        else selectedTableNames.delete(tName);
                        updateTableSelectionToolbar();
                    });
                    div.querySelector('.table-list-content').prepend(checkbox);
                }
                const tableIcon = div.querySelector('.xml-source-table-icon');
                tableIcon.style.color = sourceColor;
                tableIcon.title = sourceName ? `Importé depuis ${sourceName}` : 'Fichier XML source inconnu';
                const tableLabel = div.querySelector('[data-table-label]');
                tableLabel.textContent = tName;
                tableLabel.addEventListener('click', () => openTablePreview(tName));
                tableLabel.addEventListener('dblclick', () => openTablePreview(tName));
                const columnNames = Object.keys(appState.tables[tName][0] || {});
                tableLabel.title = columnNames.length > 0
                    ? `Cliquez pour ouvrir la table dans un onglet.\nColonnes (${columnNames.length}) :\n${columnNames.join('\n')}`
                    : 'Aucune colonne trouvée';
                listContainer.appendChild(div);
            });
            updateTableSelectionToolbar();
        }

        function updateTableSelectionToolbar() {
            const toolbar = document.getElementById('tableSelectionToolbar');
            const selectAllCheckbox = document.getElementById('selectAllTablesCheckbox');
            if (!toolbar) return;
            toolbar.classList.toggle('hidden', !tableSelectionMode);
            const selectedCount = selectedTableNames.size;
            document.getElementById('tableSelectionCount').textContent = `${selectedCount} sélectionnée${selectedCount === 1 ? '' : 's'}`;
            const totalCount = Object.keys(appState.tables).length;
            if (selectAllCheckbox) {
                selectAllCheckbox.checked = totalCount > 0 && selectedCount === totalCount;
                selectAllCheckbox.indeterminate = totalCount > 0 && selectedCount > 0 && selectedCount < totalCount;
            }
            document.getElementById('deleteSelectedTablesButton').disabled = selectedCount === 0;
        }

        function clearTableSelection() {
            tableSelectionMode = false;
            selectedTableNames.clear();
            refreshNavigationPane();
        }

        function toggleSelectAllTables() {
            const checkbox = document.getElementById('selectAllTablesCheckbox');
            if (!checkbox) return;
            const tableNames = Object.keys(appState.tables);
            if (checkbox.checked) {
                selectedTableNames = new Set(tableNames);
            } else {
                selectedTableNames.clear();
            }
            refreshNavigationPane();
        }

        function switchNavigationTab(tabName) {
            const showTables = tabName === 'tables';
            document.getElementById('navTablesPanel').classList.toggle('hidden', !showTables);
            document.getElementById('navQueriesPanel').classList.toggle('hidden', showTables);
            document.getElementById('tablesTab').setAttribute('aria-selected', String(showTables));
            document.getElementById('queriesTab').setAttribute('aria-selected', String(!showTables));
            document.getElementById('tablesTab').className = showTables
                ? 'flex-1 px-3 py-1.5 font-semibold text-red-700 border-b-2 border-red-700 bg-white'
                : 'flex-1 px-3 py-1.5 text-gray-700 border-b-2 border-transparent hover:bg-white';
            document.getElementById('queriesTab').className = showTables
                ? 'flex-1 px-3 py-1.5 text-gray-700 border-b-2 border-transparent hover:bg-white'
                : 'flex-1 px-3 py-1.5 font-semibold text-red-700 border-b-2 border-red-700 bg-white';
            if (!showTables && !sheetQueriesLoaded) loadSheetQueries();
        }

        let draggedNavigationTable = null;

        // ===== Cartes de tables, canvas et interactions deplacement =====
        function startNavigationTableDrag(event, tableName) {
            if (event.button !== 0 || event.target.closest('button, input') || tableSelectionMode) return;
            event.preventDefault();
            const dragState = {
                tableName,
                startX: event.clientX,
                startY: event.clientY,
                sourceElement: event.currentTarget,
                active: false,
                longPressed: false,
                longPressTimer: null
            };
            draggedNavigationTable = dragState;
            dragState.longPressTimer = setTimeout(() => {
                if (draggedNavigationTable !== dragState || dragState.active) return;
                dragState.longPressed = true;
                tableSelectionMode = true;
                selectedTableNames.add(tableName);
                refreshNavigationPane();
            }, 550);
            document.addEventListener('mousemove', moveNavigationTableDrag);
            document.addEventListener('mouseup', finishNavigationTableDrag);
        }

        function moveNavigationTableDrag(event) {
            if (!draggedNavigationTable || draggedNavigationTable.longPressed) return;
            if (!draggedNavigationTable.active && Math.hypot(event.clientX - draggedNavigationTable.startX, event.clientY - draggedNavigationTable.startY) < 5) return;

            event.preventDefault();
            clearTimeout(draggedNavigationTable.longPressTimer);
            draggedNavigationTable.active = true;
            draggedNavigationTable.sourceElement.classList.add('opacity-50');
            document.body.style.cursor = 'grabbing';
            document.body.style.userSelect = 'none';
            clearNavigationTableDropTarget();

            const target = document.elementFromPoint(event.clientX, event.clientY);
            const canvas = target ? target.closest('#erdCanvas') : null;
            const navigationRow = target ? target.closest('#navTableList > [data-table-name]') : null;
            if (canvas) canvas.classList.add('ring-2', 'ring-blue-500');
            else if (navigationRow) navigationRow.classList.add('border-t-2', 'border-blue-600');
        }

        function finishNavigationTableDrag(event) {
            if (!draggedNavigationTable) return;
            const dragState = draggedNavigationTable;
            clearTimeout(dragState.longPressTimer);
            const target = document.elementFromPoint(event.clientX, event.clientY);
            const canvas = target ? target.closest('#erdCanvas') : null;
            const navigationRow = target ? target.closest('#navTableList > [data-table-name]') : null;

            if (dragState.active && canvas) {
                const canvasRect = canvas.getBoundingClientRect();
                addTableToCanvas(dragState.tableName, {
                    x: Math.max(0, event.clientX - canvasRect.left + canvas.scrollLeft - 110),
                    y: Math.max(0, event.clientY - canvasRect.top + canvas.scrollTop - 20)
                });
            } else if (dragState.active && navigationRow) {
                reorderNavigationTable(dragState.tableName, navigationRow.dataset.tableName);
            }

            dragState.sourceElement.classList.remove('opacity-50');
            clearNavigationTableDropTarget();
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            draggedNavigationTable = null;
            document.removeEventListener('mousemove', moveNavigationTableDrag);
            document.removeEventListener('mouseup', finishNavigationTableDrag);
        }

        function clearNavigationTableDropTarget() {
            document.getElementById('erdCanvas').classList.remove('ring-2', 'ring-blue-500');
            document.querySelectorAll('#navTableList > div').forEach(item => {
                item.classList.remove('border-t-2', 'border-blue-600');
            });
        }

        function reorderNavigationTable(sourceTable, targetTable) {
            if (!sourceTable || sourceTable === targetTable) return;
            const order = [...appState.tableOrder];
            const sourceIndex = order.indexOf(sourceTable);
            const targetIndex = order.indexOf(targetTable);
            if (sourceIndex < 0 || targetIndex < 0) return;

            const [movedTable] = order.splice(sourceIndex, 1);
            order.splice(targetIndex, 0, movedTable);
            appState.tableOrder = order;
            refreshNavigationPane();
        }

        // Filtrer les tables dans le panneau de navigation
        function filterNavTables() {
            const query = document.getElementById('navSearchInput').value.toLowerCase();
            const items = document.getElementById('navTableList').children;
            for (let item of items) {
                const text = item.textContent.toLowerCase();
                item.style.display = text.includes(query) ? 'flex' : 'none';
            }
        }

        // Rendu des cartes de tables sur le canevas ERD
        function clampCanvasPosition(tableDef) {
            const canvas = document.getElementById('erdCanvas');
            if (!canvas || !tableDef) return tableDef;
            const width = Number(tableDef.width) || 220;
            const height = Number(tableDef.height) || 220;
            const maxX = Math.max(20, canvas.clientWidth - width - 20);
            const maxY = Math.max(20, canvas.clientHeight - height - 20);
            tableDef.x = Math.max(20, Math.min(Number(tableDef.x) || 40, maxX));
            tableDef.y = Math.max(20, Math.min(Number(tableDef.y) || 40, maxY));
            return tableDef;
        }

        function renderCanvasTables() {
            const container = document.getElementById('tablesContainer');
            container.innerHTML = '';

            appState.canvasTables.forEach(ct => {
                clampCanvasPosition(ct);
                const tableData = appState.tables[ct.name] || [];
                const fields = tableData.length > 0 ? Object.keys(tableData[0]) : [];
                const tableColor = getTableSourceColor(ct.name);

                const card = document.createElement('div');
                card.className = "table-card pointer-events-auto";
                card.style.left = ct.x + 'px';
                card.style.top = ct.y + 'px';
                card.style.width = ct.width + 'px';
                card.style.height = ct.height + 'px';
                card.dataset.tableName = ct.name;

                let fieldsHtml = '';
                fields.forEach((fName, fIdx) => {
                    const keyIcon = isKeyFieldName(fName)
                        ? '<i class="fa-solid fa-key text-amber-600 mr-1 text-[9px]"></i>'
                        : '';
                    fieldsHtml += `
                        <div class="field-row" id="field_${ct.name}_${fName}" data-table-name="${ct.name}" data-field-name="${fName}" draggable="true" onclick="handleFieldClick('${ct.name}', '${fName}')" ondblclick="handleFieldDoubleClick('${ct.name}', '${fName}')" ondragstart="startFieldDrag(event, '${ct.name}', '${fName}')" ondragover="allowFieldDrop(event)" ondrop="handleFieldDrop(event, '${ct.name}', '${fName}')" title="Double-cliquez pour ajouter au bas, ou cliquez pour lier les tables">
                            <span class="truncate">${keyIcon}${fName}</span>
                            <span class="text-[9px] text-gray-400">txt</span>
                        </div>
                    `;
                });

                card.innerHTML = `
                    <header class="flex items-center justify-between cursor-move" onmousedown="initTableDrag(event, '${ct.name}')">
                        <span class="truncate" title="${ct.name}"><i class="fa-solid fa-table mr-1" style="color: ${tableColor};"></i>${ct.name}</span>
                        <div class="flex items-center space-x-1">
                            <button onclick="removeTableFromCanvas('${ct.name}')" class="text-gray-500 hover:text-red-700 font-bold px-1" title="Retirer du schéma">×</button>
                        </div>
                    </header>
                    <div class="overflow-y-auto bg-white table-fields-scroll" style="height: calc(100% - 26px);">
                        ${fieldsHtml}
                    </div>
                    <div class="resize-handle" onmousedown="initTableResize(event, '${ct.name}')"></div>
                `;

                const fieldsScroll = card.querySelector('.table-fields-scroll');
                if (fieldsScroll) {
                    fieldsScroll.addEventListener('scroll', () => {
                        requestAnimationFrame(renderSvgJoins);
                    }, { passive: true });
                }

                container.appendChild(card);
            });
            syncQueryFieldHighlights();
            renderSvgJoins();
        }

        // Gestion du Drag & Drop des cartes de tables
        let activeDragTable = null;
        let dragOffsetX = 0;
        let dragOffsetY = 0;

        function initTableDrag(e, tableName) {
            if (e.target.tagName === 'BUTTON') return;
            activeDragTable = tableName;
            const ct = appState.canvasTables.find(t => t.name === tableName);
            const card = document.querySelector(`.table-card[data-table-name="${tableName}"]`);
            const rect = card.getBoundingClientRect();
            const canvasRect = document.getElementById('erdCanvas').getBoundingClientRect();

            dragOffsetX = e.clientX - rect.left;
            dragOffsetY = e.clientY - rect.top;

            document.addEventListener('mousemove', onTableDrag);
            document.addEventListener('mouseup', stopTableDrag);
            e.stopPropagation();
        }

        function onTableDrag(e) {
            if (!activeDragTable) return;
            const canvas = document.getElementById('erdCanvas');
            const canvasRect = canvas.getBoundingClientRect();
            const ct = appState.canvasTables.find(t => t.name === activeDragTable);
            if (ct) {
                const maxX = Math.max(20, canvas.clientWidth - ct.width - 20);
                const maxY = Math.max(20, canvas.clientHeight - ct.height - 20);
                ct.x = Math.max(20, Math.min(e.clientX - canvasRect.left - dragOffsetX + canvas.scrollLeft, maxX));
                ct.y = Math.max(20, Math.min(e.clientY - canvasRect.top - dragOffsetY + canvas.scrollTop, maxY));
                const card = document.querySelector(`.table-card[data-table-name="${activeDragTable}"]`);
                if (card) {
                    card.style.left = ct.x + 'px';
                    card.style.top = ct.y + 'px';
                }
                renderSvgJoins();
            }
        }

        function stopTableDrag() {
            activeDragTable = null;
            document.removeEventListener('mousemove', onTableDrag);
            document.removeEventListener('mouseup', stopTableDrag);
        }

        // Redimensionnement des cartes de tables
        let activeResizeTable = null;
        let initialResizeX = 0;
        let initialResizeY = 0;
        let initialWidth = 220;
        let initialHeight = 220;

        function initTableResize(e, tableName) {
            activeResizeTable = tableName;
            const ct = appState.canvasTables.find(t => t.name === tableName);
            initialResizeX = e.clientX;
            initialResizeY = e.clientY;
            initialWidth = ct.width;
            initialHeight = ct.height;

            document.addEventListener('mousemove', onTableResize);
            document.addEventListener('mouseup', stopTableResize);
            e.stopPropagation();
        }

        function onTableResize(e) {
            if (!activeResizeTable) return;
            const ct = appState.canvasTables.find(t => t.name === activeResizeTable);
            if (ct) {
                ct.width = Math.max(180, initialWidth + (e.clientX - initialResizeX));
                ct.height = Math.max(150, initialHeight + (e.clientY - initialResizeY));
                const card = document.querySelector(`.table-card[data-table-name="${activeResizeTable}"]`);
                if (card) {
                    card.style.width = ct.width + 'px';
                    card.style.height = ct.height + 'px';
                }
                renderSvgJoins();
            }
        }

        function stopTableResize() {
            activeResizeTable = null;
            document.removeEventListener('mousemove', onTableResize);
            document.removeEventListener('mouseup', stopTableResize);
        }

        // Retirer une table du schéma
        function removeTableFromCanvas(tableName) {
            appState.canvasTables = appState.canvasTables.filter(t => t.name !== tableName);
            // Supprimer aussi les liaisons associées
            appState.joins = appState.joins.filter(j => j.fromTable !== tableName && j.toTable !== tableName);
            refreshNavigationPane();
            renderCanvasTables();
            renderSvgJoins();
        }

        function deleteTableFromProject(tableName) {
            if (!Object.prototype.hasOwnProperty.call(appState.tables, tableName)) return;
            if (!confirm(`Supprimer complètement la table « ${tableName} » et toutes ses données ?`)) return;

            removeTablesFromProject([tableName]);
        }

        function deleteSelectedTables() {
            const tableNames = [...selectedTableNames].filter(name => Object.prototype.hasOwnProperty.call(appState.tables, name));
            if (tableNames.length === 0) return;
            if (!confirm(`Supprimer les ${tableNames.length} tables sélectionnées et toutes leurs données ?`)) return;

            tableSelectionMode = false;
            selectedTableNames.clear();
            removeTablesFromProject(tableNames);
        }

        function removeTablesFromProject(tableNames) {
            const namesToDelete = new Set(tableNames);
            const lowerCaseNames = new Set(tableNames.map(name => name.toLowerCase()));
            if (linkingSource && namesToDelete.has(linkingSource.table)) clearFieldJoinSelection();
            namesToDelete.forEach(tableName => {
                delete appState.tables[tableName];
                delete appState.tableSources[tableName];
            });
            appState.tableOrder = appState.tableOrder.filter(name => !namesToDelete.has(name));
            appState.canvasTables = appState.canvasTables.filter(table => !namesToDelete.has(table.name));
            appState.joins = appState.joins.filter(join => !namesToDelete.has(join.fromTable) && !namesToDelete.has(join.toTable));
            appState.queryGridRows = appState.queryGridRows.map(queryRow => {
                if (!queryRow.table || !lowerCaseNames.has(queryRow.table.toLowerCase())) return queryRow;
                return { field: '', table: '', alias: '', expression: '', total: 'GROUP BY', show: false, sort: '', criteria: '', or: '', width: queryRow.width };
            });
            selectedTableNames = new Set([...selectedTableNames].filter(name => !namesToDelete.has(name)));

            renderQueryGrid();
            refreshNavigationPane();
            renderCanvasTables();
            renderSvgJoins();
            updateActionButtonsState();
            updateStatusBar();
        }

        function deleteAllTables() {
            const tableCount = Object.keys(appState.tables).length;
            if (tableCount === 0) return;
            if (!confirm(`Supprimer les ${tableCount} tables et toutes leurs données ?`)) return;

            tableSelectionMode = false;
            selectedTableNames.clear();
            if (linkingSource) clearFieldJoinSelection();
            appState.tables = {};
            appState.tableOrder = [];
            appState.tableSources = {};
            appState.xmlSources = {};
            appState.canvasTables = [];
            appState.joins = [];
            appState.queryGridRows = [];
            activeQueryId = null;
            activeSqlText = null;
            openTableTabs = [];
            activeWorkspaceTab = 'query-design';
            lastQueryResults = [];
            lastQueryResultColumns = [];
            updateQueryDesignTabLabel();
            updateSqlButtonState();
            updateActionButtonsState();

            initQueryGrid();
            refreshNavigationPane();
            renderCanvasTables();
            renderSvgJoins();
            replaceChildrenCompat(document.getElementById('tableWorkspaceContainer'));
            document.getElementById('tableWorkspaceTitle').textContent = '';
            document.getElementById('exportTableWorkspaceButton').disabled = true;
            closeDataPreviewModal();
            showQueryDesignTab();
            savedQueryDefinition = getQueryDefinition();
            updateSaveButtonState();
            updateStatusBar();
        }

        // Ajouter une table au schéma via modal ou bouton
        function openAddTableModal() {
            const modalList = document.getElementById('addTableModalList');
            modalList.innerHTML = '';

            const tableNames = Object.keys(appState.tables);
            tableNames.forEach(tName => {
                const isOnCanvas = appState.canvasTables.some(ct => ct.name === tName);
                if (!isOnCanvas) {
                    const div = document.createElement('div');
                    div.className = "flex items-center justify-between p-2 hover:bg-gray-100 rounded cursor-pointer";
                    const tableColor = getTableSourceColor(tName);
                    div.innerHTML = `
                        <span class="font-medium text-gray-800"><i class="fa-solid fa-table mr-2" style="color: ${tableColor};"></i>${tName}</span>
                        <button onclick="addTableToCanvas('${tName}'); closeAddTableModal();" class="bg-blue-600 hover:bg-blue-700 text-white px-2 py-1 rounded text-xs">Ajouter</button>
                    `;
                    modalList.appendChild(div);
                }
            });
            if (modalList.children.length === 0) {
                modalList.innerHTML = '<div class="text-gray-400 text-center py-4 italic">Toutes les tables sont déjà sur le schéma.</div>';
            }
            document.getElementById('addTableModal').classList.remove('hidden');
        }

        function closeAddTableModal() {
            document.getElementById('addTableModal').classList.add('hidden');
        }

        function addTableToCanvas(tableName, position = null) {
            if (!appState.canvasTables.some(t => t.name === tableName)) {
                const canvas = document.getElementById('erdCanvas');
                const initialX = position ? position.x : 40 + (appState.canvasTables.length * 30) % 220;
                const initialY = position ? position.y : 40 + (appState.canvasTables.length * 30) % 120;
                const newCard = {
                    name: tableName,
                    x: initialX,
                    y: initialY,
                    width: 220,
                    height: 220
                };
                if (canvas) {
                    const maxX = Math.max(20, canvas.clientWidth - newCard.width - 20);
                    const maxY = Math.max(20, canvas.clientHeight - newCard.height - 20);
                    newCard.x = Math.max(20, Math.min(initialX, maxX));
                    newCard.y = Math.max(20, Math.min(initialY, maxY));
                }
                appState.canvasTables.push(newCard);
                refreshNavigationPane();
                renderCanvasTables();
                renderSvgJoins();
            }
        }

        function allowDrop(event) {
            if (event) {
                event.preventDefault();
                if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
            }
        }

        function dropTableCard(event) {
            if (!event) return;
            event.preventDefault();
            const tableName = event.dataTransfer ? event.dataTransfer.getData('application/x-query-table') : '';
            if (!tableName || !appState.tables[tableName]) return;

            const canvas = document.getElementById('erdCanvas');
            const canvasRect = canvas.getBoundingClientRect();
            addTableToCanvas(tableName, {
                x: Math.max(0, event.clientX - canvasRect.left + canvas.scrollLeft - 110),
                y: Math.max(0, event.clientY - canvasRect.top + canvas.scrollTop - 20)
            });
        }

        // Lier deux champs entre tables par glisser-déposer, clic ou double-clic
        let linkingSource = null;

        // ===== Creation et gestion des jointures =====
        function clearFieldJoinSelection() {
            linkingSource = null;
            document.querySelectorAll('.field-row.join-source').forEach(el => {
                el.classList.remove('join-source');
            });
        }

        function applyFieldJoinSelection(tableName, fieldName) {
            clearFieldJoinSelection();
            linkingSource = { table: tableName, field: fieldName };
            const sourceEl = document.getElementById(`field_${tableName}_${fieldName}`);
            if (sourceEl) sourceEl.classList.add('join-source');
        }

        function createJoinBetween(source, targetTable, targetField) {
            if (!source || !source.table || !source.field) return;
            if (source.table === targetTable && source.field === targetField) return;
            if (!appState.tables[targetTable]) return;

            const exists = appState.joins.some(join =>
                ((join.fromTable === source.table && join.fromField === source.field && join.toTable === targetTable && join.toField === targetField) ||
                 (join.fromTable === targetTable && join.fromField === targetField && join.toTable === source.table && join.toField === source.field))
            );
            if (exists) {
                clearFieldJoinSelection();
                return;
            }

            appState.joins.push({
                id: 'join_' + Date.now(),
                fromTable: source.table,
                fromField: source.field,
                toTable: targetTable,
                toField: targetField,
                type: 'LEFT'
            });
            clearFieldJoinSelection();
            renderSvgJoins();
        }

        function handleFieldClick(tableName, fieldName) {
            if (linkingSource && !(linkingSource.table === tableName && linkingSource.field === fieldName)) {
                createJoinBetween(linkingSource, tableName, fieldName);
                return;
            }

            applyFieldJoinSelection(tableName, fieldName);
        }

        function handleFieldDoubleClick(tableName, fieldName) {
            clearFieldJoinSelection();
            selectFieldForQuery(tableName, fieldName);
        }

        function startFieldDrag(event, tableName, fieldName) {
            applyFieldJoinSelection(tableName, fieldName);
            if (event.dataTransfer) {
                event.dataTransfer.effectAllowed = 'copy';
                event.dataTransfer.setData('application/json', JSON.stringify({ table: tableName, field: fieldName }));
                event.dataTransfer.setData('text/plain', `${tableName}.${fieldName}`);
            }
        }

        function allowFieldDrop(event) {
            event.preventDefault();
            if (event.dataTransfer) {
                event.dataTransfer.dropEffect = 'copy';
            }
        }

        function handleFieldDrop(event, targetTable, targetField) {
            event.preventDefault();
            event.stopPropagation();

            let source = linkingSource;
            if (event.dataTransfer) {
                try {
                    const payload = event.dataTransfer.getData('application/json');
                    if (payload) {
                        const dragData = JSON.parse(payload);
                        source = {
                            table: dragData.table || dragData.tableName,
                            field: dragData.field || dragData.fieldName
                        };
                    } else {
                        const plain = event.dataTransfer.getData('text/plain');
                        if (plain && plain.includes('.')) {
                            const [tableName, fieldName] = plain.split('.');
                            source = { table: tableName, field: fieldName };
                        }
                    }
                } catch (err) {
                    console.warn('Payload drag & drop non valide :', err);
                }
            }

            if (!source) {
                if (linkingSource) {
                    createJoinBetween(linkingSource, targetTable, targetField);
                }
                return;
            }

            createJoinBetween(source, targetTable, targetField);
        }

        // Rendu des traits de liaison SVG précis ancrés aux champs
        function renderSvgJoins() {
            const svg = document.getElementById('erdSvg');
            svg.innerHTML = '';

            const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
            const arrowMarker = document.createElementNS('http://www.w3.org/2000/svg', 'marker');
            arrowMarker.setAttribute('id', 'joinArrowMarker');
            arrowMarker.setAttribute('viewBox', '0 0 8 8');
            arrowMarker.setAttribute('refX', '7');
            arrowMarker.setAttribute('refY', '4');
            arrowMarker.setAttribute('markerWidth', '8');
            arrowMarker.setAttribute('markerHeight', '8');
            arrowMarker.setAttribute('markerUnits', 'userSpaceOnUse');
            arrowMarker.setAttribute('orient', 'auto-start-reverse');
            const arrowHead = document.createElementNS('http://www.w3.org/2000/svg', 'path');
            arrowHead.setAttribute('d', 'M 0 0 L 8 4 L 0 8 z');
            arrowHead.setAttribute('fill', '#b91c1c');
            arrowMarker.appendChild(arrowHead);
            defs.appendChild(arrowMarker);
            svg.appendChild(defs);

            const canvasRect = document.getElementById('erdCanvas').getBoundingClientRect();

            const getCardAnchor = (card, otherCard, fieldEl) => {
                const cardRect = card.getBoundingClientRect();
                const otherRect = otherCard.getBoundingClientRect();
                const isSourceLeft = cardRect.left <= otherRect.left;

                if (fieldEl) {
                    const fieldRect = fieldEl.getBoundingClientRect();
                    const fieldsContainer = fieldEl.closest('.table-fields-scroll');
                    const fieldsRect = fieldsContainer ? fieldsContainer.getBoundingClientRect() : null;
                    const fieldY = fieldRect.top + fieldRect.height / 2;
                    const anchorY = fieldsRect
                        ? Math.max(fieldsRect.top + 1, Math.min(fieldY, fieldsRect.bottom - 1))
                        : fieldY;
                    return {
                        x: (isSourceLeft ? cardRect.right : cardRect.left) - canvasRect.left,
                        y: anchorY - canvasRect.top
                    };
                }

                return {
                    x: isSourceLeft ? cardRect.right - canvasRect.left : cardRect.left - canvasRect.left,
                    y: cardRect.top - canvasRect.top + cardRect.height / 2
                };
            };

            appState.joins.forEach(join => {
                const cardFrom = document.querySelector(`.table-card[data-table-name="${join.fromTable}"]`);
                const cardTo = document.querySelector(`.table-card[data-table-name="${join.toTable}"]`);

                if (!cardFrom || !cardTo) return;

                const fieldFromElem = document.getElementById(`field_${join.fromTable}_${join.fromField}`);
                const fieldToElem = document.getElementById(`field_${join.toTable}_${join.toField}`);

                const fromAnchor = getCardAnchor(cardFrom, cardTo, fieldFromElem, true);
                const toAnchor = getCardAnchor(cardTo, cardFrom, fieldToElem, false);

                const x1 = fromAnchor.x;
                const y1 = fromAnchor.y;
                const x2 = toAnchor.x;
                const y2 = toAnchor.y;

                const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
                line.setAttribute('x1', x1);
                line.setAttribute('y1', y1);
                line.setAttribute('x2', x2);
                line.setAttribute('y2', y2);
                line.setAttribute('class', 'join-line');
                line.setAttribute('stroke-linecap', 'round');
                const joinStroke = join.type === 'INNER' ? '#434343' : '#b91c1c';
                line.setAttribute('stroke', joinStroke);
                if (join.type === 'LEFT') line.setAttribute('marker-end', 'url(#joinArrowMarker)');
                if (join.type === 'RIGHT') line.setAttribute('marker-start', 'url(#joinArrowMarker)');
                line.setAttribute('data-join-id', join.id);
                line.setAttribute('data-join-type', join.type);
                line.style.pointerEvents = 'all';
                line.style.cursor = 'pointer';
                const lineTitle = document.createElementNS('http://www.w3.org/2000/svg', 'title');
                lineTitle.textContent = join.type === 'LEFT'
                    ? `LEFT JOIN : toutes les lignes de ${join.fromTable}, flèche vers ${join.toTable}`
                    : join.type === 'RIGHT'
                        ? `RIGHT JOIN : toutes les lignes de ${join.toTable}, flèche vers ${join.fromTable}`
                        : 'INNER JOIN : uniquement les lignes correspondantes';
                line.appendChild(lineTitle);
                line.onclick = () => openJoinModal(join.id);
                line.ondblclick = (event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    openJoinModal(join.id);
                };
                line.onmouseover = () => {
                    line.setAttribute('stroke', '#005a9e');
                    line.setAttribute('stroke-width', '3');
                };
                line.onmouseout = () => {
                    line.setAttribute('stroke', joinStroke);
                    line.setAttribute('stroke-width', '2');
                };

                const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
                circle.setAttribute('cx', (x1 + x2) / 2);
                circle.setAttribute('cy', (y1 + y2) / 2);
                circle.setAttribute('r', '6');
                circle.setAttribute('fill', joinStroke);
                circle.setAttribute('cursor', 'pointer');
                circle.setAttribute('data-join-id', join.id);
                circle.onclick = () => openJoinModal(join.id);
                circle.ondblclick = (event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    openJoinModal(join.id);
                };

                svg.appendChild(line);
                svg.appendChild(circle);
            });
        }

        // Modal de Propriétés de Jointure (MS Access Style)
        function openJoinModal(joinId) {
            appState.activeJoinId = joinId;
            const join = appState.joins.find(j => j.id === joinId);
            if (!join) return;

            document.getElementById('joinModalDesc').textContent = `Jointure entre [${join.fromTable}.${join.fromField}] et [${join.toTable}.${join.toField}] :`;
            document.querySelector(`input[name="joinType"][value="${join.type}"]`).checked = true;
            document.getElementById('joinModal').classList.remove('hidden');
        }

        function closeJoinModal() {
            document.getElementById('joinModal').classList.add('hidden');
            appState.activeJoinId = null;
        }

        function saveJoinProperties() {
            if (!appState.activeJoinId) return;
            const selectedType = document.querySelector('input[name="joinType"]:checked').value;
            const join = appState.joins.find(j => j.id === appState.activeJoinId);
            if (join) {
                join.type = selectedType;
            }
            closeJoinModal();
            renderSvgJoins();
        }

        function deleteCurrentJoin() {
            if (!appState.activeJoinId) return;
            appState.joins = appState.joins.filter(j => j.id !== appState.activeJoinId);
            closeJoinModal();
            renderSvgJoins();
        }

        // ===== Grille de conception de requete =====
        // Initialisation de la grille de requête (Query Grid)
        function initQueryGrid() {
            appState.queryGridRows = Array.from({ length: 8 }, () => ({ field: '', table: '', alias: '', expression: '', total: 'GROUP BY', show: false, sort: '', criteria: '', or: '', width: 180 }));
            renderQueryGrid();
        }

        function addQueryColumn() {
            appState.queryGridRows.push({ field: '', table: '', alias: '', expression: '', total: 'GROUP BY', show: false, sort: '', criteria: '', or: '', width: 180 });
            renderQueryGrid();
            updateSaveButtonState();
            const gridScroller = document.querySelector('#queryGridPanel > .flex-1.overflow-auto');
            if (gridScroller) gridScroller.scrollLeft = gridScroller.scrollWidth;
            const fieldInput = document.querySelector(`#queryGridBody [data-row="${appState.queryGridRows.length - 1}"][data-col="field"]`);
            if (fieldInput) fieldInput.focus();
        }

        function renderQueryGrid() {
            const tbody = document.getElementById('queryGridBody');
            appState.queryGridRows.forEach(queryRow => {
                if (!queryRow.total) queryRow.total = 'GROUP BY';
            });
            tbody.innerHTML = '';
            const gridRows = [
                { label: 'Champ :', key: 'field', control: 'field' },
                { label: 'Libellé :', key: 'alias', control: 'text', placeholder: 'Libellé' },
                { label: 'GroupBy :', key: 'total', control: 'total' },
                { label: 'Tri :', key: 'sort', control: 'sort' },
                { label: 'Afficher :', key: 'show', control: 'checkbox' },
                { label: 'Critères :', key: 'criteria', control: 'text', placeholder: 'Ex. Like "*abc*" ou <> 0' },
                { label: 'Ou :', key: 'or', control: 'text', placeholder: 'Ou' }
            ];

            gridRows.forEach((gridRow, gridRowIndex) => {
                const tr = document.createElement('tr');
                const label = document.createElement('td');
                label.className = 'access-grid-header p-1 text-left whitespace-nowrap';
                label.style.width = '90px';
                label.style.minWidth = '90px';
                label.textContent = gridRow.label;
                tr.appendChild(label);

                appState.queryGridRows.forEach((queryRow, columnIndex) => {
                    const td = document.createElement('td');
                    td.className = `access-grid-cell p-1 ${gridRow.control === 'checkbox' ? 'text-center' : ''}`;
                    td.style.width = `${queryRow.width}px`;
                    td.style.minWidth = '0';
                    td.dataset.queryColumn = columnIndex;
                    td.addEventListener('dragover', allowQueryColumnDrop);
                    td.addEventListener('drop', event => dropQueryColumn(event, columnIndex));

                    if (gridRow.control === 'total') {
                        const select = document.createElement('select');
                        select.className = 'w-full bg-transparent px-1 focus:bg-white focus:outline-none text-xs';
                        select.dataset.row = columnIndex;
                        select.dataset.col = 'total';
                        select.onchange = updateQueryGrid;
                        [
                            ['', '(Aucun)'],
                            ['GROUP BY', 'Regrouper par'],
                            ['SUM', 'Somme'],
                            ['AVG', 'Moyenne'],
                            ['COUNT', 'Compte'],
                            ['MIN', 'Min'],
                            ['MAX', 'Max']
                        ].forEach(([value, text]) => {
                            const option = document.createElement('option');
                            option.value = value;
                            option.textContent = text;
                            select.appendChild(option);
                        });
                        select.value = queryRow.total || 'GROUP BY';
                        td.appendChild(select);
                    } else if (gridRow.control === 'sort') {
                        const select = document.createElement('select');
                        select.className = 'w-full bg-transparent px-1 focus:bg-white focus:outline-none text-xs';
                        select.dataset.row = columnIndex;
                        select.dataset.col = 'sort';
                        select.onchange = updateQueryGrid;
                        [['', '(Aucun)'], ['ASC', 'Croissant'], ['DESC', 'Décroissant']].forEach(([value, text]) => {
                            const option = document.createElement('option');
                            option.value = value;
                            option.textContent = text;
                            select.appendChild(option);
                        });
                        select.value = queryRow.sort;
                        td.appendChild(select);
                    } else if (gridRow.control === 'checkbox') {
                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox';
                        checkbox.checked = queryRow.show;
                        checkbox.dataset.row = columnIndex;
                        checkbox.dataset.col = 'show';
                        checkbox.onchange = updateQueryGrid;
                        td.appendChild(checkbox);
                    } else {
                        const input = gridRow.control === 'field'
                            ? document.createElement('select')
                            : document.createElement('input');
                        if (gridRow.control !== 'field') input.type = 'text';
                        input.className = gridRow.key === 'field'
                            ? 'min-w-0 flex-1 bg-transparent px-1 focus:bg-white focus:outline-none text-xs'
                            : 'w-full bg-transparent px-1 focus:bg-white focus:outline-none';
                        input.dataset.row = columnIndex;
                        input.dataset.col = gridRow.key;
                        if (gridRow.key === 'field') {
                            const emptyOption = document.createElement('option');
                            emptyOption.value = '';
                            emptyOption.textContent = '(Champ)';
                            input.appendChild(emptyOption);

                            Object.keys(appState.tables).forEach(tableName => {
                                const fieldNames = Object.keys((appState.tables[tableName] || [])[0] || {});
                                if (fieldNames.length === 0) return;
                                const group = document.createElement('optgroup');
                                group.label = tableName;
                                fieldNames.forEach(fieldName => {
                                    const option = document.createElement('option');
                                    option.value = `${tableName}.${fieldName}`;
                                    option.textContent = fieldName;
                                    group.appendChild(option);
                                });
                                input.appendChild(group);
                            });

                            const selectedValue = queryRow.table && queryRow.field
                                ? `${queryRow.table}.${queryRow.field}`
                                : '';
                            if (selectedValue && !Array.from(input.options).some(option => option.value === selectedValue)) {
                                const customOption = document.createElement('option');
                                customOption.value = selectedValue;
                                customOption.textContent = queryRow.expression || queryRow.field;
                                input.appendChild(customOption);
                            }
                            input.value = selectedValue;
                            input.title = 'Sélectionner un champ dans une table';
                        } else {
                            input.placeholder = gridRow.placeholder;
                            input.value = queryRow[gridRow.key];
                        }
                        input.onchange = updateQueryGrid;
                        if (gridRow.control !== 'field') {
                            input.oninput = event => updateQueryGrid(event, false);
                        }

                        if (gridRow.key === 'field') {
                            const controls = document.createElement('div');
                            controls.className = 'flex items-center gap-1';

                            const dragHandle = document.createElement('button');
                            dragHandle.type = 'button';
                            dragHandle.className = 'text-gray-400 hover:text-gray-700 cursor-grab px-1';
                            dragHandle.draggable = true;
                            dragHandle.title = 'Glisser pour déplacer le champ';
                            dragHandle.setAttribute('aria-label', 'Déplacer le champ');
                            dragHandle.innerHTML = '<i class="fa-solid fa-grip-vertical"></i>';
                            dragHandle.addEventListener('dragstart', event => startQueryColumnDrag(event, columnIndex));

                            const removeButton = document.createElement('button');
                            removeButton.type = 'button';
                            removeButton.className = 'shrink-0 rounded px-1 text-red-600 hover:bg-red-50 hover:text-red-800';
                            removeButton.title = 'Retirer ce champ de la requête';
                            removeButton.setAttribute('aria-label', 'Retirer ce champ');
                            removeButton.dataset.row = columnIndex;
                            removeButton.style.display = queryRow.field ? 'inline-flex' : 'none';
                            removeButton.innerHTML = '<i class="fa-solid fa-trash-can"></i>';
                            removeButton.onclick = () => removeQueryField(columnIndex);

                            const resizeHandle = document.createElement('button');
                            resizeHandle.type = 'button';
                            resizeHandle.className = 'w-2 shrink-0 self-stretch cursor-col-resize border-r-2 border-gray-300 hover:border-blue-600';
                            resizeHandle.title = 'Glisser pour ajuster la largeur';
                            resizeHandle.setAttribute('aria-label', 'Ajuster la largeur de la colonne');
                            resizeHandle.dataset.queryColumn = columnIndex;
                            resizeHandle.addEventListener('mousedown', event => startQueryColumnResize(event, columnIndex));

                            controls.append(dragHandle, input, removeButton, resizeHandle);
                            td.appendChild(controls);
                        } else {
                            td.appendChild(input);
                        }
                    }
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
            const table = tbody.closest('table');
            table.style.tableLayout = 'fixed';
            table.style.width = `${90 + appState.queryGridRows.reduce((totalWidth, queryRow) => totalWidth + queryRow.width, 0)}px`;
            syncQueryFieldHighlights();
        }

        let draggedQueryColumn = null;
        let resizingQueryColumn = null;

        function startQueryColumnResize(event, columnIndex) {
            event.preventDefault();
            resizingQueryColumn = {
                index: columnIndex,
                startX: event.clientX,
                startWidth: appState.queryGridRows[columnIndex].width
            };
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
            document.addEventListener('mousemove', resizeQueryColumn);
            document.addEventListener('mouseup', stopQueryColumnResize);
        }

        function resizeQueryColumn(event) {
            if (!resizingQueryColumn) return;
            const width = Math.max(100, resizingQueryColumn.startWidth + event.clientX - resizingQueryColumn.startX);
            const columnIndex = resizingQueryColumn.index;
            appState.queryGridRows[columnIndex].width = width;
            document.querySelectorAll(`#queryGridBody td[data-query-column="${columnIndex}"]`).forEach(cell => {
                cell.style.width = `${width}px`;
                cell.style.minWidth = '0';
            });
            const table = document.querySelector('#queryGridBody').closest('table');
            table.style.width = `${90 + appState.queryGridRows.reduce((totalWidth, queryRow) => totalWidth + queryRow.width, 0)}px`;
        }

        function stopQueryColumnResize() {
            resizingQueryColumn = null;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            document.removeEventListener('mousemove', resizeQueryColumn);
            document.removeEventListener('mouseup', stopQueryColumnResize);
        }

        function startQueryColumnDrag(event, columnIndex) {
            draggedQueryColumn = columnIndex;
            event.dataTransfer.effectAllowed = 'move';
            event.dataTransfer.setData('text/plain', String(columnIndex));
        }

        function allowQueryColumnDrop(event) {
            event.preventDefault();
            if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
        }

        function dropQueryColumn(event, targetIndex) {
            event.preventDefault();
            event.stopPropagation();
            const transferredIndex = Number(event.dataTransfer ? event.dataTransfer.getData('text/plain') : '');
            const sourceIndex = Number.isInteger(transferredIndex) && transferredIndex >= 0
                ? transferredIndex
                : draggedQueryColumn;
            draggedQueryColumn = null;
            if (sourceIndex === null || sourceIndex === targetIndex || sourceIndex >= appState.queryGridRows.length) return;

            const [movedColumn] = appState.queryGridRows.splice(sourceIndex, 1);
            appState.queryGridRows.splice(targetIndex, 0, movedColumn);
            renderQueryGrid();
        }

        function removeQueryField(columnIndex) {
            appState.queryGridRows.splice(columnIndex, 1);
            appState.queryGridRows.push({ field: '', table: '', alias: '', expression: '', total: 'GROUP BY', show: false, sort: '', criteria: '', or: '', width: 180 });
            renderQueryGrid();
        }

        function parseQueryFieldExpression(expression, currentTable, currentField) {
            const resolveReference = reference => {
                const parts = reference.trim().split('.').map(part => part.trim().replace(/^\[|\]$/g, ''));
                const fieldName = parts.pop() || '';
                const requestedTable = parts.join('.') || currentTable;
                const tableName = Object.keys(appState.tables).find(name => name.toLowerCase() === requestedTable.toLowerCase()) || requestedTable;
                const fieldNames = Object.keys((appState.tables[tableName] || [])[0] || {});
                const actualField = fieldNames.find(name => name.toLowerCase() === fieldName.toLowerCase());
                return { field: actualField || fieldName, table: tableName, exists: Boolean(actualField) };
            };

            const separatorIndex = expression.indexOf(':');
            if (separatorIndex < 0) {
                const source = resolveReference(expression);
                return { alias: '', field: source.field, table: source.table };
            }

            const left = resolveReference(expression.slice(0, separatorIndex));
            const right = resolveReference(expression.slice(separatorIndex + 1));
            const leftIsCurrentField = currentField && left.field.toLowerCase() === currentField.toLowerCase();
            const rightIsCurrentField = currentField && right.field.toLowerCase() === currentField.toLowerCase();
            const sourceIsLeft = leftIsCurrentField || (left.exists && !right.exists && !rightIsCurrentField);
            const source = sourceIsLeft ? left : right;
            const alias = sourceIsLeft ? right.field : left.field;
            return {
                alias,
                field: source.field,
                table: source.table
            };
        }

        function selectFieldForQuery(tableName, fieldName) {
            let columnIndex = appState.queryGridRows.findIndex(queryRow => !queryRow.field);
            if (columnIndex < 0) {
                appState.queryGridRows.push({ field: '', table: '', alias: '', expression: '', total: 'GROUP BY', show: false, sort: '', criteria: '', or: '', width: 180 });
                columnIndex = appState.queryGridRows.length - 1;
            }
            appState.queryGridRows[columnIndex] = {
                ...appState.queryGridRows[columnIndex],
                field: fieldName,
                table: tableName,
                alias: '',
                expression: fieldName,
                total: appState.queryGridRows[columnIndex].total || 'GROUP BY',
                show: true
            };
            renderQueryGrid();
            updateSaveButtonState();
        }

        function updateQueryGrid(e, normalizeCriteria = true) {
            const row = parseInt(e.target.dataset.row);
            const col = e.target.dataset.col;
            if (col === 'show') {
                appState.queryGridRows[row][col] = e.target.checked;
            } else {
                appState.queryGridRows[row][col] = e.target.value;
                if (col === 'field') {
                    const queryRow = appState.queryGridRows[row];
                    if (!queryRow.total) queryRow.total = 'GROUP BY';
                    const parsedExpression = parseQueryFieldExpression(e.target.value, queryRow.table, queryRow.field);
                    queryRow.expression = e.target.value;
                    queryRow.field = parsedExpression.field;
                    queryRow.alias = parsedExpression.alias;
                    if (parsedExpression.table) {
                        queryRow.table = parsedExpression.table;
                        const tableInput = document.querySelector(`#queryGridBody [data-row="${row}"][data-col="table"]`);
                        if (tableInput) tableInput.value = parsedExpression.table;
                    }
                    const hasField = Boolean(parsedExpression.field);
                    appState.queryGridRows[row].show = hasField;
                    const showCheckbox = document.querySelector(`#queryGridBody [data-row="${row}"][data-col="show"]`);
                    if (showCheckbox) showCheckbox.checked = hasField;
                    const removeButton = document.querySelector(`#queryGridBody [data-row="${row}"][aria-label="Retirer ce champ"]`);
                    if (removeButton) removeButton.style.display = hasField ? 'inline-flex' : 'none';
                }
                if (normalizeCriteria && (col === 'criteria' || col === 'or')) {
                    const normalizedCriteria = normalizeMultipleExclusions(e.target.value);
                    appState.queryGridRows[row][col] = normalizedCriteria;
                    e.target.value = normalizedCriteria;
                }
            }
            if (col === 'field' || col === 'alias') syncQueryFieldHighlights();
        }

        function syncQueryFieldHighlights() {
            document.querySelectorAll('.field-row.query-field-selected').forEach(element => {
                element.classList.remove('query-field-selected');
            });
            appState.queryGridRows.forEach(queryRow => {
                if (!queryRow.field || !queryRow.table) return;
                const fieldElement = document.getElementById(`field_${queryRow.table}_${queryRow.field}`);
                if (fieldElement) fieldElement.classList.add('query-field-selected');
            });
        }

        // ===== Construction et edition du SQL =====
        function quoteSqlIdentifier(identifier) {
            return `[${String(identifier).replace(/\]/g, ']]')}]`;
        }

        function queryFieldToSql(queryRow) {
            return queryRow.table
                ? `${quoteSqlIdentifier(queryRow.table)}.${quoteSqlIdentifier(queryRow.field)}`
                : quoteSqlIdentifier(queryRow.field);
        }

        function wildcardPatternToRegExp(pattern) {
            const sanitized = String(pattern === null || pattern === undefined ? '' : pattern)
                .replace(/[.+?^${}()|[\]\\]/g, '\\$&')
                .replace(/%/g, '.*')
                .replace(/_/g, '.')
                .replace(/\*/g, '.*');
            return new RegExp(`^${sanitized}$`, 'i');
        }

        function accessLikePattern(pattern) {
            return String(pattern === null || pattern === undefined ? '' : pattern)
                .replace(/%/g, '*')
                .replace(/_/g, '?');
        }

        function normalizeMultipleExclusions(criteria) {
            const value = String(criteria === null || criteria === undefined ? '' : criteria).trim();
            if (!value) return value;

            let values = null;
            if (/^(?:NOT|NI)\b/i.test(value)) {
                values = value
                    .replace(/^(?:NOT|NI)\s+/i, '')
                    .split(/\s+(?:AND\s+)?(?:NOT|NI)\s+/i);
            } else if (/^(?:<>|≠)\s*/.test(value)) {
                values = value.replace(/^(?:<>|≠)\s*/, '').split(/\s+(?:<>|≠)\s*/);
            }
            if (!values || values.length < 2 || values.some(item => !item.trim())) return value;
            if (values.some(item => /^(?:LIKE|IN|NOT|NI|AND|OR|OU)\b/i.test(item.trim()))) return value;
            return `NOT IN (${values.map(item => criteriaValueForSql(item.trim())).join(', ')})`;
        }

        function matchesCriterion(value, criteriaText) {
            const raw = String(criteriaText === null || criteriaText === undefined ? '' : criteriaText).trim().replace(/^\s*!=\s*/, '<> ');
            if (!raw) return true;

            const normalized = raw.replace(/\s+/g, ' ');
            const shorthandExclusions = normalizeMultipleExclusions(normalized);
            if (shorthandExclusions !== normalized) return matchesCriterion(value, shorthandExclusions);
            const alternatives = normalized.split(/\s+(?:OR|OU)\s+/i);
            if (alternatives.length > 1) {
                return alternatives.some(part => matchesCriterion(value, part));
            }
            const conditions = normalized.split(/\s+AND\s+/i);
            if (conditions.length > 1) {
                return conditions.every(part => matchesCriterion(value, part));
            }

            const stringValue = value === null || value === undefined ? '' : String(value).trim();
            const lowerValue = stringValue.toLowerCase();
            const trimmed = normalized;

            if (/^IS\s+NULL$/i.test(trimmed)) return stringValue === '';
            if (/^IS\s+NOT\s+NULL$/i.test(trimmed)) return stringValue !== '';
            if (/^(['"])\1$/.test(trimmed)) return stringValue === '';

            if (/^NOT\s+LIKE\s+/i.test(trimmed)) {
                const pattern = accessLikePattern(trimmed.replace(/^NOT\s+LIKE\s+/i, '').replace(/^['"]|['"]$/g, '').trim());
                const regex = wildcardPatternToRegExp(pattern);
                return !regex.test(stringValue);
            }

            if (/^LIKE\s+/i.test(trimmed)) {
                const pattern = accessLikePattern(trimmed.replace(/^LIKE\s+/i, '').replace(/^['"]|['"]$/g, '').trim());
                const regex = wildcardPatternToRegExp(pattern);
                return regex.test(stringValue);
            }

            const listMatch = trimmed.match(/^(NOT\s+IN|IN)\s*\((.*)\)$/i);
            if (listMatch) {
                const values = splitSqlList(listMatch[2]).map(item => item.trim().replace(/^['"]|['"]$/g, '').toLowerCase());
                const isIncluded = values.includes(lowerValue);
                return listMatch[1].toUpperCase() === 'NOT IN' ? !isIncluded : isIncluded;
            }

            if (/^NOT\s+/i.test(trimmed)) {
                const operand = trimmed.replace(/^NOT\s+/i, '').trim();
                if (/^(<=|>=|<>|=|<|>)\s*/.test(operand)) {
                    const comparison = operand.match(/^(<=|>=|<>|=|<|>)\s*(.*)$/);
                    return !compareValues(stringValue, comparison[1], comparison[2]);
                }
                if (/[\*%_]/.test(operand)) {
                    const regex = wildcardPatternToRegExp(accessLikePattern(operand));
                    return !regex.test(stringValue);
                }
                return !lowerValue.includes(operand.toLowerCase());
            }

            const comparison = trimmed.match(/^(<=|>=|<>|=|<|>)\s*(.*)$/);
            if (comparison) {
                return compareValues(stringValue, comparison[1], comparison[2]);
            }

            if (/^(IS\b|IN\b)/i.test(trimmed)) return trimmed.toUpperCase().startsWith('IS ') || trimmed.toUpperCase().startsWith('IN ');

            if (/[\*%_]/.test(trimmed)) {
                return wildcardPatternToRegExp(accessLikePattern(trimmed)).test(stringValue);
            }

            return lowerValue.includes(trimmed.toLowerCase().replace(/^['"]|['"]$/g, ''));
        }

        function compareValues(actual, operator, expected) {
            const normalizedActual = actual.trim();
            const normalizedExpected = expected.trim().replace(/^['"]|['"]$/g, '');
            const numericActual = Number(normalizedActual);
            const numericExpected = Number(normalizedExpected);
            const isNumeric = !Number.isNaN(numericActual) && !Number.isNaN(numericExpected);

            const left = isNumeric ? numericActual : normalizedActual.toLowerCase();
            const right = isNumeric ? numericExpected : normalizedExpected.toLowerCase();

            switch (operator) {
                case '=': return left === right;
                case '<>': return left !== right;
                case '>': return left > right;
                case '<': return left < right;
                case '>=': return left >= right;
                case '<=': return left <= right;
                default: return normalizedActual.toLowerCase().includes(normalizedExpected.toLowerCase());
            }
        }

        function splitSqlList(value) {
            const items = [];
            let current = '';
            let quote = '';
            for (const character of String(value === null || value === undefined ? '' : value)) {
                if (quote) {
                    current += character;
                    if (character === quote) quote = '';
                } else if (character === "'" || character === '"') {
                    quote = character;
                    current += character;
                } else if (character === ',') {
                    items.push(current);
                    current = '';
                } else {
                    current += character;
                }
            }
            items.push(current);
            return items;
        }

        function criteriaToSql(criteria) {
            const value = normalizeMultipleExclusions(criteria).replace(/^\s*!=\s*/, '<> ');
            if (!value) return '';

            const alternatives = value.split(/\s+(?:OR|OU)\s+/i);
            if (alternatives.length > 1) {
                return alternatives.map(part => criteriaToSql(part)).filter(Boolean).join(' OR ');
            }

            const comparison = value.match(/^(<=|>=|<>|=|<|>)\s*(.*)$/i);
            if (comparison) {
                const operand = comparison[2].trim();
                const quotedOperand = operand.match(/^(['"])([\s\S]*)\1$/);
                if (quotedOperand) {
                    return `${comparison[1]} "${quotedOperand[2].replace(/"/g, '""')}"`;
                }
                return `${comparison[1]} ${operand}`;
            }
            if (/^IS\s+NOT\s+NULL$/i.test(value) || /^IS\s+NULL$/i.test(value)) return value.toUpperCase();
            if (/^(['"])\1$/.test(value)) return '= ""';
            if (/^NOT\s+LIKE\s+/i.test(value)) {
                const pattern = value.replace(/^NOT\s+LIKE\s+/i, '').trim().replace(/^['"]|['"]$/g, '');
                return `NOT LIKE "${accessLikePattern(pattern).replace(/"/g, '""')}"`;
            }
            if (/^LIKE\s+/i.test(value)) {
                const pattern = value.replace(/^LIKE\s+/i, '').trim().replace(/^['"]|['"]$/g, '');
                return `LIKE "${accessLikePattern(pattern).replace(/"/g, '""')}"`;
            }
            if (/^NOT\s*=\s*/i.test(value)) return `<> ${value.replace(/^NOT\s*=\s*/, '')}`;
            if (/^NOT\s+/i.test(value)) {
                const operand = value.replace(/^NOT\s+/i, '').trim().replace(/^['"]|['"]$/g, '');
                return `NOT LIKE ${criteriaValueForSql(`*${operand}*`)}`;
            }
            if (/^(IS\b|IN\b)/i.test(value)) return value;
            return `LIKE "*${value.replace(/"/g, '""')}*"`;
        }

        function criteriaValueForSql(criteria) {
            const value = String(criteria === null || criteria === undefined ? '' : criteria).trim().replace(/^['"]|['"]$/g, '');
            return `'${value.replace(/'/g, "''")}'`;
        }

        function combineCriteriaToSql(criteriaValues) {
            const values = [];
            criteriaValues.forEach(criteria => {
                String(criteria || '').trim().replace(/^\s*!=\s*/, '<> ').split(/\s+(?:OR|OU)\s+/i)
                    .forEach(value => {
                        if (value.trim()) values.push(value.trim());
                    });
            });
            if (values.length > 1 && values.every(criteria => /^(?:<>|NOT\s*=)\s*/i.test(criteria))) {
                const excludedValues = values.map(criteria => criteria.replace(/^(?:<>|NOT\s*=)\s*/i, '').trim());
                return `NOT IN (${excludedValues.map(criteriaValueForSql).join(', ')})`;
            }
            return values.map(criteriaToSql).filter(Boolean).join(' OR ');
        }

        function getQueryPrimaryTable() {
            const validJoins = appState.joins.filter(join =>
                appState.tables[join.fromTable] && appState.tables[join.toTable]
            );
            if (validJoins.length > 0) return validJoins[0].fromTable;

            const selectedField = appState.queryGridRows.find(row =>
                row.show && row.field && row.table && appState.tables[row.table]
            );
            if (selectedField) return selectedField.table;

            const canvasTable = appState.canvasTables.find(table => appState.tables[table.name]);
            return canvasTable ? canvasTable.name : Object.keys(appState.tables)[0];
        }

        function generateQuerySql() {
            const availableTables = new Set([
                ...Object.keys(appState.tables),
                ...appState.canvasTables.map(table => table.name)
            ]);
            if (availableTables.size === 0) return '-- Aucune table dans la requête.';

            const firstJoin = appState.joins.find(join =>
                availableTables.has(join.fromTable) && availableTables.has(join.toTable)
            );
            const selectedField = appState.queryGridRows.find(row => row.show && row.field && row.table);
            const primaryTable = (firstJoin ? firstJoin.fromTable : '')
                || (selectedField && availableTables.has(selectedField.table) ? selectedField.table : '')
                || (() => {
                    const canvasTable = appState.canvasTables.find(table => availableTables.has(table.name));
                    return canvasTable ? canvasTable.name : '';
                })()
                || Object.keys(appState.tables)[0];

            const selectedFields = appState.queryGridRows.filter(row => row.show && row.field);
            const selectLines = selectedFields.length > 0
                ? selectedFields.map(row => {
                    const alias = row.alias ? ` AS ${quoteSqlIdentifier(row.alias)}` : '';
                    const aggregate = ['SUM', 'AVG', 'COUNT', 'MIN', 'MAX'].includes(row.total)
                        ? `${row.total}(${queryFieldToSql(row)})`
                        : queryFieldToSql(row);
                    return `    ${aggregate}${alias}`;
                })
                : ['    *'];
            const validJoins = appState.joins.filter(join =>
                availableTables.has(join.fromTable) && availableTables.has(join.toTable)
            );
            let fromSql = quoteSqlIdentifier(primaryTable);
            validJoins.forEach((join, index) => {
                const joinType = ['INNER', 'LEFT', 'RIGHT'].includes(join.type) ? join.type : 'INNER';
                const joinSql = `${joinType} JOIN ${quoteSqlIdentifier(join.toTable)} ON ${quoteSqlIdentifier(join.fromTable)}.${quoteSqlIdentifier(join.fromField)} = ${quoteSqlIdentifier(join.toTable)}.${quoteSqlIdentifier(join.toField)}`;
                fromSql = index === validJoins.length - 1
                    ? `${validJoins.length > 1 ? `(${fromSql})` : fromSql} ${joinSql}`
                    : `(${fromSql} ${joinSql})`;
            });
            const sqlLines = ['SELECT', selectLines.join(',\n'), `FROM ${fromSql}`];

            const criteriaGroups = appState.queryGridRows
                .filter(row => row.field && (row.criteria.trim() || row.or.trim()))
                .map(row => {
                    const field = queryFieldToSql(row);
                    const criteriaValues = [row.criteria, row.or]
                        .map(criteria => criteria.trim())
                        .filter(Boolean);
                    const combinedCriteria = combineCriteriaToSql(criteriaValues);
                    const hasAlternatives = criteriaValues.some(criteria => /\s+(?:OR|OU)\s+/i.test(criteria));
                    const criteriaParts = [];
                    criteriaValues.forEach(criteria => {
                        criteria.split(/\s+(?:OR|OU)\s+/i).forEach(part => criteriaParts.push(part));
                    });
                    return criteriaValues.length > 1 || hasAlternatives
                        ? (combinedCriteria.startsWith('NOT IN ')
                            ? `${field} ${combinedCriteria}`
                            : `(${criteriaParts
                                .map(criteria => `${field} ${criteriaToSql(criteria)}`)
                                .join(' OR ')})`)
                        : `${field} ${combinedCriteria}`;
                });
            if (criteriaGroups.length > 0) sqlLines.push(`WHERE ${criteriaGroups.join('\n  AND ')}`);

            const groupFields = selectedFields
                .filter(row => row.total === 'GROUP BY')
                .map(row => queryFieldToSql(row));
            if (groupFields.length > 0) sqlLines.push(`GROUP BY ${groupFields.join(', ')}`);

            const havingGroups = appState.queryGridRows
                .filter(row => row.field && ['SUM', 'AVG', 'COUNT', 'MIN', 'MAX'].includes(row.total) && (row.criteria.trim() || row.or.trim()))
                .map(row => {
                    const field = `${row.total}(${queryFieldToSql(row)})`;
                    const values = [row.criteria, row.or].map(criteria => criteria.trim()).filter(Boolean);
                    return values.length > 1
                        ? `(${values.map(criteria => `${field} ${criteriaToSql(criteria)}`).join(' OR ')})`
                        : `${field} ${criteriaToSql(values[0])}`;
                });
            if (havingGroups.length > 0) sqlLines.push(`HAVING ${havingGroups.join('\n  AND ')}`);

            const sortFields = appState.queryGridRows.filter(row => row.field && row.sort);
            if (sortFields.length > 0) {
                sqlLines.push(`ORDER BY ${sortFields.map(row => `${queryFieldToSql(row)} ${row.sort === 'DESC' ? 'DESC' : 'ASC'}`).join(', ')}`);
            }
            return `${sqlLines.join('\n')}\n;`;
        }

        function showSqlView() {
            if (!isNewQueryDraft && !savedQueries.some(query => query.id === activeQueryId) && !activeSheetQuery) return;
            document.getElementById('sqlQueryText').value = activeSqlText || generateQuerySql();
            const modal = document.getElementById('sqlViewModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function updateSqlEditorState() {
            activeSqlText = document.getElementById('sqlQueryText').value;
            updateSaveButtonState();
        }

        function applySqlView() {
            const sqlText = document.getElementById('sqlQueryText').value.trim();
            if (!sqlText) {
                alert('Saisissez une requête SQL avant de l’appliquer.');
                return;
            }

            let parsed;
            try {
                parsed = parseSheetQuerySql(sqlText);
            } catch (error) {
                if (isAdvancedAccessSql(sqlText)) {
                    applyAdvancedSqlOnly(sqlText);
                    closeSqlView();
                    return;
                }
                alert(`Impossible d’appliquer la requête SQL : ${error.message}`);
                return;
            }

            appState.canvasTables = parsed.canvasTables;
            appState.joins = parsed.joins;
            appState.queryGridRows = parsed.queryGridRows;
            activeSqlText = sqlText;
            isNewQueryDraft = !activeQueryId && !activeSheetQuery;
            renderQueryGrid();
            refreshNavigationPane();
            renderCanvasTables();
            renderSvgJoins();
            updateQueryDesignTabLabel();
            updateSqlButtonState();
            updateSaveButtonState();
            updateActionButtonsState();
            document.getElementById('statusBarMode').textContent = 'Requête SQL appliquée';
            closeSqlView();
        }

        function saveSqlView() {
            updateSqlEditorState();
            if (saveCurrentQuery() || !document.getElementById('queryNameModal').classList.contains('hidden')) {
                closeSqlView();
            }
        }

        function closeSqlView() {
            const modal = document.getElementById('sqlViewModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        // ===== Execution et affichage des resultats =====
        function updateQueryDesignTabLabel() {
            const queryTab = document.getElementById('queryDesignTab');
            if (!queryTab) return;
            const activeQuery = savedQueries.find(query => query.id === activeQueryId) || activeSheetQuery;
            const label = activeQuery ? activeQuery.name : 'Conception de requête';
            queryTab.innerHTML = '<i class="fa-solid fa-diagram-project mr-1"></i>';
            queryTab.appendChild(document.createTextNode(label));
            if (activeQuery) {
                const closeButton = document.createElement('span');
                closeButton.className = 'ml-2 cursor-pointer text-gray-500 hover:text-red-700';
                closeButton.title = 'Fermer la requête';
                closeButton.setAttribute('role', 'button');
                closeButton.setAttribute('aria-label', `Fermer ${activeQuery.name}`);
                closeButton.innerHTML = '<i class="fa-solid fa-xmark"></i>';
                closeButton.addEventListener('click', event => {
                    event.stopPropagation();
                    closeQueryDesignTab();
                });
                queryTab.appendChild(closeButton);
            }
        }

        function allowQueryDrop(event) {
            if (!event.dataTransfer || event.dataTransfer.types.includes('application/x-query-id')) {
                event.preventDefault();
                if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
            }
        }

        function dropQueryToWorkspace(event) {
            event.preventDefault();
            const queryId = event.dataTransfer ? event.dataTransfer.getData('application/x-query-id') : '';
            if (!queryId) return;
            loadSavedQuery(queryId);
        }

        function closeQueryDesignTab() {
            startNewQuery();
        }

        function renderWorkspaceTabs() {
            updateQueryDesignTabLabel();
            const tabsContainer = document.getElementById('tableWorkspaceTabs');
            tabsContainer.innerHTML = '';
            openTableTabs.forEach(tableName => {
                const tab = document.createElement('button');
                tab.type = 'button';
                tab.role = 'tab';
                tab.className = activeWorkspaceTab === tableName
                    ? 'border border-b-0 border-[#c8c6c4] bg-white px-3 py-1.5 text-xs font-semibold text-red-700'
                    : 'border border-transparent px-3 py-1.5 text-xs text-gray-700 hover:bg-white';
                tab.setAttribute('aria-selected', activeWorkspaceTab === tableName ? 'true' : 'false');
                tab.innerHTML = `<i class="fa-solid fa-table mr-1"></i>${tableName}`;
                tab.addEventListener('click', () => showTableTab(tableName));

                const closeButton = document.createElement('span');
                closeButton.className = 'ml-2 cursor-pointer text-gray-500 hover:text-red-700';
                closeButton.title = 'Fermer cet onglet';
                closeButton.setAttribute('role', 'button');
                closeButton.setAttribute('aria-label', `Fermer ${tableName}`);
                closeButton.innerHTML = '<i class="fa-solid fa-xmark"></i>';
                closeButton.addEventListener('click', event => {
                    event.stopPropagation();
                    closeTableTab(tableName);
                });
                tab.appendChild(closeButton);
                tabsContainer.appendChild(tab);
            });

            const queryTab = document.getElementById('queryDesignTab');
            const isQueryTabActive = activeWorkspaceTab === 'query-design';
            queryTab.setAttribute('aria-selected', isQueryTabActive ? 'true' : 'false');
            queryTab.className = isQueryTabActive
                ? 'border border-b-0 border-[#c8c6c4] bg-white px-3 py-1.5 text-xs font-semibold text-red-700'
                : 'border border-transparent px-3 py-1.5 text-xs text-gray-700 hover:bg-white';
        }

        function showQueryDesignTab() {
            activeWorkspaceTab = 'query-design';
            document.getElementById('erdCanvas').classList.remove('hidden');
            document.getElementById('queryGridPanel').classList.remove('hidden');
            document.getElementById('tableWorkspace').classList.add('hidden');
            document.getElementById('tableWorkspace').classList.remove('flex');
            renderWorkspaceTabs();
        }

        function showTableTab(tableName) {
            if (!appState.tables[tableName]) return;
            activeWorkspaceTab = tableName;
            document.getElementById('erdCanvas').classList.add('hidden');
            document.getElementById('queryGridPanel').classList.add('hidden');
            document.getElementById('tableWorkspace').classList.remove('hidden');
            document.getElementById('tableWorkspace').classList.add('flex');
            renderTableWorkspace(tableName);
            renderWorkspaceTabs();
        }

        function openTableTab(tableName) {
            if (!openTableTabs.includes(tableName)) openTableTabs.push(tableName);
            showTableTab(tableName);
        }

        function closeTableTab(tableName) {
            openTableTabs = openTableTabs.filter(openName => openName !== tableName);
            if (activeWorkspaceTab === tableName) {
                activeWorkspaceTab = openTableTabs.length > 0 ? openTableTabs[openTableTabs.length - 1] : 'query-design';
            }
            if (activeWorkspaceTab === 'query-design') showQueryDesignTab();
            else showTableTab(activeWorkspaceTab);
        }

        function openDataPreviewModal() {
            const container = document.getElementById('dataPreviewContainer');
            if (lastQueryResultColumns.length === 0) {
                document.getElementById('dataPreviewTitle').innerHTML = '<i class="fa-solid fa-table mr-2"></i> Feuille de Données (Datasheet View)';
                container.innerHTML = '<div class="text-gray-500 text-center py-8 italic">Aucune donnée à afficher.</div>';
                document.getElementById('dataPreviewCount').textContent = '0 enregistrement(s)';
            }
            document.getElementById('dataPreviewModal').classList.remove('hidden');
        }

        function openTablePreview(tableName) {
            openTableTab(tableName);
        }

        function prepareResultLayoutChange() {
            if (!activeSheetQuery) return true;
            const sourceIndex = sheetQueries.indexOf(activeSheetQuery);
            if (sourceIndex < 0) return false;
            if (!confirm(`La requête partagée « ${activeSheetQuery.name} » est en lecture seule. Créer une nouvelle requête pour conserver cette modification ?`)) return false;
            const duplicate = duplicateSheetQuery(sourceIndex);
            if (!duplicate) return false;
            loadSavedQuery(duplicate.id);
            return true;
        }

        function renderVirtualTable(container, columns, rows, getCellValue, onFilterChange = null, onColumnRename = null, onColumnMove = null, onColumnResize = null) {
            const rowHeight = 28;
            const overscan = 8;
            const getColumnTemplate = () => columns.map(column => `${Math.max(120, Number(column.width) || 180)}px`).join(' ');
            const columnFilters = columns.map(() => ({ selectedValues: null, textFilter: null }));
            const filterButtons = [];
            let sortState = { columnIndex: -1, direction: 1 };
            let closeColumnMenu = () => {};
            const viewport = document.createElement('div');
            viewport.className = 'h-full overflow-auto bg-white text-xs';

            const table = document.createElement('div');
            table.className = 'min-w-max';

            const header = document.createElement('div');
            header.className = 'sticky top-0 z-10 grid border-b border-gray-300 bg-gray-200';
            header.style.gridTemplateColumns = getColumnTemplate();
            columns.forEach((column, columnIndex) => {
                const cell = document.createElement('div');
                cell.className = 'flex cursor-grab items-center justify-between gap-1 border-r border-gray-300 p-1.5 text-gray-700 transition-colors hover:bg-gray-300';
                cell.draggable = Boolean(onColumnMove);
                cell.dataset.columnIndex = String(columnIndex);
                const title = document.createElement('div');
                title.className = 'min-w-0 flex-1 truncate cursor-text font-bold';
                title.textContent = column.alias || `${column.table ? column.table + '.' : ''}${column.field}`;
                title.title = title.textContent;
                title.addEventListener('dblclick', event => {
                    event.stopPropagation();
                    if (!onColumnRename) return;
                    if (title.querySelector('input')) return;
                    const currentTitle = title.textContent;
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.value = currentTitle;
                    input.className = 'w-full min-w-0 border border-[#8a8886] bg-white px-1 text-xs font-normal text-[#323130] focus:outline-none focus:border-[#b91c1c]';
                    input.setAttribute('aria-label', `Nouveau nom pour ${currentTitle}`);
                    replaceChildrenCompat(title, input);
                    let completed = false;
                    const finishRename = apply => {
                        if (completed) return;
                        completed = true;
                        const nextTitle = input.value.trim();
                        title.textContent = currentTitle;
                        if (apply && nextTitle && nextTitle !== currentTitle) onColumnRename(columnIndex, nextTitle);
                    };
                    input.addEventListener('keydown', keyEvent => {
                        if (keyEvent.key === 'Enter') {
                            keyEvent.preventDefault();
                            finishRename(true);
                        } else if (keyEvent.key === 'Escape') {
                            keyEvent.preventDefault();
                            finishRename(false);
                        }
                    });
                    input.addEventListener('blur', () => finishRename(true), { once: true });
                    input.focus();
                    input.select();
                });
                cell.addEventListener('dblclick', event => {
                    if (event.target === title || title.contains(event.target) || event.target === resizeHandle) return;
                    event.stopPropagation();
                    title.dispatchEvent(new MouseEvent('dblclick', { bubbles: false }));
                });
                let resizeStartX = null;
                let resizeStartWidth = null;
                const resizeHandle = document.createElement('span');
                resizeHandle.className = 'h-5 w-1 shrink-0 cursor-col-resize hover:bg-red-400';
                resizeHandle.title = 'Redimensionner la colonne';
                addPointerEventListener(resizeHandle, 'pointerdown', event => {
                    event.preventDefault();
                    event.stopPropagation();
                    resizeStartX = event.clientX;
                    resizeStartWidth = Math.max(120, Number(column.width) || 180);
                    const resize = moveEvent => {
                        column.width = Math.max(120, resizeStartWidth + moveEvent.clientX - resizeStartX);
                        header.style.gridTemplateColumns = getColumnTemplate();
                        table.querySelectorAll('.grid').forEach(grid => { grid.style.gridTemplateColumns = getColumnTemplate(); });
                    };
                    const stopResize = () => {
                        removePointerEventListener(document, 'pointermove', resize);
                        removePointerEventListener(document, 'pointerup', stopResize);
                        if (onColumnResize) onColumnResize();
                    };
                    addPointerEventListener(document, 'pointermove', resize);
                    addPointerEventListener(document, 'pointerup', stopResize);
                });
                const menuButton = document.createElement('button');
                menuButton.type = 'button';
                menuButton.className = 'shrink-0 rounded px-1.5 py-1 text-gray-600 hover:bg-gray-300 focus:outline-none focus:ring-1 focus:ring-blue-500';
                menuButton.setAttribute('aria-label', `Filtrer la colonne ${title.textContent}`);
                menuButton.setAttribute('aria-haspopup', 'true');
                menuButton.title = 'Trier et filtrer';
                menuButton.innerHTML = '<i class="fa-solid fa-filter"></i>';
                filterButtons[columnIndex] = menuButton;
                menuButton.addEventListener('click', event => {
                    event.stopPropagation();
                    closeColumnMenu();

                    const menu = document.createElement('div');
                    menu.className = 'fixed z-[90] w-[290px] overflow-hidden rounded border border-[#8a8886] bg-[#f3f2f1] p-2 text-sm text-[#323130] shadow-xl';
                    menu.setAttribute('role', 'dialog');
                    menu.setAttribute('aria-label', `Trier et filtrer ${title.textContent}`);
                    const menuButtonClass = 'block w-full px-2 py-1.5 text-left hover:bg-[#e1dfdd] disabled:cursor-not-allowed disabled:opacity-40 disabled:text-gray-500';
                    const makeMenuButton = (label, action, disabled = false) => {
                        const button = document.createElement('button');
                        button.type = 'button';
                        button.className = menuButtonClass;
                        button.textContent = label;
                        button.disabled = disabled;
                        button.addEventListener('click', action);
                        return button;
                    };
                    const hasActiveFilter = columnFilters[columnIndex].selectedValues !== null || columnFilters[columnIndex].textFilter !== null;

                    menu.appendChild(makeMenuButton('Trier de A à Z', () => {
                        sortState = { columnIndex, direction: 1 };
                        viewport.scrollTop = 0;
                        renderVisibleRows();
                        closeColumnMenu();
                    }));
                    menu.appendChild(makeMenuButton('Trier de Z à A', () => {
                        sortState = { columnIndex, direction: -1 };
                        viewport.scrollTop = 0;
                        renderVisibleRows();
                        closeColumnMenu();
                    }));
                    menu.appendChild(makeMenuButton(`Effacer le filtre dans ${title.textContent}`, () => {
                        columnFilters[columnIndex] = { selectedValues: null, textFilter: null };
                        viewport.scrollTop = 0;
                        renderVisibleRows();
                        closeColumnMenu();
                    }, !hasActiveFilter));

                    const textFilterSection = document.createElement('div');
                    textFilterSection.className = 'border-t border-[#c8c6c4] pt-1';
                    const textFilterToggle = makeMenuButton('Filtres de texte  ›', () => {
                        textFilterForm.classList.toggle('hidden');
                    });
                    textFilterSection.appendChild(textFilterToggle);
                    const textFilterForm = document.createElement('div');
                    textFilterForm.className = 'hidden space-y-2 px-2 pb-2';
                    const textOperator = document.createElement('select');
                    textOperator.className = 'w-full rounded border border-[#8a8886] bg-white px-2 py-1 text-[#323130] focus:outline-none focus:border-[#b91c1c]';
                    textOperator.setAttribute('aria-label', 'Opérateur du filtre de texte');
                    [
                        ['contains', 'Contient'],
                        ['notContains', 'Ne contient pas'],
                        ['equals', 'Égal à'],
                        ['notEquals', 'Différent de'],
                        ['startsWith', 'Commence par'],
                        ['endsWith', 'Finit par'],
                        ['isBlank', 'Est vide'],
                        ['isNotBlank', 'N’est pas vide']
                    ].forEach(([value, label]) => {
                        const option = document.createElement('option');
                        option.value = value;
                        option.textContent = label;
                        textOperator.appendChild(option);
                    });
                    const textValue = document.createElement('input');
                    textValue.type = 'search';
                    textValue.className = 'w-full rounded border border-[#8a8886] bg-white px-2 py-1 text-[#323130] focus:outline-none focus:border-[#b91c1c]';
                    textValue.placeholder = 'Valeur…';
                    textValue.setAttribute('aria-label', 'Valeur du filtre de texte');
                    const syncTextValue = () => {
                        textValue.disabled = textOperator.value === 'isBlank' || textOperator.value === 'isNotBlank';
                    };
                    textOperator.addEventListener('change', syncTextValue);
                    textFilterForm.append(textOperator, textValue);
                    const textFilterActions = document.createElement('div');
                    textFilterActions.className = 'flex justify-end gap-2';
                    const cancelTextFilter = makeMenuButton('Annuler', () => textFilterForm.classList.add('hidden'));
                    cancelTextFilter.className = 'rounded border border-[#8a8886] px-2 py-1 hover:bg-[#e1dfdd]';
                    const applyTextFilter = makeMenuButton('Appliquer', () => {
                        const operator = textOperator.value;
                        const criterion = textValue.value.trim().toLocaleLowerCase();
                        columnFilters[columnIndex].textFilter = criterion || operator === 'isBlank' || operator === 'isNotBlank'
                            ? { operator, value: criterion }
                            : null;
                        viewport.scrollTop = 0;
                        renderVisibleRows();
                        closeColumnMenu();
                    });
                    applyTextFilter.className = 'rounded bg-[#b91c1c] px-2 py-1 text-white hover:bg-[#991b1b]';
                    textFilterActions.append(cancelTextFilter, applyTextFilter);
                    textFilterForm.appendChild(textFilterActions);
                    textFilterSection.appendChild(textFilterForm);
                    menu.appendChild(textFilterSection);

                    const valuesSection = document.createElement('div');
                    valuesSection.className = 'mt-1 border-t border-[#c8c6c4] pt-2';
                    const searchValues = document.createElement('input');
                    searchValues.type = 'search';
                    searchValues.className = 'mb-2 w-full rounded border border-[#8a8886] bg-white px-2 py-1 text-[#323130] focus:outline-none focus:border-[#b91c1c]';
                    searchValues.placeholder = 'Rechercher une valeur…';
                    searchValues.setAttribute('aria-label', `Rechercher les valeurs de ${title.textContent}`);
                    valuesSection.appendChild(searchValues);
                    const valuesList = document.createElement('div');
                    valuesList.className = 'max-h-52 overflow-auto rounded border border-[#c8c6c4] bg-white p-1 text-[#323130]';
                    const uniqueValues = new Map();
                    rows.forEach((row, rowIndex) => {
                        const value = getCellValue(rowIndex, column);
                        const key = value === undefined || value === null || String(value) === '' ? '' : String(value);
                        if (!uniqueValues.has(key)) uniqueValues.set(key, key === '' ? '(Vides)' : key);
                    });
                    const sortedValues = [...uniqueValues.entries()].sort((first, second) =>
                        first[0].localeCompare(second[0], undefined, { numeric: true, sensitivity: 'base' })
                    );
                    const selectedValues = columnFilters[columnIndex].selectedValues === null
                        ? new Set(uniqueValues.keys())
                        : new Set(columnFilters[columnIndex].selectedValues);
                    const valueCheckboxes = [];
                    const updateSelectAll = () => {
                        selectAll.checked = valueCheckboxes.length > 0 && valueCheckboxes.every(item => item.checkbox.checked);
                        selectAll.indeterminate = valueCheckboxes.some(item => item.checkbox.checked) && !selectAll.checked;
                    };
                    const selectAllLabel = document.createElement('label');
                    selectAllLabel.className = 'mb-1 flex cursor-pointer items-center gap-2 border-b border-[#edebe9] px-2 py-1 font-medium';
                    const selectAll = document.createElement('input');
                    selectAll.type = 'checkbox';
                    selectAll.checked = selectedValues.size === uniqueValues.size;
                    selectAllLabel.append(selectAll, document.createTextNode('(Sélectionner tout)'));
                    valuesList.appendChild(selectAllLabel);
                    sortedValues.forEach(([key, label]) => {
                        const optionLabel = document.createElement('label');
                        optionLabel.className = 'flex cursor-pointer items-center gap-2 px-2 py-1 hover:bg-[#f3f9ff]';
                        optionLabel.dataset.filterValue = label.toLocaleLowerCase();
                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox';
                        checkbox.checked = selectedValues.has(key);
                        checkbox.addEventListener('change', updateSelectAll);
                        optionLabel.append(checkbox, document.createTextNode(label));
                        valuesList.appendChild(optionLabel);
                        valueCheckboxes.push({ key, checkbox, label: optionLabel });
                    });
                    selectAll.addEventListener('change', () => {
                        valueCheckboxes.forEach(item => { item.checkbox.checked = selectAll.checked; });
                        selectAll.indeterminate = false;
                    });
                    searchValues.addEventListener('input', () => {
                        const query = searchValues.value.toLocaleLowerCase();
                        valueCheckboxes.forEach(item => {
                            item.label.classList.toggle('hidden', !item.label.dataset.filterValue.includes(query));
                        });
                    });
                    valuesSection.appendChild(valuesList);

                    const valueActions = document.createElement('div');
                    valueActions.className = 'mt-2 flex justify-end gap-2';
                    const cancelValues = makeMenuButton('Annuler', () => closeColumnMenu());
                    cancelValues.className = 'rounded border border-[#8a8886] px-3 py-1 hover:bg-[#e1dfdd]';
                    const applyValues = makeMenuButton('OK', () => {
                        columnFilters[columnIndex].selectedValues = selectAll.checked
                            ? null
                            : new Set(valueCheckboxes.filter(item => item.checkbox.checked).map(item => item.key));
                        viewport.scrollTop = 0;
                        renderVisibleRows();
                        closeColumnMenu();
                    });
                    applyValues.className = 'rounded bg-[#b91c1c] px-4 py-1 text-white hover:bg-[#991b1b]';
                    valueActions.append(cancelValues, applyValues);
                    menu.append(valuesSection, valueActions);
                    document.body.appendChild(menu);

                    const rect = menuButton.getBoundingClientRect();
                    const menuWidth = 290;
                    menu.style.left = `${Math.max(8, Math.min(rect.left, window.innerWidth - menuWidth - 8))}px`;
                    menu.style.top = `${Math.max(8, Math.min(rect.bottom, window.innerHeight - Math.min(menu.offsetHeight, window.innerHeight - 16) - 8))}px`;
                    const handleOutsideClick = outsideEvent => {
                        if (!menu.contains(outsideEvent.target) && outsideEvent.target !== menuButton) closeColumnMenu();
                    };
                    const handleEscape = keyboardEvent => {
                        if (keyboardEvent.key === 'Escape') closeColumnMenu();
                    };
                    closeColumnMenu = () => {
                        menu.remove();
                        removePointerEventListener(document, 'pointerdown', handleOutsideClick);
                        document.removeEventListener('keydown', handleEscape);
                        closeColumnMenu = () => {};
                    };
                    addPointerEventListener(document, 'pointerdown', handleOutsideClick);
                    document.addEventListener('keydown', handleEscape);
                    if (columnFilters[columnIndex].textFilter) {
                        textOperator.value = columnFilters[columnIndex].textFilter.operator;
                        textValue.value = columnFilters[columnIndex].textFilter.value;
                        syncTextValue();
                    }
                });
                if (onColumnMove) {
                    cell.addEventListener('dragstart', event => {
                        cell.classList.add('cursor-grabbing', 'bg-blue-100', 'opacity-70');
                        event.dataTransfer.setData('text/plain', String(columnIndex));
                        event.dataTransfer.effectAllowed = 'move';
                    });
                    cell.addEventListener('dragend', () => {
                        cell.classList.remove('cursor-grabbing', 'bg-blue-100', 'opacity-70');
                        header.querySelectorAll('[data-column-index]').forEach(item => item.classList.remove('bg-yellow-100'));
                    });
                    cell.addEventListener('dragover', event => {
                        event.preventDefault();
                        event.dataTransfer.dropEffect = 'move';
                        cell.classList.add('bg-yellow-100');
                    });
                    cell.addEventListener('dragleave', () => cell.classList.remove('bg-yellow-100'));
                    cell.addEventListener('drop', event => {
                        event.preventDefault();
                        cell.classList.remove('bg-yellow-100');
                        const sourceIndex = Number(event.dataTransfer.getData('text/plain'));
                        if (Number.isInteger(sourceIndex) && sourceIndex !== columnIndex) onColumnMove(sourceIndex, columnIndex);
                    });
                }
                cell.append(title, resizeHandle, menuButton);
                header.appendChild(cell);
            });
            header.style.gridTemplateColumns = getColumnTemplate();
            table.appendChild(header);

            const body = document.createElement('div');
            body.className = 'relative';
            body.style.height = `${Math.max(rowHeight, rows.length * rowHeight)}px`;
            const visibleRows = document.createElement('div');
            visibleRows.className = 'absolute left-0 right-0';
            body.appendChild(visibleRows);
            const emptyState = document.createElement('div');
            emptyState.className = 'hidden p-4 text-center italic text-gray-500';
            emptyState.textContent = 'Aucun résultat pour ces filtres.';
            body.appendChild(emptyState);
            table.appendChild(body);
            viewport.appendChild(table);
            replaceChildrenCompat(container, viewport);

            const renderVisibleRows = () => {
                filterButtons.forEach((button, columnIndex) => {
                    const isFiltered = columnFilters[columnIndex].selectedValues !== null || columnFilters[columnIndex].textFilter !== null;
                    const isSorted = sortState.columnIndex === columnIndex;
                    button.classList.toggle('bg-blue-100', isFiltered || isSorted);
                    button.classList.toggle('text-blue-700', isFiltered || isSorted);
                    button.setAttribute('aria-pressed', String(isFiltered || isSorted));
                    button.innerHTML = isSorted
                        ? `<i class="fa-solid ${sortState.direction === 1 ? 'fa-sort-up' : 'fa-sort-down'}"></i>`
                        : '<i class="fa-solid fa-filter"></i>';
                });
                const filteredIndices = [];
                rows.forEach((row, rowIndex) => {
                    const matches = columns.every((column, columnIndex) => {
                        const filter = columnFilters[columnIndex];
                        const value = getCellValue(rowIndex, column);
                        const normalizedValue = value === undefined || value === null ? '' : String(value);
                        if (filter.selectedValues !== null && !filter.selectedValues.has(normalizedValue)) return false;
                        if (!filter.textFilter) return true;
                        const textValue = normalizedValue.toLocaleLowerCase();
                        const criterion = filter.textFilter.value;
                        switch (filter.textFilter.operator) {
                            case 'contains': return textValue.includes(criterion);
                            case 'notContains': return !textValue.includes(criterion);
                            case 'equals': return textValue === criterion;
                            case 'notEquals': return textValue !== criterion;
                            case 'startsWith': return textValue.startsWith(criterion);
                            case 'endsWith': return textValue.endsWith(criterion);
                            case 'isBlank': return textValue === '';
                            case 'isNotBlank': return textValue !== '';
                            default: return true;
                        }
                    });
                    if (matches) filteredIndices.push(rowIndex);
                });
                if (sortState.columnIndex >= 0) {
                    const column = columns[sortState.columnIndex];
                    filteredIndices.sort((firstIndex, secondIndex) => {
                        const firstValue = getCellValue(firstIndex, column);
                        const secondValue = getCellValue(secondIndex, column);
                        const firstText = firstValue === undefined || firstValue === null ? '' : String(firstValue);
                        const secondText = secondValue === undefined || secondValue === null ? '' : String(secondValue);
                        return firstText.localeCompare(secondText, undefined, { numeric: true, sensitivity: 'base' }) * sortState.direction;
                    });
                }
                body.style.height = `${Math.max(rowHeight, filteredIndices.length * rowHeight)}px`;
                emptyState.classList.toggle('hidden', filteredIndices.length > 0);
                if (onFilterChange) onFilterChange(filteredIndices);

                const firstRow = Math.max(0, Math.floor(viewport.scrollTop / rowHeight) - overscan);
                const lastRow = Math.min(filteredIndices.length, Math.ceil((viewport.scrollTop + viewport.clientHeight) / rowHeight) + overscan);
                visibleRows.style.transform = `translateY(${firstRow * rowHeight}px)`;
                replaceChildrenCompat(visibleRows);

                for (let rowIndex = firstRow; rowIndex < lastRow; rowIndex++) {
                    const sourceRowIndex = filteredIndices[rowIndex];
                    const row = document.createElement('div');
                    row.className = `grid h-7 border-b border-gray-200 ${rowIndex % 2 === 0 ? 'bg-white' : 'bg-gray-50'} hover:bg-blue-50`;
                    row.style.gridTemplateColumns = getColumnTemplate();
                    columns.forEach(column => {
                        const cell = document.createElement('div');
                        cell.className = 'overflow-hidden text-ellipsis whitespace-nowrap border-r border-gray-200 p-2 text-gray-800';
                        const value = getCellValue(sourceRowIndex, column);
                        cell.textContent = value === undefined || value === null ? '' : String(value);
                        cell.title = cell.textContent;
                        row.appendChild(cell);
                    });
                    visibleRows.appendChild(row);
                }
            };

            viewport.addEventListener('scroll', () => {
                closeColumnMenu();
                renderVisibleRows();
            }, { passive: true });
            if (typeof ResizeObserver !== 'undefined') {
                const resizeObserver = new ResizeObserver(renderVisibleRows);
                resizeObserver.observe(viewport);
            } else {
                window.addEventListener('resize', renderVisibleRows);
            }
            requestAnimationFrame(renderVisibleRows);
        }

        function renderTableWorkspace(tableName) {
            const tableRows = appState.tables[tableName];
            if (!Array.isArray(tableRows)) return;

            const columns = Object.keys(tableRows[0] || {}).map(field => ({ field, table: tableName, alias: '' }));
            lastQueryResults = tableRows;
            lastQueryResultColumns = columns;
            updateActionButtonsState();

            const container = document.getElementById('tableWorkspaceContainer');
            const title = document.getElementById('tableWorkspaceTitle');
            const exportButton = document.getElementById('exportTableWorkspaceButton');
            if (title) title.innerHTML = `<i class="fa-solid fa-table mr-2"></i> Table : ${tableName}`;
            if (exportButton) exportButton.disabled = columns.length === 0;

            if (columns.length === 0) {
                if (container) container.innerHTML = '<div class="text-gray-500 text-center py-8 italic">Cette table ne contient aucune colonne.</div>';
            } else if (container) {
                renderVirtualTable(container, columns, tableRows, (rowIndex, column) => tableRows[rowIndex][column.field]);
            }
        }

        // Exécuter et afficher les résultats de la requête (Datasheet View)
        function runQueryPreview() {
            if (activeQueryWorker) return;
            const requiredTableNames = [getQueryPrimaryTable()];
            appState.joins.forEach(join => {
                requiredTableNames.push(join.fromTable, join.toTable);
            });
            const requiredTables = new Set(requiredTableNames.filter(Boolean));
            const missingTables = [...requiredTables].filter(tableName => !appState.tables[tableName]);
            if (missingTables.length > 0) {
                showQueryExecutionError(`Importez les tables requises avant d’exécuter la requête : ${missingTables.join(', ')}.`);
                return;
            }
            const overlay = document.getElementById('queryExecutionOverlay');
            const executeButton = document.getElementById('executeQueryButton');
            const executionMessage = document.getElementById('queryExecutionMessage');
            overlay.classList.remove('hidden');
            overlay.classList.add('flex');
            executeButton.disabled = true;
            executionMessage.textContent = 'Exécution de la requête...';

            setTimeout(() => {
                const startedAt = performance.now();
                const workerSource = `
                    let appState = null;
                    const QUERY_MAX_INTERMEDIATE_ROWS = ${QUERY_MAX_INTERMEDIATE_ROWS};
                    const getQueryPrimaryTable = ${getQueryPrimaryTable.toString()};
                    const wildcardPatternToRegExp = ${wildcardPatternToRegExp.toString()};
                    const accessLikePattern = ${accessLikePattern.toString()};
                    const compareValues = ${compareValues.toString()};
                    const splitSqlList = ${splitSqlList.toString()};
                    const criteriaValueForSql = ${criteriaValueForSql.toString()};
                    const normalizeMultipleExclusions = ${normalizeMultipleExclusions.toString()};
                    const matchesCriterion = ${matchesCriterion.toString()};
                    const executeQueryEngine = ${executeQueryEngine.toString()};
                    self.onmessage = event => {
                        if (event.data.type !== 'execute') return;
                        try {
                            appState = event.data.state;
                            const results = executeQueryEngine();
                            self.postMessage({ type: 'result', results });
                        } catch (error) {
                            self.postMessage({ type: 'error', message: error.message || String(error) });
                        }
                    };
                `;
                const workerUrl = URL.createObjectURL(new Blob([workerSource], { type: 'application/javascript' }));
                activeQueryWorker = new Worker(workerUrl);

                const finish = () => {
                    clearTimeout(queryExecutionTimer);
                    queryExecutionTimer = null;
                    if (activeQueryWorker) activeQueryWorker.terminate();
                    activeQueryWorker = null;
                    URL.revokeObjectURL(workerUrl);
                    overlay.classList.add('hidden');
                    overlay.classList.remove('flex');
                    executeButton.disabled = false;
                };

                activeQueryWorker.onmessage = event => {
                    const elapsedMs = performance.now() - startedAt;
                    const message = event.data.type === 'error'
                        ? `Erreur lors de l'exécution de la requête : ${event.data.message}`
                        : null;
                    finish();
                    if (message) {
                        showQueryExecutionError(message);
                        return;
                    }
                    renderQueryResults(event.data.results);
                    document.getElementById('statusBarMode').textContent = elapsedMs >= 1000
                        ? `Requête terminée en ${(elapsedMs / 1000).toFixed(1)} s`
                        : 'Requête terminée';
                    if (elapsedMs >= 1000) {
                        document.getElementById('dataPreviewCount').textContent += ` en ${(elapsedMs / 1000).toFixed(1)} s`;
                    }
                };
                activeQueryWorker.onerror = event => {
                    finish();
                    showQueryExecutionError(`Erreur lors de l'exécution de la requête : ${event.message || 'Worker indisponible.'}`);
                };
                queryExecutionTimer = setTimeout(() => {
                    finish();
                    showQueryExecutionError('Exécution arrêtée automatiquement : durée supérieure à 30 secondes. Vérifiez les jointures ou ajoutez un critère.');
                }, QUERY_EXECUTION_TIMEOUT_MS);
                activeQueryWorker.postMessage({
                    type: 'execute',
                    state: {
                        tables: appState.tables,
                        joins: appState.joins,
                        queryGridRows: appState.queryGridRows,
                        canvasTables: appState.canvasTables
                    }
                });
            }, 0);
        }

        function cancelQueryExecution() {
            if (!activeQueryWorker) return;
            clearTimeout(queryExecutionTimer);
            queryExecutionTimer = null;
            activeQueryWorker.terminate();
            activeQueryWorker = null;
            const overlay = document.getElementById('queryExecutionOverlay');
            overlay.classList.add('hidden');
            overlay.classList.remove('flex');
            document.getElementById('executeQueryButton').disabled = false;
            document.getElementById('statusBarMode').textContent = 'Exécution arrêtée';
        }

        function showQueryExecutionError(message) {
            const container = document.getElementById('dataPreviewContainer');
            container.innerHTML = '';
            const errorMessage = document.createElement('div');
            errorMessage.className = 'mx-auto max-w-3xl border border-red-300 bg-red-50 p-4 text-red-800';
            errorMessage.textContent = message;
            container.appendChild(errorMessage);
            document.getElementById('dataPreviewCount').textContent = 'Échec de l’exécution';
            document.getElementById('statusBarMode').textContent = 'Erreur d’exécution';
            document.getElementById('dataPreviewModal').classList.remove('hidden');
        }

        function renderQueryResults(results) {
            const container = document.getElementById('dataPreviewContainer');
            container.innerHTML = '';
            document.getElementById('dataPreviewTitle').innerHTML = '<i class="fa-solid fa-table mr-2"></i> Feuille de Données (Datasheet View)';

            // Récupérer les colonnes actives cochées dans la grille
            let activeColumns = appState.queryGridRows
                .map((row, queryRowIndex) => ({ row, queryRowIndex }))
                .filter(({ row }) => row.show && row.field)
                .map(({ row, queryRowIndex }) => ({
                    field: row.field,
                    table: row.table,
                    alias: row.alias,
                    width: row.width,
                    queryRowIndex
                }));
            if (activeColumns.length === 0) {
                // Par défaut, afficher toutes les colonnes de la première table
                const firstTable = Object.keys(appState.tables)[0];
                if (firstTable) {
                    activeColumns = Object.keys(appState.tables[firstTable][0] || {}).map(f => ({ field: f, table: firstTable, width: 180, queryRowIndex: -1 }));
                }
            }
            lastQueryResults = results;
            lastQueryResultColumns = activeColumns;
            updateActionButtonsState();
            document.getElementById('exportQueryResultButton').disabled = false;

            if (results.length === 0) {
                container.innerHTML = '<div class="text-gray-500 text-center py-8 italic">Aucun enregistrement ne correspond aux critères de la requête.</div>';
                document.getElementById('dataPreviewCount').textContent = '0 enregistrement(s)';
                document.getElementById('dataPreviewModal').classList.remove('hidden');
                return;
            }

            renderVirtualTable(container, activeColumns, results, (rowIndex, column) => {
                const row = results[rowIndex];
                const key = column.table ? `${column.table}.${column.field}` : column.field;
                return row[key] !== undefined ? row[key] : row[column.field];
            }, filteredIndices => {
                lastQueryResults = filteredIndices.map(rowIndex => results[rowIndex]);
                const count = document.getElementById('dataPreviewCount');
                count.textContent = filteredIndices.length === results.length
                    ? `${results.length} enregistrement(s) trouvé(s)`
                    : `${filteredIndices.length} / ${results.length} enregistrement(s)`;
                document.getElementById('exportQueryResultButton').disabled = filteredIndices.length === 0;
            }, (columnIndex, nextTitle) => {
                if (!prepareResultLayoutChange()) return;
                const queryRowIndex = activeColumns[columnIndex].queryRowIndex;
                if (queryRowIndex < 0) return;
                appState.queryGridRows[queryRowIndex].alias = nextTitle;
                renderQueryGrid();
                updateSaveButtonState();
                renderQueryResults(results);
            }, (sourceIndex, targetIndex) => {
                if (!prepareResultLayoutChange()) return;
                if (sourceIndex === undefined) {
                    updateSaveButtonState();
                    return;
                }
                const sourceRowIndex = activeColumns[sourceIndex].queryRowIndex;
                const targetRowIndex = activeColumns[targetIndex].queryRowIndex;
                if (sourceRowIndex < 0 || targetRowIndex < 0) return;
                const [movedRow] = appState.queryGridRows.splice(sourceRowIndex, 1);
                appState.queryGridRows.splice(targetRowIndex > sourceRowIndex ? targetRowIndex - 1 : targetRowIndex, 0, movedRow);
                renderQueryGrid();
                updateSaveButtonState();
                renderQueryResults(results);
            }, () => {
                if (!prepareResultLayoutChange()) return;
                activeColumns.forEach(column => {
                    if (column.queryRowIndex >= 0) {
                        appState.queryGridRows[column.queryRowIndex].width = column.width;
                    }
                });
                updateSaveButtonState();
            });
            document.getElementById('dataPreviewModal').classList.remove('hidden');
        }

        async function exportCurrentContextToExcel() {
            showExportLoading();
            await new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 25)));
            try {
                if (activeWorkspaceTab !== 'query-design' && activeWorkspaceTab && appState.tables[activeWorkspaceTab]) {
                    await exportTableToExcel(activeWorkspaceTab);
                    return;
                }
                await exportQueryResults();
            } finally {
                hideExportLoading();
            }
        }

        function exportTableToExcel(tableName) {
            return new Promise((resolve, reject) => {
                try {
                    if (!tableName || !appState.tables[tableName] || typeof XLSX === 'undefined') {
                        resolve();
                        return;
                    }
                    const worksheet = XLSX.utils.json_to_sheet(appState.tables[tableName]);
                    const workbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(workbook, worksheet, tableName.substring(0, 31));
                    const baseName = appState.fileName.replace(/\.[^/.]+$/, '');
                    XLSX.writeFile(workbook, `${baseName}_${tableName}_Table.xlsx`);
                    resolve();
                } catch (error) {
                    reject(error);
                }
            });
        }

        function exportQueryResults() {
            return new Promise((resolve, reject) => {
                try {
                    if (typeof XLSX === 'undefined' || lastQueryResultColumns.length === 0) {
                        resolve();
                        return;
                    }

                    const exportRows = lastQueryResults.map(row => {
                        const exportedRow = {};
                        lastQueryResultColumns.forEach(column => {
                            const key = column.table ? `${column.table}.${column.field}` : column.field;
                            const heading = column.alias || key;
                            exportedRow[heading] = row[key] !== undefined ? row[key] : (row[column.field] !== undefined && row[column.field] !== null ? row[column.field] : '');
                        });
                        return exportedRow;
                    });

                    const worksheet = XLSX.utils.json_to_sheet(exportRows, {
                        header: lastQueryResultColumns.map(column => column.alias || `${column.table ? column.table + '.' : ''}${column.field}`)
                    });
                    const workbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(workbook, worksheet, 'Résultats');
                    const baseName = appState.fileName.replace(/\.[^/.]+$/, '');
                    XLSX.writeFile(workbook, `${baseName}_Resultat_Requete.xlsx`);
                    resolve();
                } catch (error) {
                    reject(error);
                }
            });
        }

        // Moteur de requête local simulant MS Access
        function executeQueryEngine() {
            if (Object.keys(appState.tables).length === 0) return [];

            const primaryTable = getQueryPrimaryTable();
            let rows = [...appState.tables[primaryTable].map(r => {
                let obj = {};
                for (let k in r) obj[`${primaryTable}.${k}`] = r[k];
                obj['__table'] = primaryTable;
                return obj;
            })];

            // Appliquer les jointures définies
            appState.joins.forEach(join => {
                const targetTableData = appState.tables[join.toTable] || [];
                const targetRowsByJoinValue = new Map();
                targetTableData.forEach(targetRow => {
                    const targetJoinValue = targetRow[join.toField];
                    const matchingRows = targetRowsByJoinValue.get(targetJoinValue);
                    if (matchingRows) matchingRows.push(targetRow);
                    else targetRowsByJoinValue.set(targetJoinValue, [targetRow]);
                });
                let newRows = [];

                rows.forEach(row => {
                    const joinVal = row[`${join.fromTable}.${join.fromField}`];
                    const matchingRows = targetRowsByJoinValue.get(joinVal) || [];

                    if (matchingRows.length > 0) {
                        matchingRows.forEach(targetRow => {
                            const combined = { ...row };
                            for (let k in targetRow) {
                                combined[`${join.toTable}.${k}`] = targetRow[k];
                            }
                            newRows.push(combined);
                        });
                    } else if (join.type === 'LEFT' || join.type === 'RIGHT') {
                        const combined = { ...row };
                        const targetColumns = Object.keys(targetTableData[0] || {});
                        targetColumns.forEach(column => {
                            combined[`${join.toTable}.${column}`] = '';
                        });
                        newRows.push(combined);
                    }
                });
                if (newRows.length > 0) rows = newRows;
                if (rows.length > QUERY_MAX_INTERMEDIATE_ROWS) {
                    throw new Error(`Trop de lignes intermédiaires (${rows.length.toLocaleString()}).`);
                }
            });

            // Appliquer les critères de la grille
            appState.queryGridRows.forEach(qRow => {
                if (!qRow.field || !qRow.criteria) return;
                const fieldKey = qRow.table ? `${qRow.table}.${qRow.field}` : qRow.field;
                const criteria = qRow.criteria.trim();

                rows = rows.filter(row => {
                    let val = row[fieldKey] !== undefined ? row[fieldKey] : row[qRow.field];
                    if (val === undefined) return false;
                    return matchesCriterion(val, criteria);
                });
            });

            // Appliquer les tris (sort)
            appState.queryGridRows.forEach(qRow => {
                if (!qRow.field || !qRow.sort) return;
                const fieldKey = qRow.table ? `${qRow.table}.${qRow.field}` : qRow.field;
                rows.sort((a, b) => {
                    let va = a[fieldKey] !== undefined ? a[fieldKey] : a[qRow.field];
                    let vb = b[fieldKey] !== undefined ? b[fieldKey] : b[qRow.field];
                    if (va < vb) return qRow.sort === 'ASC' ? -1 : 1;
                    if (va > vb) return qRow.sort === 'ASC' ? 1 : -1;
                    return 0;
                });
            });

            const groupedQueryRows = appState.queryGridRows.filter(qRow => qRow.field && qRow.total);
            if (groupedQueryRows.length > 0) {
                const groupRows = groupedQueryRows.filter(qRow => qRow.total === 'GROUP BY');
                const aggregateRows = groupedQueryRows.filter(qRow => ['SUM', 'AVG', 'COUNT', 'MIN', 'MAX'].includes(qRow.total));
                const buckets = new Map();
                rows.forEach(row => {
                    const key = JSON.stringify(groupRows.map(qRow => {
                        const fieldKey = qRow.table ? `${qRow.table}.${qRow.field}` : qRow.field;
                        return row[fieldKey] !== undefined ? row[fieldKey] : row[qRow.field];
                    }));
                    if (!buckets.has(key)) buckets.set(key, []);
                    buckets.get(key).push(row);
                });
                rows = Array.from(buckets.values()).map(bucket => {
                    const result = { ...bucket[0] };
                    aggregateRows.forEach(qRow => {
                        const fieldKey = qRow.table ? `${qRow.table}.${qRow.field}` : qRow.field;
                        const values = bucket.map(row => row[fieldKey] !== undefined ? row[fieldKey] : row[qRow.field])
                            .filter(value => value !== null && value !== undefined && value !== '');
                        let aggregateValue = '';
                        if (qRow.total === 'COUNT') aggregateValue = values.length;
                        else if (values.length > 0) {
                            const numericValues = values.map(Number).filter(value => !Number.isNaN(value));
                            if (qRow.total === 'SUM') aggregateValue = numericValues.reduce((sum, value) => sum + value, 0);
                            if (qRow.total === 'AVG') aggregateValue = numericValues.length > 0 ? numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length : '';
                            if (qRow.total === 'MIN') aggregateValue = values.reduce((min, value) => value < min ? value : min);
                            if (qRow.total === 'MAX') aggregateValue = values.reduce((max, value) => value > max ? value : max);
                        }
                        result[fieldKey] = aggregateValue;
                    });
                    return result;
                });
            }

            let groupingColumns = appState.queryGridRows
                .filter(qRow => qRow.show && qRow.field)
                .map(qRow => ({ field: qRow.field, table: qRow.table }));
            if (groupingColumns.length === 0) {
                groupingColumns = Object.keys(appState.tables[primaryTable][0] || {})
                    .map(field => ({ field, table: primaryTable }));
            }

            const groupedRows = new Map();
            rows.forEach(row => {
                const groupKey = JSON.stringify(groupingColumns.map(column => {
                    const fieldKey = column.table ? `${column.table}.${column.field}` : column.field;
                    return row[fieldKey] !== undefined ? row[fieldKey] : (row[column.field] !== undefined ? row[column.field] : null);
                }));
                if (!groupedRows.has(groupKey)) groupedRows.set(groupKey, row);
            });

            return Array.from(groupedRows.values());
        }

        // ===== Exportation des donnees et du projet =====
        // Exporter vers Excel (.xlsx) multi-onglets via SheetJS
        function executeAndExportExcel() {
            const tableNames = Object.keys(appState.tables);
            if (tableNames.length === 0) {
                alert("Aucune table à exporter.");
                return;
            }

            const wb = XLSX.utils.book_new();

            tableNames.forEach(tName => {
                const data = appState.tables[tName];
                const ws = XLSX.utils.json_to_sheet(data);
                // Nom d'onglet limité à 31 caractères (contrainte Excel)
                const sheetName = tName.substring(0, 31);
                XLSX.utils.book_append_sheet(wb, ws, sheetName);
            });

            const outputFilename = appState.fileName.replace(/\.[^/.]+$/, "") + "_ExportAccess.xlsx";
            XLSX.writeFile(wb, outputFilename);
        }

        // Export du projet en JSON
        function exportJSONProject() {
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(appState, null, 2));
            const dlAnchor = document.createElement('a');
            dlAnchor.setAttribute("href", dataStr);
            dlAnchor.setAttribute("download", appState.fileName + "_project.json");
            document.body.appendChild(dlAnchor);
            dlAnchor.click();
            dlAnchor.remove();
        }

        // ===== Fenetres, barre d'etat et cycle de fermeture =====
        function closeDataPreviewModal() {
            document.getElementById('dataPreviewModal').classList.add('hidden');
        }

        function updateStatusBar() {
            const tableCount = Object.keys(appState.tables).length;
            let totalRows = 0;
            for (let t in appState.tables) {
                totalRows += appState.tables[t].length;
            }
            document.getElementById('statusBarTablesCount').textContent = `${tableCount} table(s) chargée(s)`;
            document.getElementById('statusBarRowsCount').textContent = `${totalRows} ligne(s) au total`;
            document.getElementById('statusBarMode').textContent = 'Prêt';
        }

        function minimizeWindow() { alert("Fenêtre minimisée"); }
        function maximizeWindow() {
            const element = document.documentElement;
            const requestFullscreen = element.requestFullscreen || element.webkitRequestFullscreen || element.msRequestFullscreen;
            if (requestFullscreen) requestFullscreen.call(element);
        }
        function closeWindow() {
            if (getQueryDefinition() !== savedQueryDefinition) {
                if (confirm('Enregistrer les modifications avant de quitter ?')) {
                    if (!savedQueries.some(query => query.id === activeQueryId)) {
                        pendingCloseAfterSave = true;
                        openQueryNameModal();
                        return;
                    }
                    if (!saveCurrentQuery()) return;
                } else if (!confirm('Quitter sans enregistrer les modifications ?')) {
                    return;
                } else {
                    skipUnloadWarning = true;
                }
            }
            window.close();
        }
