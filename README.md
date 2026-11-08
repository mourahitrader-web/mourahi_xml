# BigSi_XML

Une application web hors-ligne pour visualiser, interroger et exporter des donn\u00e9es XML sous forme de tables, avec une interface inspir\u00e9e de Microsoft Access.

## \u2705 Fonctionnalit\u00e9s

- \ud83d\udcc1 **Import de fichiers XML** : Chargez vos fichiers XML locaux pour les analyser comme des tables de base de donn\u00e9es
- \ud83d\udcb0 **Visualisation des tables** : Affichez les donn\u00e9es sous forme de tables interactives
- \ud83d\udc40 **Création de jointures** : \u00c9tablissez des relations entre les tables
- \u2563\ufe0f **Constructeur de requ\u00eates visuel** : Cr\u00e9ez des requ\u00eates sans \u00e9crire de SQL
- \ud83d\udcc8 **Export Excel** : Exportez les r\u00e9sultats vers des fichiers Excel
- \ud83d\uddd2 **Gestion des requ\u00eates** : Sauvegardez et rechargez vos requ\u00eates
- \ud83d\udcbe **Mode hors-ligne** : Fonctionne compl\u00e8tement sans connexion Internet apr\u00e8s le premier chargement

## \u26a1 Optimisations de Performance

Pour les gros fichiers XML, BigSi_XML inclut maintenant des optimisations avanc\u00e9es :

- \u26a1 **Web Workers** : Parsing en arrière-plan pour une UI r\u00e9active
- \u26a1 **Cache IndexedDB** : M\u00e9morisation des fichiers d\u00e9j\u00e0 import\u00e9s
- \u26a1 **Parsing it\u00e9ratif** : Gestion des XML tr\u00e8s imbriqu\u00e9s
- \u26a1 **Parsing parall\u00e8le** : Traitement simultan\u00e9 de plusieurs fichiers

Voir [PERFORMANCE_OPTIMIZATIONS.md](PERFORMANCE_OPTIMIZATIONS.md) pour plus de d\u00e9tails.

## \u0040 Installation

### En local

1. Clonez le d\u00e9p\u00f4t ou t\u00e9l\u00e9chargez les fichiers
2. Ouvrez `index.html` dans un navigateur web moderne
3. L'application fonctionne imm\u00e9diatement (pas de serveur requis)

### Sur un serveur web

1. Copiez tous les fichiers sur votre serveur
2. Acc\u00e9dez \u00e0 l'application via l'URL de votre serveur
3. L'application se chargera et fonctionnera normalement

## \u2728 Utilisation

### Premier pas

1. Cliquez sur **"Ajouter XML"** dans le ruban
2. S\u00e9lectionnez un ou plusieurs fichiers XML depuis votre ordinateur
3. Les tables seront automatiquement extraites et affich\u00e9es

### Cr\u00e9er une requ\u00eate

1. Passez en mode **"Avanc\u00e9"** dans le ruban
2. Cliquez sur **"Nouvelle requ\u00eate"**
3. Ajoutez des tables au canevas
4. Cr\u00e9ez des jointures entre les tables
5. Ajoutez des crit\u00e8res dans la grille
6. Cliquez sur **"Ex\u00e9cuter la requ\u00eate"**

### Exporter les r\u00e9sultats

1. Apr\u00e8s avoir ex\u00e9cut\u00e9 une requ\u00eate
2. Cliquez sur **"Exporter Excel"** dans le ruban
3. Le fichier Excel sera t\u00e9l\u00e9charg\u00e9 automatiquement

## \u2196\ufe0f Navigation

- **Onglet "Affichage"** : Mode simple pour visualiser les tables
- **Onglet "Avanc\u00e9"** : Mode complet avec toutes les fonctionnalit\u00e9s de requ\u00eetage
- **Onglet "\u00c0 propos"** : Informations sur l'application

## \u2764 Contribuer

Les contributions sont les bienvenues ! Voici comment contribuer :

1. Forker le projet
2. Cr\u00e9er une branche pour votre fonctionnalit\u00e9 (`git checkout -b feature/ma-fonctionnalite`)
3. Commiter vos modifications (`git commit -m 'Ajout de ma fonctionnalit\u00e9'`)
4. Pousser vers la branche (`git push origin feature/ma-fonctionnalite`)
5. Ouvrir une Pull Request

## \u2705 Compatibilit\u00e9

- Chrome (recommand\u00e9)
- Firefox
- Edge
- Safari (avec certaines limitations)
- Mobile : Support partiel (meilleure exp\u00e9rience sur desktop)

## \u2699\ufe0f Technologies utilis\u00e9es

- **Frontend** : HTML5, CSS3, JavaScript (ES6+)
- **UI Framework** : Tailwind CSS
- **Ic\u00f4nes** : Font Awesome
- **Export Excel** : SheetJS (xlsx)
- **Storage** : IndexedDB, LocalStorage
- **Parsing** : DOMParser, Web Workers

## \u2728 Probl\u00e8mes connus

- Les tr\u00e8s gros fichiers XML (>100 Mo) peuvent prendre du temps \u00e0 parser
- Certaines fonctionnalit\u00e9s avanc\u00e9es de SQL ne sont pas support\u00e9es
- L'interface mobile n'est pas optimis\u00e9e

## \u2122 License

Ce projet est sous license MIT. Voir le fichier [LICENSE](LICENSE) pour plus de d\u00e9tails.

## \u2709 Contact

Pour toute question ou suggestion, veuillez ouvrir une issue sur GitHub.

---

**BigSi_XML** - Une application de visualisation et requ\u00eetage XML simple et puissante.
