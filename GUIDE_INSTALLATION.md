# 📦 Guide d'Installation - Module 7 VBA Excel

> Comment importer et utiliser les fichiers pratiques du cours

---

## 📁 Fichiers Fournis

Vous avez téléchargé les fichiers suivants :

| Fichier | Type | Description |
|---------|------|-------------|
| **Module-7-Procedures-Parametres.md** | 📄 Cours | Support théorique complet |
| **Module_Exemples.bas** | 💻 VBA | Code des 4 exemples du cours |
| **Module_Exercices_Solutions.bas** | 💻 VBA | Solutions des 3 exercices + mini-projet |
| **Donnees_Test_Remises.csv** | 📊 Données | Données pour exercice 2 (remises) |
| **Donnees_Test_Catalogue.csv** | 📊 Données | Données pour mini-projet (catalogue) |

---

## 🚀 Installation en 5 Étapes

### Étape 1 : Créer un Nouveau Classeur Excel

1. Ouvrir **Excel 2021/2024/Microsoft 365**
2. Créer un nouveau classeur vierge
3. Enregistrer sous le nom : **`Module-7-Procedures-Parametres.xlsm`**
   - ⚠️ **Important** : Choisir le format **`.xlsm`** (Excel avec macros)
   - Emplacement : Dossier de travail accessible

---

### Étape 2 : Activer l'Onglet Développeur

**Si l'onglet "Développeur" est déjà visible** → Passer à l'étape 3

**Sinon** :
1. Cliquer sur **Fichier > Options**
2. Sélectionner **Personnaliser le ruban** (menu de gauche)
3. Dans la colonne de droite, **cocher "Développeur"**
4. Cliquer sur **OK**

➡️ L'onglet **Développeur** apparaît maintenant dans le ruban

---

### Étape 3 : Importer les Modules VBA

#### 3.1 Ouvrir l'Éditeur VBA
- Appuyer sur **Alt + F11**
- Ou aller dans **Développeur > Visual Basic**

#### 3.2 Importer le Module des Exemples
1. Dans l'éditeur VBA : **Fichier > Importer un fichier...**
2. Naviguer jusqu'au fichier **`Module_Exemples.bas`**
3. Cliquer sur **Ouvrir**
4. ✅ Le module apparaît dans l'arborescence sous "Modules"

#### 3.3 Importer le Module des Exercices
1. Répéter l'opération : **Fichier > Importer un fichier...**
2. Sélectionner **`Module_Exercices_Solutions.bas`**
3. Cliquer sur **Ouvrir**
4. ✅ Vous avez maintenant 2 modules

**Résultat attendu dans l'Explorateur de projet** :
```
VBAProject (Module-7-Procedures-Parametres.xlsm)
├── Microsoft Excel Objets
│   ├── Feuil1 (Feuil1)
│   └── ThisWorkbook
└── Modules
    ├── Module_Exemples
    └── Module_Exercices_Solutions
```

---

### Étape 4 : Importer les Données de Test

#### 4.1 Données pour l'Exercice 2 (Remises)

1. Créer une nouvelle feuille : **Clic droit > Insérer > Feuille de calcul**
2. Renommer en **`Test_Remises`**
3. Aller dans **Données > Obtenir des données > À partir d'un fichier > À partir d'un fichier texte/CSV**
4. Sélectionner **`Donnees_Test_Remises.csv`**
5. Vérifier l'aperçu, cliquer sur **Charger**
6. Les données apparaissent en colonne A

**Alternative rapide** : Copier-coller manuel
```
A1: Montant HT
A2: 50
A3: 250
A4: 750
A5: 1500
A6: 75
A7: 125
A8: 450
A9: 599
A10: 999
A11: 1250
```

#### 4.2 Données pour le Mini-Projet (Catalogue)

1. Créer une nouvelle feuille : **`Catalogue`**
2. Importer **`Donnees_Test_Catalogue.csv`** (même méthode)
3. Ou copier-coller :

| Catégorie | Nom | Prix HT |
|-----------|-----|---------|
| Alimentaire | pâtes bio | 2.5 |
| Hygiène | SAVON liquide | 3.8 |
| Électronique | souris sans fil | 15 |
| Alimentaire | huile d'olive | 8.5 |
| Mobilier | chaise@bureau | 120 |
| Alimentaire | café en grains | 12.9 |
| Hygiène | dentifrice blancheur | 4.2 |
| Électronique | clavier mécanique | 89 |
| Alimentaire | miel bio | 9.8 |
| Hygiène | shampoing doux | 6.5 |

---

### Étape 5 : Tester l'Installation

#### 5.1 Test Rapide des Exemples

1. Dans Excel, appuyer sur **Alt + F8** (ou **Développeur > Macros**)
2. Sélectionner **`TesterTousLesExemples`**
3. Cliquer sur **Exécuter**
4. ✅ Des messages s'affichent, la feuille se remplit

**Vérifier** :
- Fenêtre Exécution (**Ctrl + G** dans VBE) : traces de Debug.Print
- Feuil1 : calculs de TVA affichés

#### 5.2 Test des Exercices

1. **Alt + F8** → Sélectionner **`TesterTousLesExercices`**
2. Cliquer sur **Exécuter**
3. ✅ Plusieurs feuilles sont créées/remplies :
   - `Test_Remises` : calculs de remises
   - `Catalogue` : traitement complet du catalogue

---

## 🎯 Utiliser les Modules

### Module_Exemples : 4 Procédures de Test

| Procédure | Description | Comment lancer |
|-----------|-------------|----------------|
| `TestAfficherMessage` | Affiche des messages personnalisés | Alt+F8 → Exécuter |
| `TestCalculerTVA` | Calcule TVA avec différents taux | Alt+F8 → Exécuter |
| `TestByRefByVal` | Démontre différence ByRef/ByVal | Alt+F8 → Exécuter |
| `TestFonctionsVBA` | Teste fonctions String/Date | Alt+F8 → Exécuter |
| **`TesterTousLesExemples`** | **Lance tous les tests** | **Alt+F8 → Exécuter** |

**Astuce** : Ouvrir la **Fenêtre Exécution** (Ctrl+G dans VBE) pour voir les traces `Debug.Print`

---

### Module_Exercices_Solutions : Solutions Complètes

#### Exercice 1 : Validation Email

```vba
' Tester la fonction
Sub Test()
    Debug.Print ValiderEmail("user@example.com")  ' True
    Debug.Print ValiderEmail("invalide")          ' False
End Sub
```

**Ou lancer** : Alt+F8 → `TestValiderEmail`

#### Exercice 2 : Calcul de Remises

**Lancer** : Alt+F8 → `TestCalculRemise`
- Utilise automatiquement la feuille `Test_Remises`
- Remplit les colonnes B, C, D

#### Exercice 3 : Génération de Références

**Lancer** : Alt+F8 → `TestGenererReference`
- Affiche 10 exemples de références dans la fenêtre Exécution

#### Mini-Projet : Catalogue Produits

**Lancer** : Alt+F8 → `TestCatalogue`
- Crée/nettoie la feuille `Catalogue`
- Remplit les données de test
- Traite tout le catalogue (colonnes D à H)

---

## 🔧 Débogage Pas-à-Pas

### Comment Explorer le Code en Détail

1. **Ouvrir VBE** : Alt+F11
2. **Trouver une procédure** : Double-cliquer sur un module → Chercher une fonction
3. **Poser un point d'arrêt** :
   - Cliquer dans la marge gauche (point rouge)
   - Ou curseur sur la ligne + **F9**
4. **Exécuter en mode debug** : **F8** (pas-à-pas)
5. **Voir les valeurs** : Passer la souris sur les variables

**Exemple** : Déboguer `CalculerTVA`
1. Ouvrir `Module_Exemples`
2. Chercher la fonction `CalculerTVA`
3. F9 sur la ligne `If montantHT < 0 Then`
4. Lancer `TestCalculerTVA` (F5 ou Alt+F8)
5. Le code s'arrête au point d'arrêt
6. F8 pour avancer ligne par ligne

---

## 📚 Pratiquer les Exercices

### Méthode Recommandée (Apprentissage Actif)

1. **Lire le cours** : `Module-7-Procedures-Parametres.md`
2. **Pour chaque exercice** :
   - Créer un nouveau module : `Module_MonCode`
   - Essayer de coder sans regarder la solution
   - Tester avec les données fournies
   - Comparer avec la solution dans `Module_Exercices_Solutions`
   - Refactoriser si besoin

3. **Déboguer** :
   - F8 pour exécuter pas-à-pas
   - Ctrl+G pour voir les traces Debug.Print
   - F9 pour poser des points d'arrêt

---

## ⚠️ Résolution de Problèmes

### Erreur : "Les macros ont été désactivées"

**Cause** : Paramètres de sécurité Excel

**Solution 1** : Activer pour cette session
1. Barre jaune en haut du classeur
2. Cliquer sur **Activer le contenu**

**Solution 2** : Emplacement approuvé (recommandé)
1. **Fichier > Options > Centre de gestion de la confidentialité**
2. **Paramètres du Centre de gestion de la confidentialité**
3. **Emplacements approuvés**
4. **Ajouter un nouvel emplacement**
5. Sélectionner le dossier contenant vos fichiers Excel
6. Cocher **Les sous-dossiers de cet emplacement sont également approuvés**
7. OK

---

### Erreur : "Projet ou bibliothèque introuvable"

**Cause** : Référence manquante

**Solution** :
1. Dans VBE : **Outils > Références**
2. Décocher toutes les références marquées "MANQUANT"
3. OK

---

### Erreur : "Variable non définie"

**Cause** : `Option Explicit` force la déclaration des variables

**Solution** : Déclarer toutes les variables en début de procédure
```vba
Dim maVariable As String
```

---

### Les données CSV ne s'importent pas correctement

**Cause** : Séparateur régional (virgule vs point-virgule)

**Solution** :
1. Ouvrir le CSV dans un éditeur de texte (Notepad++)
2. Vérifier le séparateur utilisé (`,` ou `;`)
3. Dans Excel : **Données > Obtenir des données > À partir d'un fichier texte/CSV**
4. Cliquer sur **Transformer les données**
5. Ajuster le délimiteur si nécessaire

---

## 🎓 Parcours Recommandé

### Niveau Débutant

1. ✅ Lire le cours théorique (sections 1 à 4)
2. ✅ Exécuter `TesterTousLesExemples`
3. ✅ Lire le code des exemples dans VBE
4. ✅ Modifier légèrement les exemples (changer valeurs)
5. ✅ Tenter l'Exercice 1 (validation email)

### Niveau Intermédiaire

1. ✅ Faire les Exercices 1, 2 et 3 sans regarder les solutions
2. ✅ Comparer avec les solutions
3. ✅ Déboguer pas-à-pas avec F8
4. ✅ Répondre au QCM du cours (objectif 80%)

### Niveau Avancé

1. ✅ Réaliser le mini-projet en 2h chrono
2. ✅ Atteindre 70/100 selon la grille d'évaluation
3. ✅ Refactoriser : optimiser le code, gérer plus d'erreurs
4. ✅ Ajouter des fonctionnalités (export CSV, UserForm, etc.)

---

## 📞 Besoin d'Aide ?

### Ressources Documentaires
- 📖 **Cours complet** : `Module-7-Procedures-Parametres.md`
- 📖 **Microsoft Learn** : [Documentation VBA officielle](https://learn.microsoft.com/fr-fr/office/vba/api/overview/excel)
- 🌐 **Excel-Pratique** : [Forums FR](https://www.excel-pratique.com/fr/vba)

### Communautés
- 💬 **Stack Overflow** : [Tag VBA](https://stackoverflow.com/questions/tagged/vba)
- 💬 **Reddit** : [r/vba](https://www.reddit.com/r/vba/)

---

## ✅ Checklist d'Installation Complète

Vérifiez que tout fonctionne :

- [ ] Classeur `.xlsm` créé et enregistré
- [ ] Onglet Développeur visible
- [ ] 2 modules VBA importés (Exemples + Exercices)
- [ ] Feuille `Test_Remises` avec données
- [ ] Feuille `Catalogue` avec données
- [ ] Test `TesterTousLesExemples` → ✅ Succès
- [ ] Test `TesterTousLesExercices` → ✅ Succès
- [ ] Fenêtre Exécution (Ctrl+G) affiche des traces
- [ ] Débogage pas-à-pas (F8) fonctionne

**🎉 Si tous les points sont cochés : Installation réussie !**

---

## 📌 Raccourcis Clavier Essentiels

| Raccourci | Action |
|-----------|--------|
| **Alt + F11** | Ouvrir/Fermer VBE |
| **Alt + F8** | Liste des macros |
| **F5** | Exécuter la procédure courante |
| **F8** | Pas-à-pas (débogage) |
| **F9** | Point d'arrêt |
| **Ctrl + G** | Fenêtre Exécution (Debug.Print) |
| **Ctrl + Espace** | Auto-complétion |

---

**Version** : 1.0 (05/11/2025)
**Auteur** : Formation VBA Excel - TOSA & ICDL
**Durée d'installation** : 15-20 minutes

---

*Bon apprentissage ! 🚀*
