# Module 7 - Déclarer des Procédures avec des Paramètres
**Niveau 3 - Intermédiaire/Avancé**

> **Expert-Formateur VBA Excel** - Formation certifiante TOSA & ICDL

---

## 🎯 Objectifs Mesurables

À l'issue de ce module, l'apprenant sera capable de :
- ✅ **Déclarer et appeler** des procédures (Sub) et fonctions (Function) avec paramètres en moins de 10 minutes
- ✅ **Distinguer et utiliser** correctement ByRef et ByVal selon le contexte métier
- ✅ **Intégrer** 15+ fonctions VBA intégrées dans des procédures personnalisées
- ✅ **Créer** une bibliothèque de fonctions réutilisables avec paramètres typés

---

## 📊 Compétences TOSA Visées

| Compétence | Objectif Observable | Critère | Niveau |
|------------|---------------------|---------|--------|
| Procédures paramétrées | Créer Sub/Function avec arguments | Code sans erreur, paramètres typés | I/A |
| Passage de paramètres | Choisir ByRef vs ByVal | Justification technique correcte | I/A |
| Fonctions intégrées | Utiliser String/Math/Date VBA | 80% de réussite au QCM | I |
| Modularité du code | Découper code en fonctions | Code réutilisable, DRY | A |

---

## 📋 Pré-requis

### Connaissances Requises
- ✓ Bases VBA : déclaration de variables, types de données
- ✓ Structures conditionnelles : If/Then/Else, Select Case
- ✓ Boucles : For/Next, For Each
- ✓ Notion de Sub et Function simple (sans paramètres)

### Test de Positionnement (5 min)
**Question 1** : Quelle est la différence entre Sub et Function ?
**Question 2** : Écrivez une procédure simple qui affiche "Bonjour" dans une MsgBox.
**Question 3** : Qu'est-ce qu'une variable locale ?

➡️ **Si 3/3 correct** : Niveau confirmé, poursuivre le module
➡️ **Si < 2/3** : Revoir Module 5 (Procédures de base)

---

## 📖 Notions Clés

### 1. Procédure (Sub) vs Fonction (Function)

| Caractéristique | Sub | Function |
|-----------------|-----|----------|
| **Retourne une valeur** | ❌ Non | ✅ Oui (un seul résultat) |
| **Appel** | `Call MaProcedure(arg1)` ou `MaProcedure arg1` | `resultat = MaFonction(arg1)` |
| **Usage typique** | Actions (modification, affichage) | Calculs, transformations |
| **Exemple métier** | `GenererRapport`, `EnvoyerEmail` | `CalculerTVA`, `ValiderEmail` |

### 2. Passage de Paramètres : ByRef vs ByVal

#### ByRef (By Reference) - **Par Défaut en VBA**
- ✅ Passe **l'adresse mémoire** de la variable
- ✅ La procédure peut **modifier** la variable originale
- ⚡ **Plus rapide** pour les objets/tableaux (pas de copie)
- ⚠️ **Risque** : effets de bord si modification non intentionnelle

#### ByVal (By Value)
- ✅ Passe une **copie** de la valeur
- ✅ La procédure ne peut **pas modifier** la variable originale
- 🛡️ **Plus sûr** pour protéger les données
- ⚠️ **Coût mémoire** si données volumineuses

### 3. Fonctions VBA Intégrées

#### Catégories Principales

**Chaînes de caractères (String)**
- `Len(chaine)` : Longueur
- `UCase(chaine)` : Majuscules
- `LCase(chaine)` : Minuscules
- `Left(chaine, n)` : n premiers caractères
- `Right(chaine, n)` : n derniers caractères
- `Mid(chaine, debut, longueur)` : Extraction
- `Trim(chaine)` : Supprimer espaces début/fin
- `Replace(chaine, ancien, nouveau)` : Remplacer
- `InStr(chaine, recherche)` : Position d'une sous-chaîne

**Mathématiques**
- `Round(nombre, decimales)` : Arrondir
- `Int(nombre)` : Partie entière
- `Abs(nombre)` : Valeur absolue
- `Sqr(nombre)` : Racine carrée
- `Rnd()` : Nombre aléatoire 0-1

**Dates**
- `Date` : Date du jour
- `Now` : Date et heure actuelles
- `DateAdd(intervalle, nombre, date)` : Ajouter durée
- `DateDiff(intervalle, date1, date2)` : Différence
- `Format(date, "dd/mm/yyyy")` : Formater
- `Year(date)`, `Month(date)`, `Day(date)` : Extraire composants

**Conversion & Test**
- `CStr(valeur)`, `CInt(valeur)`, `CDbl(valeur)` : Conversion de type
- `IsNumeric(valeur)` : Tester si numérique
- `IsDate(valeur)` : Tester si date valide
- `IsEmpty(variable)` : Tester si vide
- `IsNull(valeur)` : Tester si Null

---

## 🎬 Démonstration Guidée

### Chemin UI
1. **Alt+F11** → Ouvrir l'éditeur VBE
2. **Insertion > Module** → Créer un nouveau module standard
3. Copier les exemples ci-dessous
4. **F5** ou **Alt+F8** → Exécuter les procédures de test

---

### 📘 Exemple 1 : Procédure avec Paramètres Simples

```vba
Option Explicit

'═══════════════════════════════════════════════════════════
' Procédure : AfficherMessage
' But       : Afficher un message personnalisé
' Entrées   : prenom (String), age (Integer)
' Sorties   : MsgBox avec message formaté
'═══════════════════════════════════════════════════════════

Public Sub AfficherMessage(ByVal prenom As String, ByVal age As Integer)
    Dim message As String

    ' Construction du message
    message = "Bonjour " & prenom & " !" & vbCrLf & _
              "Vous avez " & age & " ans."

    ' Affichage
    MsgBox message, vbInformation, "Message Personnalisé"
End Sub

' ─── Procédure de Test ───
Public Sub TestAfficherMessage()
    ' Appel avec arguments littéraux
    Call AfficherMessage("Marie", 28)

    ' Appel sans Call (syntaxe alternative)
    AfficherMessage "Pierre", 35

    ' Appel avec variables
    Dim nomUtilisateur As String
    Dim ageUtilisateur As Integer

    nomUtilisateur = "Sophie"
    ageUtilisateur = 42

    AfficherMessage nomUtilisateur, ageUtilisateur
End Sub
```

**🔍 Points Clés** :
- ✅ `ByVal` utilisé car on ne veut **pas modifier** les variables originales
- ✅ **Typage explicite** : `As String`, `As Integer`
- ✅ Deux syntaxes d'appel possibles : avec ou sans `Call`

---

### 📘 Exemple 2 : Fonction avec Retour de Valeur

```vba
Option Explicit

'═══════════════════════════════════════════════════════════
' Fonction  : CalculerTVA
' But       : Calculer le montant TTC à partir du HT
' Entrées   : montantHT (Double), tauxTVA (Double)
' Sortie    : Montant TTC (Double)
' Exemple   : montantTTC = CalculerTVA(100, 0.2)
'═══════════════════════════════════════════════════════════

Public Function CalculerTVA(ByVal montantHT As Double, _
                            Optional ByVal tauxTVA As Double = 0.2) As Double
    On Error GoTo ErrHandler

    ' ─── Validation des Entrées ───
    If montantHT < 0 Then
        MsgBox "Le montant HT ne peut pas être négatif", vbExclamation
        CalculerTVA = 0
        Exit Function
    End If

    If tauxTVA < 0 Or tauxTVA > 1 Then
        MsgBox "Le taux de TVA doit être entre 0 et 1", vbExclamation
        CalculerTVA = 0
        Exit Function
    End If

    ' ─── Calcul ───
    CalculerTVA = montantHT * (1 + tauxTVA)

    Exit Function

ErrHandler:
    CalculerTVA = 0
    Debug.Print "Erreur dans CalculerTVA : " & Err.Description
End Function

' ─── Procédure de Test ───
Public Sub TestCalculerTVA()
    Dim prixHT As Double
    Dim prixTTC As Double

    ' Test 1 : Avec taux par défaut (20%)
    prixHT = 100
    prixTTC = CalculerTVA(prixHT)
    Debug.Print "100€ HT = " & prixTTC & "€ TTC (TVA 20%)"

    ' Test 2 : Avec taux personnalisé (5.5%)
    prixTTC = CalculerTVA(100, 0.055)
    Debug.Print "100€ HT = " & prixTTC & "€ TTC (TVA 5.5%)"

    ' Test 3 : Affichage dans Excel
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("A1").Value = "Montant HT"
        .Range("B1").Value = "Montant TTC"
        .Range("A2").Value = 150
        .Range("B2").Formula = "=A2*1.2" ' Ou utiliser la fonction
        .Range("B2").Value = CalculerTVA(.Range("A2").Value)
    End With
End Sub
```

**🔍 Points Clés** :
- ✅ **Paramètre optionnel** : `Optional ByVal tauxTVA As Double = 0.2`
- ✅ **Validation des entrées** avant calcul
- ✅ **Gestion d'erreurs** avec `On Error GoTo`
- ✅ La fonction retourne un `Double` via son nom

---

### 📘 Exemple 3 : ByRef vs ByVal - Illustration Pratique

```vba
Option Explicit

'═══════════════════════════════════════════════════════════
' Démonstration du passage ByRef vs ByVal
'═══════════════════════════════════════════════════════════

' ─── ByRef : La procédure PEUT modifier la variable ───
Public Sub DoublerValeur_ByRef(ByRef nombre As Long)
    nombre = nombre * 2
    Debug.Print "Dans DoublerValeur_ByRef : " & nombre
End Sub

' ─── ByVal : La procédure ne PEUT PAS modifier la variable originale ───
Public Sub DoublerValeur_ByVal(ByVal nombre As Long)
    nombre = nombre * 2
    Debug.Print "Dans DoublerValeur_ByVal : " & nombre
End Sub

' ─── Procédure de Test ───
Public Sub TestByRefByVal()
    Dim monNombre As Long

    ' ═══ Test ByRef ═══
    monNombre = 10
    Debug.Print "AVANT ByRef : " & monNombre
    DoublerValeur_ByRef monNombre
    Debug.Print "APRES ByRef : " & monNombre ' ➡️ Résultat : 20 (modifié !)

    Debug.Print String(50, "-")

    ' ═══ Test ByVal ═══
    monNombre = 10
    Debug.Print "AVANT ByVal : " & monNombre
    DoublerValeur_ByVal monNombre
    Debug.Print "APRES ByVal : " & monNombre ' ➡️ Résultat : 10 (inchangé !)
End Sub
```

**📊 Résultat Attendu dans la Fenêtre Exécution (Ctrl+G)** :
```
AVANT ByRef : 10
Dans DoublerValeur_ByRef : 20
APRES ByRef : 20
--------------------------------------------------
AVANT ByVal : 10
Dans DoublerValeur_ByVal : 20
APRES ByVal : 10
```

---

### 📘 Exemple 4 : Cas Métier - Utilisation des Fonctions VBA Intégrées

```vba
Option Explicit

'═══════════════════════════════════════════════════════════
' Fonction  : NettoyerTexte
' But       : Nettoyer une chaîne (trim, majuscules, accents)
' Entrées   : texte (String)
' Sortie    : Texte nettoyé (String)
'═══════════════════════════════════════════════════════════

Public Function NettoyerTexte(ByVal texte As String) As String
    Dim resultat As String

    ' ─── Étape 1 : Supprimer espaces début/fin ───
    resultat = Trim(texte)

    ' ─── Étape 2 : Convertir en majuscules ───
    resultat = UCase(resultat)

    ' ─── Étape 3 : Remplacer les espaces multiples par un seul ───
    Do While InStr(resultat, "  ") > 0
        resultat = Replace(resultat, "  ", " ")
    Loop

    ' ─── Retour ───
    NettoyerTexte = resultat
End Function

'═══════════════════════════════════════════════════════════
' Fonction  : FormaterCodePostal
' But       : Formater un code postal français (5 chiffres)
' Entrées   : codePostal (String)
' Sortie    : Code postal formaté ou "" si invalide
'═══════════════════════════════════════════════════════════

Public Function FormaterCodePostal(ByVal codePostal As String) As String
    Dim codeNettoye As String

    ' ─── Nettoyer : supprimer espaces ───
    codeNettoye = Replace(Trim(codePostal), " ", "")

    ' ─── Valider : doit être 5 chiffres ───
    If Len(codeNettoye) <> 5 Then
        FormaterCodePostal = ""
        Exit Function
    End If

    If Not IsNumeric(codeNettoye) Then
        FormaterCodePostal = ""
        Exit Function
    End If

    ' ─── Formater : ajouter zéro devant si nécessaire ───
    FormaterCodePostal = Format(codeNettoye, "00000")
End Function

'═══════════════════════════════════════════════════════════
' Fonction  : CalculerAge
' But       : Calculer l'âge à partir de la date de naissance
' Entrées   : dateNaissance (Date)
' Sortie    : Age en années (Integer)
'═══════════════════════════════════════════════════════════

Public Function CalculerAge(ByVal dateNaissance As Date) As Integer
    Dim age As Integer

    ' ─── Calcul de base ───
    age = Year(Date) - Year(dateNaissance)

    ' ─── Ajustement si anniversaire pas encore passé cette année ───
    If Month(Date) < Month(dateNaissance) Then
        age = age - 1
    ElseIf Month(Date) = Month(dateNaissance) Then
        If Day(Date) < Day(dateNaissance) Then
            age = age - 1
        End If
    End If

    CalculerAge = age
End Function

' ─── Procédure de Test Complète ───
Public Sub TestFonctionsVBA()
    Dim texte As String
    Dim cp As String
    Dim dateNaiss As Date

    ' ═══ Test NettoyerTexte ═══
    texte = "   Jean-Pierre   DUPONT   "
    Debug.Print "Avant : [" & texte & "]"
    Debug.Print "Après : [" & NettoyerTexte(texte) & "]"

    Debug.Print String(50, "-")

    ' ═══ Test FormaterCodePostal ═══
    cp = " 75001"
    Debug.Print "CP [" & cp & "] → [" & FormaterCodePostal(cp) & "]"

    cp = "1234" ' Invalide
    Debug.Print "CP [" & cp & "] → [" & FormaterCodePostal(cp) & "]"

    Debug.Print String(50, "-")

    ' ═══ Test CalculerAge ═══
    dateNaiss = DateSerial(1990, 3, 15)
    Debug.Print "Né le " & Format(dateNaiss, "dd/mm/yyyy") & _
                " → Age : " & CalculerAge(dateNaiss) & " ans"
End Sub
```

**🔍 Fonctions VBA Utilisées** :
- `Trim()` : Supprimer espaces
- `UCase()` : Majuscules
- `InStr()` : Rechercher position
- `Replace()` : Remplacer
- `Len()` : Longueur
- `IsNumeric()` : Tester si numérique
- `Format()` : Formater
- `Year()`, `Month()`, `Day()` : Extraire composants de date
- `Date` : Date du jour
- `DateSerial()` : Créer une date

---

## ✍️ Pratique Guidée

### Exercice 1 : Créer une Fonction de Validation Email ⭐⭐

**Objectif** : L'apprenant sera capable de créer une fonction de validation d'email en 15 minutes avec critères de réussite clairs.

**Contexte** : Vous devez valider les adresses email saisies dans une feuille Excel.

**Consignes** :
1. Créer une fonction `ValiderEmail(email As String) As Boolean`
2. La fonction doit retourner `True` si l'email est valide, `False` sinon
3. Critères de validation minimale :
   - Contient exactement un `@`
   - Contient au moins un `.` après le `@`
   - Longueur minimale : 5 caractères
   - Pas d'espaces

**Critères de Réussite** :
- [ ] La fonction retourne `Boolean`
- [ ] Les 4 critères de validation sont implémentés
- [ ] Code commenté et indenté
- [ ] Procédure de test avec 5 cas (3 valides, 2 invalides)
- [ ] Utilisation de fonctions VBA intégrées : `InStr()`, `Len()`, `InStrRev()`

**Aide au Débogage** :
- F8 : Exécuter ligne par ligne
- Debug.Print : Afficher résultats intermédiaires
- Fenêtre Exécution (Ctrl+G) : Voir les traces

**💡 Indice** :
```vba
' Structure de base
Public Function ValiderEmail(ByVal email As String) As Boolean
    ' 1. Vérifier longueur
    ' 2. Vérifier présence @ (InStr)
    ' 3. Vérifier présence . après @ (InStrRev)
    ' 4. Vérifier absence d'espaces (InStr avec " ")
    ' 5. Retourner True si tous critères OK
End Function
```

---

### Exercice 2 : Procédure de Calcul de Remise ⭐⭐⭐

**Objectif** : Créer une procédure qui calcule et applique une remise selon le montant d'achat.

**Contexte** : Votre entreprise applique des remises progressives :
- < 100€ : 0%
- 100€ - 499€ : 5%
- 500€ - 999€ : 10%
- ≥ 1000€ : 15%

**Consignes** :
1. Créer une fonction `CalculerMontantRemise(montantHT As Double) As Double`
2. Créer une fonction `ObtenirTauxRemise(montantHT As Double) As Double`
3. Créer une procédure `AppliquerRemisesPlage(plageDebut As Range)`
4. La procédure doit traiter une plage de cellules (colonne A = Montants HT)
5. Écrire en colonne B : Taux de remise (%)
6. Écrire en colonne C : Montant remise (€)
7. Écrire en colonne D : Montant final (€)

**Données de Test** :
```
A1: Montant HT  |  B1: Taux  |  C1: Remise  |  D1: Final
A2: 50          |  B2: ?     |  C2: ?       |  D2: ?
A3: 250         |  B3: ?     |  C3: ?       |  D3: ?
A4: 750         |  B4: ?     |  C4: ?       |  D4: ?
A5: 1500        |  B5: ?     |  C5: ?       |  D5: ?
```

**Critères de Réussite** :
- [ ] Les 3 procédures/fonctions sont créées
- [ ] Logique de remise correcte (4 tranches)
- [ ] Paramètres typés (As Double, As Range)
- [ ] Gestion d'erreurs avec `On Error GoTo`
- [ ] Utilisation de `Round()` pour arrondir à 2 décimales
- [ ] Code modulaire (séparation calcul / application)

**Résultat Attendu** :
```
A2: 50    →  B2: 0%    C2: 0.00€      D2: 50.00€
A3: 250   →  B3: 5%    C3: 12.50€     D3: 237.50€
A4: 750   →  B4: 10%   C4: 75.00€     D4: 675.00€
A5: 1500  →  B5: 15%   C5: 225.00€    D5: 1275.00€
```

---

### Exercice 3 : Fonction de Génération de Référence ⭐⭐⭐

**Objectif** : Créer une fonction qui génère une référence unique au format standard.

**Contexte** : Vous devez générer des références produit au format : `PROD-YYYY-XXXXX`
- `PROD` : Préfixe fixe
- `YYYY` : Année en cours
- `XXXXX` : Numéro séquentiel sur 5 chiffres (avec zéros devant)

**Consignes** :
1. Créer une fonction `GenererReference(numero As Long) As String`
2. Utiliser les fonctions VBA : `Format()`, `Year()`, `Date`
3. Valider que le numéro est compris entre 1 et 99999
4. Créer une procédure de test qui génère 10 références

**Exemple** :
```vba
GenererReference(1)      → "PROD-2025-00001"
GenererReference(42)     → "PROD-2025-00042"
GenererReference(12345)  → "PROD-2025-12345"
```

**Critères de Réussite** :
- [ ] Format exact respecté (15 caractères)
- [ ] Année dynamique (pas de valeur codée en dur)
- [ ] Zéros devant le numéro (Format avec "00000")
- [ ] Validation de la plage 1-99999
- [ ] Code commenté avec en-tête

**💡 Indice** :
```vba
Public Function GenererReference(ByVal numero As Long) As String
    Dim annee As Integer
    Dim numeroFormate As String

    ' Validation
    ' Récupérer année en cours
    ' Formater le numéro sur 5 chiffres
    ' Concaténer les parties
End Function
```

---

## 📝 Évaluation Formative - QCM (20 questions)

### Section 1 : Appel de Procédures et Fonctions

**Question 1** : Quelle est la syntaxe correcte pour appeler une procédure avec paramètres ?
- A) `Call MaProcedure(arg1, arg2)`
- B) `MaProcedure arg1, arg2`
- C) `MaProcedure(arg1, arg2)`
- D) A, B et C sont corrects ✅

**Feedback** :
- ✅ **D correct** : Les trois syntaxes sont valides. `Call` est optionnel, et les parenthèses aussi si pas d'utilisation de la valeur de retour.

---

**Question 2** : Comment appeler une fonction et récupérer sa valeur ?
- A) `Call MaFonction(arg1)`
- B) `resultat = MaFonction(arg1)` ✅
- C) `MaFonction arg1`
- D) `Get MaFonction(arg1)`

**Feedback** :
- ✅ **B correct** : On récupère la valeur retournée via l'opérateur `=`
- ❌ **A et C** : Ces syntaxes ignorent la valeur de retour
- ❌ **D** : `Get` n'existe pas pour cet usage en VBA

---

**Question 3** : Quelle déclaration permet de rendre un paramètre optionnel ?
- A) `Sub Test(Optional x As Integer)`
- B) `Sub Test(x As Integer = 10)` ✅
- C) `Sub Test([x As Integer])`
- D) A et B sont corrects

**Feedback** :
- ✅ **D correct** : `Optional` + valeur par défaut sont tous deux valides
- Syntaxe complète : `Sub Test(Optional x As Integer = 10)`

---

**Question 4** : Peut-on appeler une Function sans récupérer sa valeur de retour ?
- A) Non, c'est une erreur de compilation
- B) Oui, mais c'est déconseillé ✅
- C) Oui, c'est obligatoire pour les fonctions de type Sub
- D) Non, il faut utiliser Call

**Feedback** :
- ✅ **B correct** : C'est possible mais peu logique. Si on n'utilise pas la valeur retournée, mieux vaut créer une Sub
- VBA n'empêche pas l'appel sans récupération, mais c'est une mauvaise pratique

---

### Section 2 : ByRef vs ByVal

**Question 5** : Quel est le mode de passage par défaut en VBA ?
- A) ByVal
- B) ByRef ✅
- C) ByAddress
- D) Aucun (doit être spécifié)

**Feedback** :
- ✅ **B correct** : Si non spécifié, VBA utilise **ByRef** par défaut
- ⚠️ **Attention** : Contrairement à d'autres langages comme C# (ByVal par défaut)

---

**Question 6** : Quelle affirmation est vraie pour ByVal ?
- A) Passe l'adresse mémoire de la variable
- B) La procédure peut modifier la variable originale
- C) Passe une copie de la valeur ✅
- D) Est plus rapide pour les gros tableaux

**Feedback** :
- ✅ **C correct** : ByVal crée une **copie** de la valeur
- ❌ **A et B** : Décrivent ByRef
- ❌ **D** : ByVal est plus lent pour les grosses structures (coût de copie)

---

**Question 7** : Quand utiliser ByRef ?
- A) Pour protéger les données d'origine
- B) Quand on veut modifier la variable passée ✅
- C) Toujours, c'est plus rapide
- D) Jamais, c'est dangereux

**Feedback** :
- ✅ **B correct** : ByRef permet à la procédure de modifier la variable originale
- Usage typique : retourner plusieurs valeurs via des paramètres
- ❌ **A** : C'est ByVal qui protège
- ❌ **C et D** : Dépend du contexte

---

**Question 8** : Quel code modifie la variable `x` ?
```vba
Sub Test1(ByVal n As Integer)
    n = n * 2
End Sub

Sub Test2(ByRef n As Integer)
    n = n * 2
End Sub
```
- A) Test1 uniquement
- B) Test2 uniquement ✅
- C) Les deux
- D) Aucun

**Feedback** :
- ✅ **B correct** : Seul ByRef modifie la variable originale
- Test1 modifie la **copie locale**, mais pas la variable passée en argument

---

### Section 3 : Fonctions VBA Intégrées

**Question 9** : Que retourne `Len("Bonjour")` ?
- A) 6
- B) 7 ✅
- C) 8
- D) Erreur

**Feedback** :
- ✅ **B correct** : "Bonjour" contient 7 caractères
- `Len()` compte tous les caractères, espaces inclus

---

**Question 10** : Quelle fonction extrait "VBA" de "Formation VBA Excel" ?
- A) `Mid("Formation VBA Excel", 11, 3)` ✅
- B) `Left("Formation VBA Excel", 3)`
- C) `Right("Formation VBA Excel", 3)`
- D) `Extract("Formation VBA Excel", "VBA")`

**Feedback** :
- ✅ **A correct** : `Mid(chaîne, position_départ, longueur)`
- Position 11 = début de "VBA", longueur 3
- ❌ **B** : Retourne "For"
- ❌ **C** : Retourne "cel"
- ❌ **D** : `Extract()` n'existe pas en VBA

---

**Question 11** : Comment obtenir la date du jour ?
- A) `Today()`
- B) `CurrentDate()`
- C) `Date` ✅
- D) `GetDate()`

**Feedback** :
- ✅ **C correct** : `Date` (sans parenthèses) retourne la date du jour
- `Now` retourne date + heure
- Les autres fonctions n'existent pas en VBA

---

**Question 12** : Que fait `Round(3.7456, 2)` ?
- A) Retourne 3
- B) Retourne 3.74
- C) Retourne 3.75 ✅
- D) Retourne 4

**Feedback** :
- ✅ **C correct** : Arrondit à 2 décimales → 3.75
- Syntaxe : `Round(nombre, nombre_de_décimales)`

---

**Question 13** : Comment tester si une variable est numérique ?
- A) `If IsNumber(x) Then`
- B) `If IsNumeric(x) Then` ✅
- C) `If TypeOf x Is Number Then`
- D) `If x = Number Then`

**Feedback** :
- ✅ **B correct** : `IsNumeric()` est la fonction VBA standard
- Retourne `True` si la valeur peut être convertie en nombre

---

**Question 14** : Quelle fonction convertit "hello" en "HELLO" ?
- A) `Upper("hello")`
- B) `UCase("hello")` ✅
- C) `ToUpper("hello")`
- D) `Uppercase("hello")`

**Feedback** :
- ✅ **B correct** : `UCase()` = UpperCase en VBA
- `LCase()` pour minuscules
- Les autres fonctions n'existent pas en VBA (mais dans d'autres langages)

---

**Question 15** : Comment supprimer les espaces début/fin de " test " ?
- A) `Trim(" test ")` ✅
- B) `Strip(" test ")`
- C) `Clean(" test ")`
- D) `RemoveSpaces(" test ")`

**Feedback** :
- ✅ **A correct** : `Trim()` supprime espaces début + fin → "test"
- `LTrim()` = espaces à gauche uniquement
- `RTrim()` = espaces à droite uniquement

---

### Section 4 : Cas Pratiques

**Question 16** : Quelle fonction VBA permet de chercher la position d'un caractère dans une chaîne ?
- A) `Find()`
- B) `Search()`
- C) `InStr()` ✅
- D) `IndexOf()`

**Feedback** :
- ✅ **C correct** : `InStr(chaîne, recherche)` retourne la position (1-based)
- Retourne 0 si non trouvé
- `InStrRev()` pour chercher depuis la fin

---

**Question 17** : Comment extraire l'année d'une date ?
- A) `GetYear(date)`
- B) `Year(date)` ✅
- C) `date.Year`
- D) `Format(date, "yyyy")`

**Feedback** :
- ✅ **B et D corrects** :
  - `Year(date)` retourne un Integer
  - `Format(date, "yyyy")` retourne un String
- **B est plus direct** pour un calcul

---

**Question 18** : Que retourne `Replace("Bonjour", "o", "0")` ?
- A) "Bonj0ur"
- B) "B0nj0ur" ✅
- C) "Bonjour"
- D) Erreur

**Feedback** :
- ✅ **B correct** : Remplace **tous** les "o" par "0"
- Pour remplacer une seule occurrence, ajouter paramètre count : `Replace(chaîne, ancien, nouveau, start, count)`

---

**Question 19** : Comment valider qu'une chaîne est une date ?
- A) `If IsDate(chaine) Then` ✅
- B) `If TypeOf chaine Is Date Then`
- C) `If chaine.IsDate Then`
- D) `If ValidDate(chaine) Then`

**Feedback** :
- ✅ **A correct** : `IsDate()` teste si la chaîne peut être convertie en date
- Prend en compte les paramètres régionaux (format date)

---

**Question 20** : Quelle syntaxe crée une fonction qui retourne un String ?
- A) `Sub MaFonction() As String`
- B) `Function MaFonction() As String` ✅
- C) `Function MaFonction() Returns String`
- D) `String Function MaFonction()`

**Feedback** :
- ✅ **B correct** : `Function NomFonction() As TypeRetour`
- La valeur est retournée en affectant le nom de la fonction : `MaFonction = "résultat"`

---

## 🏆 Évaluation Sommative - Mini-Projet

### Projet : Système de Gestion de Références Produits

**Contexte Professionnel** :
Vous travaillez pour une PME qui doit gérer un catalogue de produits. Votre mission est de créer un système VBA permettant de :
1. Générer des références produits automatiques
2. Valider les données saisies (nom, prix, catégorie)
3. Calculer le prix TTC selon la catégorie
4. Nettoyer et formater les données

---

**Cahier des Charges** :

#### Fonctions à Créer

**1. `GenererReferenceProduit(categorie As String, numero As Long) As String`**
- Format : `CAT-YYYY-NNNNN`
- Catégorie sur 3 lettres en majuscules
- Année sur 4 chiffres
- Numéro sur 5 chiffres avec zéros devant
- Exemple : `ALI-2025-00042` pour Alimentaire

**2. `ValiderNomProduit(nom As String) As Boolean`**
- Longueur entre 3 et 50 caractères
- Pas de caractères spéciaux (@, #, $, %, etc.)
- Retourne True si valide

**3. `CalculerPrixTTC(prixHT As Double, categorie As String) As Double`**
- Alimentaire (ALI) : TVA 5.5%
- Hygiène (HYG) : TVA 20%
- Électronique (ELE) : TVA 20%
- Retourne 0 si catégorie inconnue

**4. `FormaterNomProduit(nom As String) As String`**
- Supprimer espaces début/fin
- Premier caractère de chaque mot en majuscule
- Exemple : " ordinateur portable " → "Ordinateur Portable"

**5. `ObtenirCodeCategorie(nomCategorie As String) As String`**
- Convertir nom complet vers code 3 lettres
- "Alimentaire" → "ALI"
- "Hygiène" → "HYG"
- "Électronique" → "ELE"
- Non sensible à la casse

#### Procédure Principale

**`TraiterCatalogueProduits(feuille As Worksheet)`**
- Traiter les lignes 2 à dernière ligne remplie
- Colonne A : Catégorie (nom complet)
- Colonne B : Nom produit (à nettoyer)
- Colonne C : Prix HT
- Colonne D : À remplir → Code catégorie
- Colonne E : À remplir → Référence produit
- Colonne F : À remplir → Nom nettoyé
- Colonne G : À remplir → Prix TTC
- Colonne H : À remplir → Statut validation (OK/ERREUR + raison)

---

**Données de Test** (Feuil1) :

| A (Catégorie) | B (Nom) | C (Prix HT) |
|---------------|---------|-------------|
| Alimentaire | " pâtes bio " | 2.5 |
| Hygiène | "SAVON liquide" | 3.8 |
| Électronique | "   souris sans fil   " | 15.0 |
| Alimentaire | "huile d'olive" | 8.5 |
| Mobilier | "chaise@bureau" | 120.0 |

---

**Résultat Attendu** :

| D (Code) | E (Référence) | F (Nom Nettoyé) | G (Prix TTC) | H (Statut) |
|----------|---------------|-----------------|--------------|------------|
| ALI | ALI-2025-00001 | Pâtes Bio | 2.64 | OK |
| HYG | HYG-2025-00002 | Savon Liquide | 4.56 | OK |
| ELE | ELE-2025-00003 | Souris Sans Fil | 18.00 | OK |
| ALI | ALI-2025-00004 | Huile D'Olive | 8.97 | OK |
|  |  | chaise@bureau | 0.00 | ERREUR: Nom invalide |

---

**Critères d'Évaluation** (Total : 100 points)

| Critère | Points | Détail |
|---------|--------|--------|
| **Exactitude fonctionnelle** | 40 | Toutes les fonctions produisent les résultats attendus |
| **Qualité du code** | 20 | Option Explicit, typage, commentaires, indentation |
| **Gestion d'erreurs** | 15 | On Error GoTo, validation des entrées |
| **Utilisation fonctions VBA** | 15 | Minimum 8 fonctions intégrées différentes |
| **Modularité** | 10 | Code réutilisable, pas de duplication (DRY) |

**Seuil de Réussite** : 70/100

---

**Aide au Démarrage** :

```vba
Option Explicit

'═══════════════════════════════════════════════════════════
' Module   : GestionProduits
' But      : Système de gestion de catalogue produits
' Auteur   : [Votre Nom]
' Date     : 05/11/2025
'═══════════════════════════════════════════════════════════

' ─── Variable de module pour le compteur ───
Private compteurReference As Long

Public Function GenererReferenceProduit(ByVal categorie As String, _
                                       ByVal numero As Long) As String
    ' TODO : Implémenter
    ' Utiliser : UCase(), Year(), Date, Format()
End Function

Public Function ValiderNomProduit(ByVal nom As String) As Boolean
    ' TODO : Implémenter
    ' Utiliser : Len(), InStr()
End Function

Public Function CalculerPrixTTC(ByVal prixHT As Double, _
                               ByVal categorie As String) As Double
    ' TODO : Implémenter
    ' Utiliser : Select Case, Round()
End Function

Public Function FormaterNomProduit(ByVal nom As String) As String
    ' TODO : Implémenter
    ' Utiliser : Trim(), StrConv() avec vbProperCase
End Function

Public Function ObtenirCodeCategorie(ByVal nomCategorie As String) As String
    ' TODO : Implémenter
    ' Utiliser : UCase(), Left() ou Select Case
End Function

Public Sub TraiterCatalogueProduits(ByVal feuille As Worksheet)
    On Error GoTo ErrHandler

    Dim derniereLigne As Long
    Dim i As Long

    ' TODO : Implémenter la boucle de traitement

    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

' ─── Procédure de Test ───
Public Sub TestCatalogue()
    TraiterCatalogueProduits ThisWorkbook.Worksheets("Feuil1")
    MsgBox "Traitement terminé !", vbInformation
End Sub
```

**Temps Imparti** : 2 heures

---

## 🔄 Remédiation - Erreurs Fréquentes

### Erreur 1 : Confusion entre Sub et Function

**❌ Code Problématique** :
```vba
Sub CalculerTotal(montant As Double)
    CalculerTotal = montant * 1.2 ' ❌ Sub ne peut pas retourner de valeur !
End Sub
```

**✅ Solution** :
```vba
Function CalculerTotal(ByVal montant As Double) As Double
    CalculerTotal = montant * 1.2 ' ✅ Function retourne une valeur
End Function
```

**📚 Explication** :
- `Sub` = procédure qui **agit** (affichage, modification)
- `Function` = fonction qui **calcule** et retourne un résultat

---

### Erreur 2 : Oublier ByVal/ByRef

**❌ Code Problématique** :
```vba
' ByRef par défaut → modification involontaire !
Sub AfficherDouble(nombre As Long)
    nombre = nombre * 2
    Debug.Print nombre
End Sub

' Appel
Dim x As Long
x = 10
AfficherDouble x
Debug.Print x ' ❌ Affiche 20 au lieu de 10 !
```

**✅ Solution** :
```vba
' ByVal explicite → protège la variable
Sub AfficherDouble(ByVal nombre As Long)
    nombre = nombre * 2
    Debug.Print nombre ' Affiche 20
End Sub

' Appel
Dim x As Long
x = 10
AfficherDouble x
Debug.Print x ' ✅ Affiche toujours 10
```

**📚 Règle d'Or** :
- **Toujours spécifier** `ByVal` ou `ByRef` explicitement
- **Par défaut** : utiliser `ByVal` (sauf besoin de modification)

---

### Erreur 3 : Paramètre Optional sans Valeur par Défaut

**❌ Code Problématique** :
```vba
Function Calculer(ByVal x As Double, Optional y As Double) As Double
    Calculer = x + y ' ❌ Si y non fourni → erreur !
End Function
```

**✅ Solution 1 : Valeur par Défaut** :
```vba
Function Calculer(ByVal x As Double, Optional ByVal y As Double = 0) As Double
    Calculer = x + y ' ✅ y vaut 0 si non fourni
End Function
```

**✅ Solution 2 : Tester IsMissing** :
```vba
Function Calculer(ByVal x As Double, Optional y As Variant) As Double
    Dim valeurY As Double

    If IsMissing(y) Then
        valeurY = 0
    Else
        valeurY = CDbl(y)
    End If

    Calculer = x + valeurY
End Function
```

**📚 Note** : `IsMissing()` ne fonctionne qu'avec type `Variant`

---

### Erreur 4 : Mauvaise Utilisation des Fonctions String

**❌ Code Problématique** :
```vba
Dim nom As String
nom = " Jean Dupont "
If Mid(nom, 1, 4) = "Jean" Then ' ❌ Faux à cause des espaces !
    Debug.Print "Trouvé"
End If
```

**✅ Solution** :
```vba
Dim nom As String
nom = " Jean Dupont "
nom = Trim(nom) ' ✅ Nettoyer d'abord
If Left(nom, 4) = "Jean" Then ' ✅ ou Mid(nom, 1, 4)
    Debug.Print "Trouvé"
End If
```

**📚 Checklist Manipulation String** :
1. **Toujours** `Trim()` avant comparaison
2. **Penser** à la casse : `UCase()` ou `LCase()` pour comparaison insensible
3. **Valider** la longueur avec `Len()` avant `Mid()`/`Left()`/`Right()`

---

### Erreur 5 : Ne Pas Valider les Paramètres

**❌ Code Problématique** :
```vba
Function DiviserNombres(ByVal a As Double, ByVal b As Double) As Double
    DiviserNombres = a / b ' ❌ Division par zéro possible !
End Function
```

**✅ Solution** :
```vba
Function DiviserNombres(ByVal a As Double, ByVal b As Double) As Double
    On Error GoTo ErrHandler

    ' ─── Validation ───
    If b = 0 Then
        MsgBox "Division par zéro impossible", vbExclamation
        DiviserNombres = 0
        Exit Function
    End If

    ' ─── Calcul ───
    DiviserNombres = a / b
    Exit Function

ErrHandler:
    DiviserNombres = 0
    Debug.Print "Erreur : " & Err.Description
End Function
```

**📚 Checklist Validation** :
- [ ] Tester les valeurs nulles/vides
- [ ] Tester les divisions par zéro
- [ ] Tester les plages (min/max)
- [ ] Tester les types attendus (`IsNumeric`, `IsDate`)

---

## 🔗 Ressources Externes

### Documentation Officielle
- 📖 **[Microsoft Learn - Procédures VBA](https://learn.microsoft.com/fr-fr/office/vba/language/reference/user-interface-help/sub-statement)** → Syntaxe Sub et Function
- 📖 **[Microsoft Learn - Fonctions VBA](https://learn.microsoft.com/fr-fr/office/vba/language/reference/functions-visual-basic-for-applications)** → Liste complète des fonctions intégrées
- 📖 **[ByRef vs ByVal](https://learn.microsoft.com/fr-fr/office/vba/language/concepts/getting-started/passing-arguments-by-value-and-by-reference)** → Différences expliquées par Microsoft

### Tutoriels Pratiques
- 🎥 **[Leila Gharani - VBA Functions](https://www.youtube.com/@LeilaGharani)** → Tutoriels vidéo (EN, sous-titres FR)
- 🌐 **[Excel-Pratique - Procédures](https://www.excel-pratique.com/fr/vba/procedures)** → Cours et exemples en français
- 🌐 **[XLerateur - Fonctions](https://www.xlerateur.com/)** → Bonnes pratiques professionnelles

### Communautés
- 💬 **[Stack Overflow - Tag VBA](https://stackoverflow.com/questions/tagged/vba)** → Questions/Réponses
- 💬 **[Reddit r/vba](https://www.reddit.com/r/vba/)** → Entraide communautaire

---

## ⏭️ Module Suivant

### Module 8 : Gestion des Erreurs et Débogage Avancé

**Contenu à venir** :
- On Error GoTo : Gestion des erreurs structurée
- Err.Number et Err.Description : Identifier les erreurs
- Resume, Resume Next, Resume Label
- Debug.Print et Debug.Assert : Traçage avancé
- Fenêtre Espions et pile d'appels
- Création de logs d'erreurs

**Pré-requis pour le Module 8** :
- ✓ Maîtrise des procédures avec paramètres (Module 7)
- ✓ Comprendre les structures conditionnelles
- ✓ Savoir utiliser la fenêtre Exécution (Ctrl+G)

---

## 📌 Mémo Récapitulatif

### Syntaxe des Procédures

```vba
' ═══ Sub (Procédure) ═══
Public Sub NomProcedure(ByVal param1 As Type, ByRef param2 As Type)
    ' Actions
End Sub

' ═══ Function (Fonction) ═══
Public Function NomFonction(ByVal param As Type) As TypeRetour
    NomFonction = resultat ' Retour de valeur
End Function

' ═══ Paramètre Optionnel ═══
Function Calcul(ByVal x As Double, Optional ByVal y As Double = 0) As Double
    Calcul = x + y
End Function
```

### ByRef vs ByVal

| Aspect | ByRef | ByVal |
|--------|-------|-------|
| **Passe** | Adresse mémoire | Copie de la valeur |
| **Modification** | ✅ Modifie l'original | ❌ Ne modifie pas |
| **Performance** | Rapide (gros objets) | Lent (copie) |
| **Sécurité** | Risque d'effets de bord | ✅ Protégé |
| **Par défaut** | ✅ Oui | ❌ Non |

### Top 20 Fonctions VBA Intégrées

| Catégorie | Fonctions |
|-----------|-----------|
| **String** | `Len()`, `Trim()`, `UCase()`, `LCase()`, `Left()`, `Right()`, `Mid()`, `InStr()`, `Replace()` |
| **Math** | `Round()`, `Int()`, `Abs()`, `Sqr()`, `Rnd()` |
| **Date** | `Date`, `Now`, `Year()`, `Month()`, `Day()`, `DateAdd()`, `DateDiff()`, `Format()` |
| **Conversion** | `CStr()`, `CInt()`, `CDbl()`, `CLng()` |
| **Test** | `IsNumeric()`, `IsDate()`, `IsEmpty()`, `IsNull()` |

---

## ✅ Checklist de Fin de Module

Avant de passer au Module 8, assurez-vous de pouvoir :

- [ ] Expliquer la différence entre Sub et Function
- [ ] Créer une fonction avec paramètres typés et valeur de retour
- [ ] Choisir entre ByRef et ByVal selon le contexte
- [ ] Utiliser 10+ fonctions VBA intégrées dans votre code
- [ ] Valider les paramètres d'entrée dans vos fonctions
- [ ] Créer des procédures modulaires et réutilisables
- [ ] Déboguer pas-à-pas (F8) une fonction complexe
- [ ] Avoir réussi 80% du QCM
- [ ] Avoir terminé 2/3 exercices pratiques
- [ ] Avoir obtenu 70/100 au mini-projet

**🎓 Si checklist complète** → Vous êtes prêt pour le Module 8 !
**⚠️ Si < 80%** → Revoir les sections marquées et refaire les exercices

---

**Version** : 1.0 (05/11/2025)
**Auteur** : Expert-Formateur VBA Excel - Certifications TOSA & ICDL
**Durée estimée** : 8-10 heures (théorie + pratique)
**Niveau** : Intermédiaire/Avancé (Niveau 3)

---

*Ce cours est conforme aux référentiels TOSA Programmation et ICDL Advanced Spreadsheets.*
