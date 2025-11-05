Attribute VB_Name = "Module_Exercices_Solutions"
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' Module      : Module_Exercices_Solutions
' Description : Solutions des exercices Module 7 - Procédures avec Paramètres
' Auteur      : Formation VBA Excel - TOSA & ICDL
' Date        : 05/11/2025
'═══════════════════════════════════════════════════════════════════════════════

' ═══════════════════════════════════════════════════════════════════════════════
' EXERCICE 1 : VALIDATION EMAIL ⭐⭐
' ═══════════════════════════════════════════════════════════════════════════════

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : ValiderEmail
' But       : Valider une adresse email selon critères de base
' Entrées   : email (String)
' Sortie    : Boolean (True si valide)
' Critères  : 1 @, au moins 1 . après @, longueur min 5, pas d'espaces
'───────────────────────────────────────────────────────────────────────────────

Public Function ValiderEmail(ByVal email As String) As Boolean
    Dim posArobase As Long
    Dim posPoint As Long

    ' ─── Initialisation ───
    ValiderEmail = False

    ' ─── Critère 1 : Longueur minimale 5 caractères ───
    If Len(email) < 5 Then
        Exit Function
    End If

    ' ─── Critère 2 : Pas d'espaces ───
    If InStr(email, " ") > 0 Then
        Exit Function
    End If

    ' ─── Critère 3 : Exactement un @ ───
    posArobase = InStr(email, "@")
    If posArobase = 0 Then
        Exit Function ' Pas de @
    End If

    ' Vérifier qu'il n'y a qu'un seul @
    If InStr(posArobase + 1, email, "@") > 0 Then
        Exit Function ' Plus d'un @
    End If

    ' ─── Critère 4 : Au moins un . après le @ ───
    posPoint = InStr(posArobase + 1, email, ".")
    If posPoint = 0 Then
        Exit Function ' Pas de . après @
    End If

    ' ─── Si tous les critères OK ───
    ValiderEmail = True
End Function

' ─── Procédure de Test ───
Public Sub TestValiderEmail()
    Dim emails(1 To 8) As String
    Dim i As Integer
    Dim resultat As Boolean

    ' ═══ Cas de test ═══
    emails(1) = "user@example.com"      ' VALIDE
    emails(2) = "jean.dupont@mail.fr"   ' VALIDE
    emails(3) = "contact@site.co.uk"    ' VALIDE
    emails(4) = "invalide"              ' INVALIDE (pas de @)
    emails(5) = "test@test"             ' INVALIDE (pas de . après @)
    emails(6) = "user @mail.com"        ' INVALIDE (espace)
    emails(7) = "a@b.c"                 ' VALIDE (limite basse)
    emails(8) = "user@@mail.com"        ' INVALIDE (2 @)

    Debug.Print String(80, "═")
    Debug.Print "TEST VALIDATION EMAIL"
    Debug.Print String(80, "═")

    For i = 1 To UBound(emails)
        resultat = ValiderEmail(emails(i))
        Debug.Print "[" & IIf(resultat, "✓", "✗") & "] " & _
                    emails(i) & " → " & IIf(resultat, "VALIDE", "INVALIDE")
    Next i

    MsgBox "Test terminé ! Voir la fenêtre Exécution (Ctrl+G)", vbInformation
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' EXERCICE 2 : CALCUL DE REMISE ⭐⭐⭐
' ═══════════════════════════════════════════════════════════════════════════════

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : ObtenirTauxRemise
' But       : Déterminer le taux de remise selon le montant HT
' Entrées   : montantHT (Double)
' Sortie    : Taux de remise (Double) entre 0 et 0.15
' Barème    : < 100 : 0% | 100-499 : 5% | 500-999 : 10% | >= 1000 : 15%
'───────────────────────────────────────────────────────────────────────────────

Public Function ObtenirTauxRemise(ByVal montantHT As Double) As Double
    Select Case montantHT
        Case Is < 100
            ObtenirTauxRemise = 0       ' 0%
        Case 100 To 499.99
            ObtenirTauxRemise = 0.05    ' 5%
        Case 500 To 999.99
            ObtenirTauxRemise = 0.1     ' 10%
        Case Is >= 1000
            ObtenirTauxRemise = 0.15    ' 15%
        Case Else
            ObtenirTauxRemise = 0
    End Select
End Function

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : CalculerMontantRemise
' But       : Calculer le montant de la remise en euros
' Entrées   : montantHT (Double)
' Sortie    : Montant remise (Double)
'───────────────────────────────────────────────────────────────────────────────

Public Function CalculerMontantRemise(ByVal montantHT As Double) As Double
    Dim tauxRemise As Double

    tauxRemise = ObtenirTauxRemise(montantHT)
    CalculerMontantRemise = Round(montantHT * tauxRemise, 2)
End Function

'───────────────────────────────────────────────────────────────────────────────
' Procédure : AppliquerRemisesPlage
' But       : Appliquer les remises sur une plage Excel
' Entrées   : plageDebut (Range) - première cellule de la colonne A
' Sorties   : Colonnes B (Taux), C (Remise), D (Final)
'───────────────────────────────────────────────────────────────────────────────

Public Sub AppliquerRemisesPlage(ByVal plageDebut As Range)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim i As Long
    Dim montantHT As Double
    Dim tauxRemise As Double
    Dim montantRemise As Double
    Dim montantFinal As Double

    ' ─── Initialisation ───
    Set ws = plageDebut.Worksheet
    derniereLigne = ws.Cells(ws.Rows.Count, plageDebut.Column).End(xlUp).Row

    ' ─── En-têtes ───
    With ws
        .Cells(plageDebut.Row, 1).Value = "Montant HT"
        .Cells(plageDebut.Row, 2).Value = "Taux Remise"
        .Cells(plageDebut.Row, 3).Value = "Montant Remise"
        .Cells(plageDebut.Row, 4).Value = "Montant Final"

        ' Mise en forme en-têtes
        .Range(.Cells(plageDebut.Row, 1), .Cells(plageDebut.Row, 4)).Font.Bold = True
    End With

    ' ─── Traitement ligne par ligne ───
    For i = plageDebut.Row + 1 To derniereLigne
        ' Récupérer le montant HT
        montantHT = ws.Cells(i, 1).Value

        ' Calculer
        tauxRemise = ObtenirTauxRemise(montantHT)
        montantRemise = CalculerMontantRemise(montantHT)
        montantFinal = montantHT - montantRemise

        ' Écrire les résultats
        With ws
            .Cells(i, 2).Value = Format(tauxRemise, "0%")
            .Cells(i, 3).Value = Format(montantRemise, "0.00") & " €"
            .Cells(i, 4).Value = Format(montantFinal, "0.00") & " €"
        End With
    Next i

    MsgBox "Remises appliquées avec succès sur " & (derniereLigne - plageDebut.Row) & " lignes", _
           vbInformation, "Traitement terminé"

    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Erreur dans AppliquerRemisesPlage"
End Sub

' ─── Procédure de Test ───
Public Sub TestCalculRemise()
    Dim ws As Worksheet

    ' Créer ou récupérer la feuille de test
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Test_Remises")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Test_Remises"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' ═══ Données de test ═══
    With ws
        .Range("A2").Value = 50
        .Range("A3").Value = 250
        .Range("A4").Value = 750
        .Range("A5").Value = 1500
    End With

    ' ═══ Appliquer les remises ═══
    AppliquerRemisesPlage ws.Range("A1")
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' EXERCICE 3 : GÉNÉRATION DE RÉFÉRENCE ⭐⭐⭐
' ═══════════════════════════════════════════════════════════════════════════════

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : GenererReference
' But       : Générer une référence produit au format PROD-YYYY-XXXXX
' Entrées   : numero (Long) entre 1 et 99999
' Sortie    : Référence formatée (String)
' Exemple   : GenererReference(42) → "PROD-2025-00042"
'───────────────────────────────────────────────────────────────────────────────

Public Function GenererReference(ByVal numero As Long) As String
    Dim annee As Integer
    Dim numeroFormate As String

    ' ─── Validation ───
    If numero < 1 Or numero > 99999 Then
        GenererReference = "ERREUR"
        Exit Function
    End If

    ' ─── Récupérer année en cours ───
    annee = Year(Date)

    ' ─── Formater le numéro sur 5 chiffres ───
    numeroFormate = Format(numero, "00000")

    ' ─── Concaténer ───
    GenererReference = "PROD-" & annee & "-" & numeroFormate
End Function

' ─── Procédure de Test ───
Public Sub TestGenererReference()
    Dim i As Long
    Dim refs(1 To 10) As Long

    ' ═══ Cas de test ═══
    refs(1) = 1
    refs(2) = 42
    refs(3) = 100
    refs(4) = 999
    refs(5) = 1234
    refs(6) = 12345
    refs(7) = 99999
    refs(8) = 0         ' Invalide
    refs(9) = 100000    ' Invalide
    refs(10) = -5       ' Invalide

    Debug.Print String(80, "═")
    Debug.Print "TEST GÉNÉRATION DE RÉFÉRENCES"
    Debug.Print String(80, "═")

    For i = 1 To UBound(refs)
        Debug.Print "Numéro " & Format(refs(i), "000000") & " → " & GenererReference(refs(i))
    Next i

    MsgBox "Test terminé ! Voir la fenêtre Exécution (Ctrl+G)", vbInformation
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' MINI-PROJET : SYSTÈME DE GESTION DE CATALOGUE PRODUITS
' ═══════════════════════════════════════════════════════════════════════════════

' ─── Variable de module pour le compteur ───
Private compteurReference As Long

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : GenererReferenceProduit
' But       : Générer référence au format CAT-YYYY-NNNNN
' Entrées   : categorie (String), numero (Long)
' Sortie    : Référence (String)
'───────────────────────────────────────────────────────────────────────────────

Public Function GenererReferenceProduit(ByVal categorie As String, _
                                       ByVal numero As Long) As String
    Dim catCode As String
    Dim annee As Integer

    ' ─── Valider et formater catégorie (3 lettres majuscules) ───
    catCode = UCase(Left(Trim(categorie), 3))

    ' ─── Année en cours ───
    annee = Year(Date)

    ' ─── Formater ───
    GenererReferenceProduit = catCode & "-" & annee & "-" & Format(numero, "00000")
End Function

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : ValiderNomProduit
' But       : Valider le nom d'un produit
' Entrées   : nom (String)
' Sortie    : Boolean
' Critères  : 3-50 caractères, pas de caractères spéciaux
'───────────────────────────────────────────────────────────────────────────────

Public Function ValiderNomProduit(ByVal nom As String) As Boolean
    Dim longueur As Integer
    Dim caracteresInterdits As String
    Dim i As Integer

    ' ─── Initialisation ───
    ValiderNomProduit = False
    caracteresInterdits = "@#$%^&*+=[]{}|<>?;:"

    ' ─── Nettoyer ───
    nom = Trim(nom)
    longueur = Len(nom)

    ' ─── Critère 1 : Longueur entre 3 et 50 ───
    If longueur < 3 Or longueur > 50 Then
        Exit Function
    End If

    ' ─── Critère 2 : Pas de caractères spéciaux ───
    For i = 1 To Len(caracteresInterdits)
        If InStr(nom, Mid(caracteresInterdits, i, 1)) > 0 Then
            Exit Function
        End If
    Next i

    ' ─── Validation OK ───
    ValiderNomProduit = True
End Function

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : CalculerPrixTTC
' But       : Calculer prix TTC selon catégorie
' Entrées   : prixHT (Double), categorie (String)
' Sortie    : Prix TTC (Double)
' TVA       : ALI=5.5%, HYG=20%, ELE=20%
'───────────────────────────────────────────────────────────────────────────────

Public Function CalculerPrixTTC(ByVal prixHT As Double, _
                               ByVal categorie As String) As Double
    Dim tauxTVA As Double

    ' ─── Déterminer le taux de TVA ───
    Select Case UCase(categorie)
        Case "ALI", "ALIMENTAIRE"
            tauxTVA = 0.055    ' 5.5%
        Case "HYG", "HYGIÈNE", "HYGIENE"
            tauxTVA = 0.2      ' 20%
        Case "ELE", "ÉLECTRONIQUE", "ELECTRONIQUE"
            tauxTVA = 0.2      ' 20%
        Case Else
            CalculerPrixTTC = 0 ' Catégorie inconnue
            Exit Function
    End Select

    ' ─── Calculer TTC ───
    CalculerPrixTTC = Round(prixHT * (1 + tauxTVA), 2)
End Function

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : FormaterNomProduit
' But       : Formater nom produit (première lettre en majuscule)
' Entrées   : nom (String)
' Sortie    : Nom formaté (String)
'───────────────────────────────────────────────────────────────────────────────

Public Function FormaterNomProduit(ByVal nom As String) As String
    ' ─── Supprimer espaces début/fin ───
    nom = Trim(nom)

    ' ─── Première lettre de chaque mot en majuscule ───
    FormaterNomProduit = StrConv(nom, vbProperCase)
End Function

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : ObtenirCodeCategorie
' But       : Convertir nom catégorie vers code 3 lettres
' Entrées   : nomCategorie (String)
' Sortie    : Code 3 lettres (String)
'───────────────────────────────────────────────────────────────────────────────

Public Function ObtenirCodeCategorie(ByVal nomCategorie As String) As String
    Select Case UCase(Trim(nomCategorie))
        Case "ALIMENTAIRE"
            ObtenirCodeCategorie = "ALI"
        Case "HYGIÈNE", "HYGIENE"
            ObtenirCodeCategorie = "HYG"
        Case "ÉLECTRONIQUE", "ELECTRONIQUE"
            ObtenirCodeCategorie = "ELE"
        Case Else
            ObtenirCodeCategorie = "" ' Catégorie inconnue
    End Select
End Function

'───────────────────────────────────────────────────────────────────────────────
' Procédure : TraiterCatalogueProduits
' But       : Traiter un catalogue complet
' Entrées   : feuille (Worksheet)
' Sorties   : Colonnes D à H remplies
'───────────────────────────────────────────────────────────────────────────────

Public Sub TraiterCatalogueProduits(ByVal feuille As Worksheet)
    On Error GoTo ErrHandler

    Dim derniereLigne As Long
    Dim i As Long
    Dim categorie As String
    Dim nomProduit As String
    Dim prixHT As Double
    Dim codeCategorie As String
    Dim reference As String
    Dim nomNettoye As String
    Dim prixTTC As Double
    Dim statut As String

    ' ─── Initialisation ───
    derniereLigne = feuille.Cells(feuille.Rows.Count, "A").End(xlUp).Row
    compteurReference = 0

    ' ─── En-têtes ───
    With feuille
        .Range("D1").Value = "Code Catégorie"
        .Range("E1").Value = "Référence"
        .Range("F1").Value = "Nom Nettoyé"
        .Range("G1").Value = "Prix TTC"
        .Range("H1").Value = "Statut"
        .Range("A1:H1").Font.Bold = True
    End With

    ' ─── Traitement ligne par ligne ───
    For i = 2 To derniereLigne
        ' Récupérer les données
        categorie = feuille.Cells(i, 1).Value
        nomProduit = feuille.Cells(i, 2).Value
        prixHT = feuille.Cells(i, 3).Value

        ' ═══ Traitement ═══

        ' Code catégorie
        codeCategorie = ObtenirCodeCategorie(categorie)

        ' Validation du nom
        If Not ValiderNomProduit(nomProduit) Then
            ' ═══ ERREUR : Nom invalide ═══
            feuille.Cells(i, 4).Value = ""
            feuille.Cells(i, 5).Value = ""
            feuille.Cells(i, 6).Value = nomProduit
            feuille.Cells(i, 7).Value = 0
            feuille.Cells(i, 8).Value = "ERREUR: Nom invalide"
            feuille.Cells(i, 8).Font.Color = RGB(255, 0, 0) ' Rouge
        ElseIf codeCategorie = "" Then
            ' ═══ ERREUR : Catégorie inconnue ═══
            feuille.Cells(i, 4).Value = ""
            feuille.Cells(i, 5).Value = ""
            feuille.Cells(i, 6).Value = FormaterNomProduit(nomProduit)
            feuille.Cells(i, 7).Value = 0
            feuille.Cells(i, 8).Value = "ERREUR: Catégorie inconnue"
            feuille.Cells(i, 8).Font.Color = RGB(255, 0, 0)
        Else
            ' ═══ OK : Traitement complet ═══
            compteurReference = compteurReference + 1

            ' Génération référence
            reference = GenererReferenceProduit(codeCategorie, compteurReference)

            ' Nom nettoyé
            nomNettoye = FormaterNomProduit(nomProduit)

            ' Prix TTC
            prixTTC = CalculerPrixTTC(prixHT, codeCategorie)

            ' Écrire les résultats
            feuille.Cells(i, 4).Value = codeCategorie
            feuille.Cells(i, 5).Value = reference
            feuille.Cells(i, 6).Value = nomNettoye
            feuille.Cells(i, 7).Value = Format(prixTTC, "0.00")
            feuille.Cells(i, 8).Value = "OK"
            feuille.Cells(i, 8).Font.Color = RGB(0, 128, 0) ' Vert
        End If
    Next i

    ' ─── Ajuster largeur colonnes ───
    feuille.Columns("A:H").AutoFit

    MsgBox "Traitement terminé ! " & compteurReference & " produits traités.", _
           vbInformation, "Catalogue Produits"

    Exit Sub

ErrHandler:
    MsgBox "Erreur ligne " & i & " : " & Err.Description, vbCritical, "Erreur"
End Sub

' ─── Procédure de Test du Mini-Projet ───
Public Sub TestCatalogue()
    Dim ws As Worksheet

    ' Créer ou récupérer la feuille de test
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Catalogue")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Catalogue"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' ═══ En-têtes ═══
    With ws
        .Range("A1").Value = "Catégorie"
        .Range("B1").Value = "Nom"
        .Range("C1").Value = "Prix HT"
        .Range("A1:C1").Font.Bold = True
    End With

    ' ═══ Données de test ═══
    With ws
        .Range("A2").Value = "Alimentaire"
        .Range("B2").Value = " pâtes bio "
        .Range("C2").Value = 2.5

        .Range("A3").Value = "Hygiène"
        .Range("B3").Value = "SAVON liquide"
        .Range("C3").Value = 3.8

        .Range("A4").Value = "Électronique"
        .Range("B4").Value = "   souris sans fil   "
        .Range("C4").Value = 15

        .Range("A5").Value = "Alimentaire"
        .Range("B5").Value = "huile d'olive"
        .Range("C5").Value = 8.5

        .Range("A6").Value = "Mobilier"
        .Range("B6").Value = "chaise@bureau"
        .Range("C6").Value = 120
    End With

    ' ═══ Traiter le catalogue ═══
    TraiterCatalogueProduits ws

    ' ═══ Activer la feuille ═══
    ws.Activate
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' PROCÉDURE DE TEST COMPLÈTE - TOUS LES EXERCICES
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub TesterTousLesExercices()
    MsgBox "Début des tests. Ouvrez la fenêtre Exécution (Ctrl+G) pour certains résultats.", vbInformation

    ' Test Exercice 1
    TestValiderEmail

    ' Test Exercice 2
    TestCalculRemise

    ' Test Exercice 3
    TestGenererReference

    ' Test Mini-Projet
    TestCatalogue

    MsgBox "Tous les tests des exercices sont terminés !", vbInformation
End Sub
