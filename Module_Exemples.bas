Attribute VB_Name = "Module_Exemples"
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' Module      : Module_Exemples
' Description : Exemples du cours Module 7 - Procédures avec Paramètres
' Auteur      : Formation VBA Excel - TOSA & ICDL
' Date        : 05/11/2025
'═══════════════════════════════════════════════════════════════════════════════

' ═══════════════════════════════════════════════════════════════════════════════
' EXEMPLE 1 : PROCÉDURE AVEC PARAMÈTRES SIMPLES
' ═══════════════════════════════════════════════════════════════════════════════

'───────────────────────────────────────────────────────────────────────────────
' Procédure : AfficherMessage
' But       : Afficher un message personnalisé
' Entrées   : prenom (String), age (Integer)
' Sorties   : MsgBox avec message formaté
'───────────────────────────────────────────────────────────────────────────────

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

' ═══════════════════════════════════════════════════════════════════════════════
' EXEMPLE 2 : FONCTION AVEC RETOUR DE VALEUR
' ═══════════════════════════════════════════════════════════════════════════════

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : CalculerTVA
' But       : Calculer le montant TTC à partir du HT
' Entrées   : montantHT (Double), tauxTVA (Double)
' Sortie    : Montant TTC (Double)
' Exemple   : montantTTC = CalculerTVA(100, 0.2)
'───────────────────────────────────────────────────────────────────────────────

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
        .Range("B2").Value = CalculerTVA(.Range("A2").Value)
    End With

    MsgBox "Tests terminés ! Voir la fenêtre Exécution (Ctrl+G)", vbInformation
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' EXEMPLE 3 : BYREF VS BYVAL - ILLUSTRATION PRATIQUE
' ═══════════════════════════════════════════════════════════════════════════════

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

' ═══════════════════════════════════════════════════════════════════════════════
' EXEMPLE 4 : CAS MÉTIER - FONCTIONS VBA INTÉGRÉES
' ═══════════════════════════════════════════════════════════════════════════════

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : NettoyerTexte
' But       : Nettoyer une chaîne (trim, majuscules, espaces multiples)
' Entrées   : texte (String)
' Sortie    : Texte nettoyé (String)
'───────────────────────────────────────────────────────────────────────────────

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

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : FormaterCodePostal
' But       : Formater un code postal français (5 chiffres)
' Entrées   : codePostal (String)
' Sortie    : Code postal formaté ou "" si invalide
'───────────────────────────────────────────────────────────────────────────────

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

'───────────────────────────────────────────────────────────────────────────────
' Fonction  : CalculerAge
' But       : Calculer l'âge à partir de la date de naissance
' Entrées   : dateNaissance (Date)
' Sortie    : Age en années (Integer)
'───────────────────────────────────────────────────────────────────────────────

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

    MsgBox "Tests terminés ! Voir la fenêtre Exécution (Ctrl+G)", vbInformation
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' PROCÉDURE DE TEST COMPLÈTE - TOUS LES EXEMPLES
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub TesterTousLesExemples()
    MsgBox "Début des tests. Ouvrez la fenêtre Exécution (Ctrl+G) pour voir les résultats.", vbInformation

    Debug.Print String(80, "═")
    Debug.Print "TEST DES EXEMPLES - MODULE 7"
    Debug.Print String(80, "═")

    Debug.Print vbCrLf & "1. Test AfficherMessage"
    TestAfficherMessage

    Debug.Print vbCrLf & "2. Test CalculerTVA"
    TestCalculerTVA

    Debug.Print vbCrLf & "3. Test ByRef vs ByVal"
    TestByRefByVal

    Debug.Print vbCrLf & "4. Test Fonctions VBA Intégrées"
    TestFonctionsVBA

    Debug.Print vbCrLf & String(80, "═")
    Debug.Print "TOUS LES TESTS TERMINÉS !"
    Debug.Print String(80, "═")

    MsgBox "Tous les tests sont terminés ! Consultez la fenêtre Exécution (Ctrl+G)", vbInformation
End Sub
