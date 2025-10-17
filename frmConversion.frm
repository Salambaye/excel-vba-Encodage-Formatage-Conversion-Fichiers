VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConversion 
   Caption         =   "UserForm2"
   ClientHeight    =   10875
   ClientLeft      =   -30
   ClientTop       =   -60
   ClientWidth     =   20640
   OleObjectBlob   =   "frmConversion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
''Private Sub Label1_Click()
''
''End Sub
''
''Private Sub CommandButton1_Click()
''
''End Sub
''
''Private Sub lblFormatageActuel_Click()
''
''End Sub
''
''Private Sub UserForm_Click()
''
''End Sub
'
''' dans frmConversion
''Private m_CheminFichier As String
''
''
''
''Public Property Let CheminFichier(val As String)
''    m_CheminFichier = val
''End Property
''
''Public Property Get CheminFichier() As String
''    CheminFichier = m_CheminFichier
''End Property
'
'Public CheminFichier As String
'Private encodageDetecte As String
'Private formatageDetecte As String
'Private extensionDetectee As String
'
'Private Sub optANSI_Click()
'
'End Sub
'
'Private Sub optUTF8_Click()
'
'End Sub
'
'Private Sub optWindows_Click()
'
'End Sub
'
'Private Sub UserForm_Initialize()
'    ' Initialisation du formulaire
'    Me.Caption = "Conversion de fichier"
'End Sub
'
'Public Sub DetecterParametresFichier()
'    Dim fso As Object
'    Dim ts As Object
'    Dim contenu As String
'    Dim premierOctet() As Byte
'    Dim fichierNum As Integer
'
'    On Error GoTo ErreurDetection
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' Détecter l'extension
'    extensionDetectee = UCase(fso.GetExtensionName(CheminFichier))
'
'    ' Détecter l'encodage
'    encodageDetecte = DetecterEncodage(CheminFichier)
'
'    ' Détecter le formatage (Windows CRLF ou Unix LF)
'    formatageDetecte = DetecterFormatage(CheminFichier)
'
'    ' Afficher les informations détectées
'    lblFichierActuel.Caption = "Fichier : " & fso.GetFileName(CheminFichier)
'    lblEncodageActuel.Caption = "Encodage actuel : " & encodageDetecte
'    lblFormatageActuel.Caption = "Formatage actuel : " & formatageDetecte
'    lblExtensionActuelle.Caption = "Extension actuelle : " & extensionDetectee
'
'    ' Cocher les options correspondant au fichier actuel
'    Select Case UCase(encodageDetecte)
'        Case "UTF-8"
'            optUTF8.Value = True
'        Case "ANSI"
'            optANSI.Value = True
'        Case Else
'            optUTF8.Value = True ' Par défaut
'    End Select
'
'    Select Case UCase(formatageDetecte)
'        Case "WINDOWS"
'            optWindows.Value = True
'        Case "UNIX"
'            optUnix.Value = True
'        Case Else
'            optWindows.Value = True ' Par défaut
'    End Select
'
'    Select Case UCase(extensionDetectee)
'        Case "TXT"
'            optTxt.Value = True
'        Case "CSV"
'            optCsv.Value = True
'        Case Else
'            optTxt.Value = True ' Par défaut
'    End Select
'
'    Exit Sub
'
'ErreurDetection:
'    MsgBox "Erreur lors de la détection : " & Err.Description, vbExclamation
'    ' Valeurs par défaut en cas d'erreur
'    optUTF8.Value = True
'    optWindows.Value = True
'    optTxt.Value = True
'End Sub
'
'Private Function DetecterEncodage(fichier As String) As String
'    Dim fichierNum As Integer
'    Dim bom(1 To 3) As Byte
'    Dim i As Integer
'
'    On Error GoTo ErreurEncodage
'
'    fichierNum = FreeFile
'    Open fichier For Binary Access Read As #fichierNum
'
'    ' Lire les 3 premiers octets pour détecter le BOM UTF-8
'    If LOF(fichierNum) >= 3 Then
'        For i = 1 To 3
'            Get #fichierNum, , bom(i)
'        Next i
'
'        ' BOM UTF-8 : EF BB BF
'        If bom(1) = &HEF And bom(2) = &HBB And bom(3) = &HBF Then
'            DetecterEncodage = "UTF-8"
'            Close #fichierNum
'            Exit Function
'        End If
'    End If
'
'    Close #fichierNum
'
'    ' Si pas de BOM, considérer comme ANSI
'    DetecterEncodage = "ANSI"
'    Exit Function
'
'ErreurEncodage:
'    If fichierNum > 0 Then Close #fichierNum
'    DetecterEncodage = "ANSI"
'End Function
'
'Private Function DetecterFormatage(fichier As String) As String
'    Dim fso As Object
'    Dim ts As Object
'    Dim contenu As String
'    Dim posLF As Long
'    Dim posCRLF As Long
'
'    On Error GoTo ErreurFormatage
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set ts = fso.OpenTextFile(fichier, 1, False, -2) ' -2 pour ouvrir comme système
'
'    ' Lire une partie du fichier (premiers 10000 caractères)
'    If Not ts.AtEndOfStream Then
'        contenu = ts.Read(10000)
'    End If
'    ts.Close
'
'    ' Chercher CRLF et LF
'    posCRLF = InStr(1, contenu, vbCrLf)
'    posLF = InStr(1, contenu, vbLf)
'
'    If posCRLF > 0 Then
'        DetecterFormatage = "Windows"
'    ElseIf posLF > 0 Then
'        DetecterFormatage = "Unix"
'    Else
'        DetecterFormatage = "Windows" ' Par défaut
'    End If
'
'    Exit Function
'
'ErreurFormatage:
'    DetecterFormatage = "Windows"
'End Function
'
'Private Sub btnConvertir_Click()
'    Dim encodageCible As String
'    Dim formatageCible As String
'    Dim extensionCible As String
'    Dim fdlgDossier As FileDialog
'    Dim dossierSauvegarde As String
'    Dim fso As Object
'    Dim nomFichierSansExt As String
'    Dim fichierSortie As String
'
'    On Error GoTo ErreurConversion
'
'    ' Récupérer les options sélectionnées
'    If optUTF8.Value Then
'        encodageCible = "UTF-8"
'    Else
'        encodageCible = "ANSI"
'    End If
'
'    If optWindows.Value Then
'        formatageCible = "Windows"
'    Else
'        formatageCible = "Unix"
'    End If
'
'    If optTxt.Value Then
'        extensionCible = "txt"
'    Else
'        extensionCible = "csv"
'    End If
'
'    ' Sélection du dossier de sauvegarde
'    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
'    With fdlgDossier
'        .Title = "Sélectionner le dossier de sauvegarde"
'        .AllowMultiSelect = False
'        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
'    End With
'
'    If fdlgDossier.Show <> -1 Then
'        MsgBox "Sélection annulée.", vbInformation
'        Exit Sub
'    End If
'
'    dossierSauvegarde = fdlgDossier.SelectedItems(1)
'
'    If Dir(dossierSauvegarde, vbDirectory) = "" Then
'        MsgBox "Le dossier sélectionné n'est pas accessible.", vbCritical
'        Exit Sub
'    End If
'
'    ' Créer le nom du fichier de sortie
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    nomFichierSansExt = fso.GetBaseName(CheminFichier)
'
'    If Right(dossierSauvegarde, 1) <> "\" Then
'        dossierSauvegarde = dossierSauvegarde & "\"
'    End If
'
'    fichierSortie = dossierSauvegarde & nomFichierSansExt & "_converti." & extensionCible
'
'    ' Vérifier si le fichier existe déjà
'    If Dir(fichierSortie) <> "" Then
'        If MsgBox("Le fichier existe déjà :" & vbCrLf & fichierSortie & vbCrLf & vbCrLf & _
'                  "Voulez-vous le remplacer ?", vbYesNo + vbQuestion) = vbNo Then
'            Exit Sub
'        End If
'    End If
'
'    ' Effectuer la conversion
'    ConvertirFichierComplet CheminFichier, fichierSortie, encodageCible, formatageCible
'
'    MsgBox "Conversion terminée avec succès !" & vbCrLf & vbCrLf & _
'           "Fichier créé : " & fichierSortie, vbInformation
'
'    ' Ouvrir le dossier contenant le fichier converti
'    Shell "explorer.exe /select,""" & fichierSortie & """", vbNormalFocus
'
'    Unload Me
'    Exit Sub
'
'ErreurConversion:
'    MsgBox "Erreur lors de la conversion : " & Err.Description, vbCritical
'End Sub
'
'Private Sub ConvertirFichierComplet(source As String, destination As String, _
'                                     encodage As String, formatage As String)
'    Dim fso As Object
'    Dim tsInput As Object
'    Dim tsOutput As Object
'    Dim contenu As String
'    Dim encodageInput As Integer
'    Dim encodageOutput As Integer
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' Déterminer l'encodage d'entrée
'    If DetecterEncodage(source) = "UTF-8" Then
'        encodageInput = -1 ' UTF-8
'    Else
'        encodageInput = 0 ' ANSI
'    End If
'
'    ' Lire le fichier source
'    Set tsInput = fso.OpenTextFile(source, 1, False, encodageInput)
'    contenu = tsInput.ReadAll
'    tsInput.Close
'
'    ' Appliquer le formatage
'    If formatage = "Windows" Then
'        ' Normaliser d'abord tout en LF, puis convertir en CRLF
'        contenu = Replace(contenu, vbCrLf, vbLf)
'        contenu = Replace(contenu, vbCr, vbLf)
'        contenu = Replace(contenu, vbLf, vbCrLf)
'    Else ' Unix
'        ' Convertir tout en LF
'        contenu = Replace(contenu, vbCrLf, vbLf)
'        contenu = Replace(contenu, vbCr, vbLf)
'    End If
'
'    ' Déterminer l'encodage de sortie
'    If encodage = "UTF-8" Then
'        encodageOutput = -1
'    Else
'        encodageOutput = 0
'    End If
'
'    ' Écrire le fichier de sortie
'    Set tsOutput = fso.OpenTextFile(destination, 2, True, encodageOutput)
'    tsOutput.Write contenu
'    tsOutput.Close
'
'    Set fso = Nothing
'End Sub
'
'Private Sub btnAnnuler_Click()
'    Unload Me
'End Sub
'






'' ========================================
'' USERFORM frmConversion - Code amélioré
'' ========================================
'
'Option Explicit
'
'Public CheminFichier As String
'Private encodageDetecte As String
'Private formatageDetecte As String
'Private extensionDetectee As String
'
'Private Sub UserForm_Initialize()
'    ' Initialisation du formulaire
'    Me.Caption = "Conversion de fichier"
'
'    ' Améliorer l'apparence des frames
'    With Me.FrameEncodage
'        .BackColor = RGB(230, 245, 230)
'        .BorderColor = RGB(0, 120, 100)
'        .SpecialEffect = fmSpecialEffectRaised
'    End With
'
'    With Me.FrameFormatage
'        .BackColor = RGB(255, 250, 220)
'        .BorderColor = RGB(0, 120, 100)
'        .SpecialEffect = fmSpecialEffectRaised
'    End With
'
'    With Me.FrameExtension
'        .BackColor = RGB(255, 240, 230)
'        .BorderColor = RGB(0, 120, 100)
'        .SpecialEffect = fmSpecialEffectRaised
'    End With
'
'    ' Améliorer les boutons radio - les rendre plus visibles
'    AmeliorerBoutonsRadio
'End Sub
'
'Private Sub AmeliorerBoutonsRadio()
'    ' Augmenter la taille et améliorer l'apparence des boutons radio
'    Dim ctrl As Control
'
'    For Each ctrl In Me.Controls
'        If TypeName(ctrl) = "OptionButton" Then
'            With ctrl
'                .Font.Size = 12
'                .Font.Bold = True
'                .Height = 40
'                .Width = 140
'                ' Espacement pour meilleure lisibilité
'                Select Case ctrl.Name
'                    Case "optUTF8"
'                        .Left = 20
'                        .Top = 30
'                    Case "optANSI"
'                        .Left = 180
'                        .Top = 30
'                    Case "optWindows"
'                        .Left = 20
'                        .Top = 30
'                    Case "optUnix"
'                        .Left = 180
'                        .Top = 30
'                    Case "optTxt"
'                        .Left = 20
'                        .Top = 30
'                    Case "optCsv"
'                        .Left = 180
'                        .Top = 30
'                End Select
'            End With
'        End If
'    Next ctrl
'End Sub
'
'Public Sub InitialiserAvecFichier(chemin As String)
'    CheminFichier = chemin
'    DetecterParametresFichier
'End Sub
'
'Public Sub DetecterParametresFichier()
'    Dim fso As Object
'
'    On Error GoTo ErreurDetection
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' Détecter l'extension
'    extensionDetectee = UCase(fso.GetExtensionName(CheminFichier))
'
'    ' Détecter l'encodage (fonction améliorée)
'    encodageDetecte = DetecterEncodageAmeliore(CheminFichier)
'
'    ' Détecter le formatage
'    formatageDetecte = DetecterFormatage(CheminFichier)
'
'    ' Afficher les informations avec style amélioré
'    With lblFichierActuel
'        .Caption = "Fichier : " & fso.GetFileName(CheminFichier)
'        .Font.Size = 10
'        .Font.Bold = True
'        .BackColor = RGB(173, 216, 230)
'    End With
'
'    With lblEncodageActuel
'        .Caption = "Encodage actuel : " & encodageDetecte
'        .Font.Size = 10
'        .BackColor = RGB(173, 216, 230)
'    End With
'
'    With lblFormatageActuel
'        .Caption = "Formatage actuel : " & formatageDetecte
'        .Font.Size = 10
'        .BackColor = RGB(173, 216, 230)
'    End With
'
'    With lblExtensionActuelle
'        .Caption = "Extension actuelle : " & extensionDetectee
'        .Font.Size = 10
'        .BackColor = RGB(173, 216, 230)
'    End With
'
'    ' Cocher les options correspondantes
'    Select Case UCase(encodageDetecte)
'        Case "UTF-8", "UTF-8 BOM"
'            optUTF8.Value = True
'        Case "ANSI"
'            optANSI.Value = True
'        Case Else
'            optUTF8.Value = True
'    End Select
'
'    Select Case UCase(formatageDetecte)
'        Case "WINDOWS"
'            optWindows.Value = True
'        Case "UNIX"
'            optUnix.Value = True
'        Case Else
'            optWindows.Value = True
'    End Select
'
'    Select Case UCase(extensionDetectee)
'        Case "TXT"
'            optTxt.Value = True
'        Case "CSV"
'            optCsv.Value = True
'        Case Else
'            optTxt.Value = True
'    End Select
'
'    Exit Sub
'
'ErreurDetection:
'    MsgBox "Erreur lors de la détection : " & Err.Description, vbExclamation
'    optUTF8.Value = True
'    optWindows.Value = True
'    optTxt.Value = True
'End Sub
'
'Private Function DetecterEncodageAmeliore(fichier As String) As String
'    ' Fonction améliorée pour détecter correctement UTF-8
'    Dim fichierNum As Integer
'    Dim bom(1 To 4) As Byte
'    Dim buffer() As Byte
'    Dim i As Long
'    Dim tailleBuffer As Long
'    Dim hasNonASCII As Boolean
'    Dim isValidUTF8 As Boolean
'
'    On Error GoTo ErreurEncodage
'
'    fichierNum = FreeFile
'    Open fichier For Binary Access Read As #fichierNum
'
'    If LOF(fichierNum) = 0 Then
'        Close #fichierNum
'        DetecterEncodageAmeliore = "ANSI"
'        Exit Function
'    End If
'
'    ' Vérifier le BOM UTF-8 (EF BB BF)
'    If LOF(fichierNum) >= 3 Then
'        Get #fichierNum, , bom(1)
'        Get #fichierNum, , bom(2)
'        Get #fichierNum, , bom(3)
'
'        If bom(1) = &HEF And bom(2) = &HBB And bom(3) = &HBF Then
'            Close #fichierNum
'            DetecterEncodageAmeliore = "UTF-8 BOM"
'            Exit Function
'        End If
'    End If
'
'    ' Revenir au début du fichier
'    Close #fichierNum
'    fichierNum = FreeFile
'    Open fichier For Binary Access Read As #fichierNum
'
'    ' Lire un échantillon du fichier (premiers 4000 octets max)
'    tailleBuffer = LOF(fichierNum)
'    If tailleBuffer > 4000 Then tailleBuffer = 4000
'
'    ReDim buffer(1 To tailleBuffer)
'    Get #fichierNum, , buffer
'    Close #fichierNum
'
'    ' Analyser les octets pour détecter UTF-8 sans BOM
'    hasNonASCII = False
'    isValidUTF8 = True
'    i = 1
'
'    Do While i <= tailleBuffer
'        If buffer(i) > 127 Then
'            hasNonASCII = True
'
'            ' Vérifier séquence UTF-8
'            If (buffer(i) And &HE0) = &HC0 Then
'                ' Séquence 2 octets (110xxxxx 10xxxxxx)
'                If i + 1 > tailleBuffer Then Exit Do
'                If (buffer(i + 1) And &HC0) <> &H80 Then
'                    isValidUTF8 = False
'                    Exit Do
'                End If
'                i = i + 2
'            ElseIf (buffer(i) And &HF0) = &HE0 Then
'                ' Séquence 3 octets (1110xxxx 10xxxxxx 10xxxxxx)
'                If i + 2 > tailleBuffer Then Exit Do
'                If (buffer(i + 1) And &HC0) <> &H80 Or (buffer(i + 2) And &HC0) <> &H80 Then
'                    isValidUTF8 = False
'                    Exit Do
'                End If
'                i = i + 3
'            ElseIf (buffer(i) And &HF8) = &HF0 Then
'                ' Séquence 4 octets
'                If i + 3 > tailleBuffer Then Exit Do
'                If (buffer(i + 1) And &HC0) <> &H80 Or (buffer(i + 2) And &HC0) <> &H80 Or _
'                   (buffer(i + 3) And &HC0) <> &H80 Then
'                    isValidUTF8 = False
'                    Exit Do
'                End If
'                i = i + 4
'            Else
'                ' Octet invalide pour UTF-8
'                isValidUTF8 = False
'                Exit Do
'            End If
'        Else
'            i = i + 1
'        End If
'    Loop
'
'    ' Déterminer l'encodage
'    If hasNonASCII And isValidUTF8 Then
'        DetecterEncodageAmeliore = "UTF-8"
'    Else
'        DetecterEncodageAmeliore = "ANSI"
'    End If
'
'    Exit Function
'
'ErreurEncodage:
'    If fichierNum > 0 Then Close #fichierNum
'    DetecterEncodageAmeliore = "ANSI"
'End Function
'
'Private Function DetecterFormatage(fichier As String) As String
'    Dim fso As Object
'    Dim ts As Object
'    Dim contenu As String
'    Dim posLF As Long
'    Dim posCRLF As Long
'
'    On Error GoTo ErreurFormatage
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set ts = fso.OpenTextFile(fichier, 1, False, -2)
'
'    If Not ts.AtEndOfStream Then
'        contenu = ts.Read(10000)
'    End If
'    ts.Close
'
'    posCRLF = InStr(1, contenu, vbCrLf)
'    posLF = InStr(1, contenu, vbLf)
'
'    If posCRLF > 0 Then
'        DetecterFormatage = "Windows"
'    ElseIf posLF > 0 Then
'        DetecterFormatage = "Unix"
'    Else
'        DetecterFormatage = "Windows"
'    End If
'
'    Exit Function
'
'ErreurFormatage:
'    DetecterFormatage = "Windows"
'End Function
'
'Private Sub btnConvertir_Click()
'    Dim encodageCible As String
'    Dim formatageCible As String
'    Dim extensionCible As String
'    Dim fdlgDossier As FileDialog
'    Dim dossierSauvegarde As String
'    Dim fso As Object
'    Dim nomFichierSansExt As String
'    Dim fichierSortie As String
'
'    On Error GoTo ErreurConversion
'
'    If optUTF8.Value Then
'        encodageCible = "UTF-8"
'    Else
'        encodageCible = "ANSI"
'    End If
'
'    If optWindows.Value Then
'        formatageCible = "Windows"
'    Else
'        formatageCible = "Unix"
'    End If
'
'    If optTxt.Value Then
'        extensionCible = "txt"
'    Else
'        extensionCible = "csv"
'    End If
'
'    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
'    With fdlgDossier
'        .Title = "Sélectionner le dossier de sauvegarde"
'        .AllowMultiSelect = False
'        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
'    End With
'
'    If fdlgDossier.Show <> -1 Then
'        MsgBox "Sélection annulée.", vbInformation
'        Exit Sub
'    End If
'
'    dossierSauvegarde = fdlgDossier.SelectedItems(1)
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    nomFichierSansExt = fso.GetBaseName(CheminFichier)
'
'    If Right(dossierSauvegarde, 1) <> "\" Then
'        dossierSauvegarde = dossierSauvegarde & "\"
'    End If
'
'    fichierSortie = dossierSauvegarde & nomFichierSansExt & "_converti." & extensionCible
'
'    If Dir(fichierSortie) <> "" Then
'        If MsgBox("Le fichier existe déjà :" & vbCrLf & fichierSortie & vbCrLf & vbCrLf & _
'                  "Voulez-vous le remplacer ?", vbYesNo + vbQuestion) = vbNo Then
'            Exit Sub
'        End If
'    End If
'
'    ConvertirFichierComplet CheminFichier, fichierSortie, encodageCible, formatageCible
'
'    MsgBox "Conversion terminée avec succès !" & vbCrLf & vbCrLf & _
'           "Fichier créé : " & fichierSortie, vbInformation
'
'    Unload Me
'    Exit Sub
'
'ErreurConversion:
'    MsgBox "Erreur lors de la conversion : " & Err.Description, vbCritical
'End Sub
'
'Private Sub ConvertirFichierComplet(source As String, destination As String, _
'                                   encodage As String, formatage As String)
'    Dim fso As Object
'    Dim tsInput As Object
'    Dim tsOutput As Object
'    Dim contenu As String
'    Dim encodageInput As Integer
'    Dim encodageOutput As Integer
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    If DetecterEncodageAmeliore(source) Like "UTF-8*" Then
'        encodageInput = -1
'    Else
'        encodageInput = 0
'    End If
'
'    Set tsInput = fso.OpenTextFile(source, 1, False, encodageInput)
'    contenu = tsInput.ReadAll
'    tsInput.Close
'
'    If formatage = "Windows" Then
'        contenu = Replace(contenu, vbCrLf, vbLf)
'        contenu = Replace(contenu, vbCr, vbLf)
'        contenu = Replace(contenu, vbLf, vbCrLf)
'    Else
'        contenu = Replace(contenu, vbCrLf, vbLf)
'        contenu = Replace(contenu, vbCr, vbLf)
'    End If
'
'    If encodage = "UTF-8" Then
'        encodageOutput = -1
'    Else
'        encodageOutput = 0
'    End If
'
'    Set tsOutput = fso.OpenTextFile(destination, 2, True, encodageOutput)
'    tsOutput.Write contenu
'    tsOutput.Close
'
'    Set fso = Nothing
'End Sub
'
'Private Sub btnAnnuler_Click()
'    Unload Me
'End Sub




' ========================================
' USERFORM frmConversion - Code amélioré
' ========================================

Option Explicit

Public CheminFichier As String
Private encodageDetecte As String
Private formatageDetecte As String
Private extensionDetectee As String



Private Sub Label2_Click()

End Sub

Private Sub FrameExtension_Click()

End Sub

Private Sub FrameFormatage_Click()

End Sub

Private Sub optANSI_Click()

End Sub

Private Sub optWindows_Click()

End Sub

Private Sub UserForm_Initialize()
    ' Initialisation du formulaire
    Me.Caption = "Conversion de fichier"

    ' Améliorer l'apparence des frames
    With Me.FrameEncodage
        .BackColor = RGB(230, 245, 230)
        .BorderColor = RGB(0, 120, 100)
        .SpecialEffect = fmSpecialEffectRaised
    End With

    With Me.FrameFormatage
        .BackColor = RGB(255, 250, 220)
        .BorderColor = RGB(0, 120, 100)
        .SpecialEffect = fmSpecialEffectRaised
    End With

    With Me.FrameExtension
        .BackColor = RGB(255, 240, 230)
        .BorderColor = RGB(0, 120, 100)
        .SpecialEffect = fmSpecialEffectRaised
    End With

    ' Améliorer les boutons radio - les rendre plus visibles
    AmeliorerBoutonsRadio
End Sub

Private Sub AmeliorerBoutonsRadio()
    ' Augmenter la taille et améliorer l'apparence des boutons radio
    Dim ctrl As Control

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" Then
            With ctrl
                .Font.Size = 18
                .Font.Bold = True
                .Height = 48
                .Width = 150
                ' Espacement pour meilleure lisibilité
'                Select Case ctrl.Name
'                    Case "optUTF8"
'                        .Left = 20
'                        .Top = 30
'                    Case "optANSI"
'                        .Left = 180
'                        .Top = 30
'                    Case "optWindows"
'                        .Left = 20
'                        .Top = 30
'                    Case "optUnix"
'                        .Left = 180
'                        .Top = 30
'                    Case "optTxt"
'                        .Left = 20
'                        .Top = 30
'                    Case "optCsv"
'                        .Left = 180
'                        .Top = 30
'                End Select
            End With
        End If
    Next ctrl
End Sub

Public Sub InitialiserAvecFichier(chemin As String)
    CheminFichier = chemin
    DetecterParametresFichier
End Sub

Public Sub DetecterParametresFichier()
    Dim fso As Object

    On Error GoTo ErreurDetection

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Détecter l'extension
    extensionDetectee = UCase(fso.GetExtensionName(CheminFichier))

    ' Détecter l'encodage (fonction améliorée)
    encodageDetecte = DetecterEncodage(CheminFichier)

    ' Détecter le formatage
    formatageDetecte = DetecterFormatage(CheminFichier)

    ' Afficher les informations avec style amélioré
    With lblFichierActuel
        .Caption = "Fichier : " & fso.GetFileName(CheminFichier)
'        .Font.Size = 11
'        .Font.Bold = True
        .ForeColor = RGB(0, 51, 102) ' Bleu foncé
        .BackColor = RGB(220, 240, 255) ' Bleu très clair
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(100, 150, 200)
        .SpecialEffect = fmSpecialEffectFlat
        '.TextAlign = fmTextAlignLeft
    End With

'    With lblEncodageActuel
'        .Caption = "Encodage actuel : " & encodageDetecte
''        .Font.Size = 10
''        .Font.Bold = False
'        .ForeColor = RGB(0, 51, 102)
'        .BackColor = RGB(230, 245, 255) ' Bleu clair dégradé
'        .BorderStyle = fmBorderStyleSingle
'        .BorderColor = RGB(100, 150, 200)
'        .SpecialEffect = fmSpecialEffectFlat
''        ' Couleur spéciale si UTF-8
''        If UCase(encodageDetecte) Like "UTF-8*" Then
''            .BackColor = RGB(200, 255, 200) ' Vert clair pour UTF-8
''            .ForeColor = RGB(0, 100, 0) ' Vert foncé
''        End If
'    End With

    With lblFormatageActuel
        .Caption = "Formatage actuel : " & formatageDetecte
'        .Font.Size = 10
'        .Font.Bold = False
        .ForeColor = RGB(0, 51, 102)
        .BackColor = RGB(240, 248, 255) ' Bleu clair alice
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(100, 150, 200)
        .SpecialEffect = fmSpecialEffectFlat
    End With

    With lblExtensionActuelle
        .Caption = "Extension actuelle : " & extensionDetectee
'        .Font.Size = 10
'        .Font.Bold = False
        .ForeColor = RGB(0, 51, 102)
        .BackColor = RGB(235, 245, 255) ' Bleu azur pâle
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(100, 150, 200)
        .SpecialEffect = fmSpecialEffectFlat
    End With

    ' Cocher les options correspondantes
    Select Case UCase(encodageDetecte)
        Case "UTF8", "UTF-8 BOM", "UTF8"
            optUTF8.Value = True
        Case "ANSI"
            optANSI.Value = True
        Case Else
            optUTF8.Value = True
    End Select

    Select Case UCase(formatageDetecte)
        Case "WINDOWS"
            optWindows.Value = True
        Case "UNIX"
'            optUnix.Value = True
        Case Else
            optWindows.Value = True
    End Select

    Select Case UCase(extensionDetectee)
        Case "TXT"
            optTxt.Value = True
        Case "CSV"
            optCsv.Value = True
        Case Else
            optTxt.Value = True
    End Select

    Exit Sub

ErreurDetection:
    MsgBox "Erreur lors de la détection : " & Err.Description, vbExclamation
    optUTF8.Value = True
    optWindows.Value = True
    optTxt.Value = True
End Sub
Private Function DetecterEncodage(fichier As String) As String
    Dim fichierNum As Integer
    Dim bom(1 To 3) As Byte
    Dim i As Integer

    On Error GoTo ErreurEncodage

    fichierNum = FreeFile
    Open fichier For Binary Access Read As #fichierNum

    ' Lire les 3 premiers octets pour détecter le BOM UTF-8
    If LOF(fichierNum) >= 3 Then
        For i = 1 To 3
            Get #fichierNum, , bom(i)
        Next i

        ' BOM UTF-8 : EF BB BF
        If bom(1) = &HEF And bom(2) = &HBB And bom(3) = &HBF Then
            DetecterEncodage = "UTF-8"
            Close #fichierNum
            Exit Function
        End If
    End If

    Close #fichierNum

    ' Si pas de BOM, considérer comme ANSI
    Dim contenu As String, b As Byte, isPureAscii As Boolean
    isPureAscii = True
    
    Do While Not EOF(fichierNum)
        Get #fichierNum, , b
        If b > 127 Then isPureAscii = False: Exit Do
    Loop
    Close #fichierNum
    
    If isPureAscii Then
        DetecterEncodage = "ANSI"
    Else
        DetecterEncodage = "UTF-8"  ' UTF-8 sans BOM probable
    End If


'    DetecterEncodage = "ANSI"
'    Exit Function
'Private Function DetecterEncodageAmeliore(fichier As String) As String
'    ' Fonction ultra-robuste pour détecter UTF-8
'    Dim adoStream As Object
'    Dim contenu As String
'    Dim fichierNum As Integer
'    Dim bom(1 To 3) As Byte
'    Dim i As Long
'    Dim countUTF8 As Long
'    Dim countErrors As Long
'    Dim b As Long
'
'    On Error GoTo ErreurEncodage
'
'    ' Méthode 1 : Vérifier le BOM
'    fichierNum = FreeFile
'    Open fichier For Binary Access Read As #fichierNum
'
'    If LOF(fichierNum) >= 3 Then
'        Get #fichierNum, , bom(1)
'        Get #fichierNum, , bom(2)
'        Get #fichierNum, , bom(3)
'
'        If bom(1) = &HEF And bom(2) = &HBB And bom(3) = &HBF Then
'            Close #fichierNum
'            DetecterEncodageAmeliore = "UTF-8"
'            Exit Function
'        End If
'    End If
'    Close #fichierNum
'
''    ' Méthode 2 : Utiliser ADODB.Stream pour tester la lecture UTF-8
''    On Error Resume Next
''    Set adoStream = CreateObject("ADODB.Stream")
''    If Err.Number <> 0 Then
''        On Error GoTo ErreurEncodage
''        GoTo MethodeAlternative
''    End If
''    On Error GoTo ErreurEncodage
''
''    With adoStream
''        .Type = 2 ' adTypeText
''        .Charset = "UTF-8"
''        .Open
''        .LoadFromFile fichier
''
''        ' Si on peut lire sans erreur, c'est probablement UTF-8
''        If Not .EOS Then
''            contenu = .ReadText(1000) ' Lire un échantillon
''
''            ' Vérifier si le contenu contient des caractères non-ASCII
''            Dim hasSpecialChars As Boolean
''            hasSpecialChars = False
''
''            For i = 1 To Len(contenu)
''                If AscW(Mid(contenu, i, 1)) > 127 Then
''                    hasSpecialChars = True
''                    Exit For
''                End If
''            Next i
''
''            .Close
''
''            ' Si caractères spéciaux détectés avec succès, c'est UTF-8
''            If hasSpecialChars Then
''                DetecterEncodageAmeliore = "UTF-8"
''                Set adoStream = Nothing
''                Exit Function
''            End If
''        Else
''            .Close
''        End If
''    End With
''    Set adoStream = Nothing
''
''MethodeAlternative:
''    ' Méthode 3 : Analyse heuristique des octets
''    Dim fso As Object
''    Dim ts As Object
''    Dim ligne As String
''    Dim utf8Score As Long
''    Dim ansiScore As Long
''    Dim ligneCount As Long
''
''    Set fso = CreateObject("Scripting.FileSystemObject")
''
''    ' Essayer d'ouvrir en UTF-8
''    On Error Resume Next
''    Set ts = fso.OpenTextFile(fichier, 1, False, -1) ' -1 = UTF-8
''    If Err.Number = 0 Then
''        ligneCount = 0
''        Do While Not ts.AtEndOfStream And ligneCount < 50
''            ligne = ts.ReadLine
''            ' Vérifier la présence de caractères accentués français communs
''            If InStr(1, ligne, "é") > 0 Or InStr(1, ligne, "è") > 0 Or _
''               InStr(1, ligne, "à") > 0 Or InStr(1, ligne, "ç") > 0 Or _
''               InStr(1, ligne, "ê") > 0 Or InStr(1, ligne, "â") > 0 Then
''                utf8Score = utf8Score + 1
''            End If
''            ' Vérifier si pas de caractères bizarres (mojibake)
''            If InStr(1, ligne, "Ã©") = 0 And InStr(1, ligne, "Ã¨") = 0 And _
''               InStr(1, ligne, "Ã ") = 0 Then
''                utf8Score = utf8Score + 1
''            Else
''                ansiScore = ansiScore + 5
''            End If
''            ligneCount = ligneCount + 1
''        Loop
''        ts.Close
''    End If
''    Err.Clear
''    On Error GoTo ErreurEncodage
''
''     'Essayer d 'ouvrir en ANSI
''    Set ts = fso.OpenTextFile(fichier, 1, False, 0) ' 0 = ANSI
''    If Not Err.Number Then
''        ligneCount = 0
''        Do While Not ts.AtEndOfStream And ligneCount < 50
''            ligne = ts.ReadLine
''            ' Vérifier caractères accentués corrects en ANSI
''            If InStr(1, ligne, "é") > 0 Or InStr(1, ligne, "è") > 0 Then
''                ansiScore = ansiScore + 1
''            End If
''            ligneCount = ligneCount + 1
''        Loop
''        ts.Close
''    End If
''
''    ' Décision basée sur les scores
''    If utf8Score > ansiScore Then
''        DetecterEncodageAmeliore = "UTF-8"
''    Else
''        DetecterEncodageAmeliore = "ANSI"
''    End If
''
'    Exit Function

ErreurEncodage:
    If fichierNum > 0 Then Close #fichierNum
    On Error Resume Next
'    If Not adoStream Is Nothing Then adoStream.Close
'    If Not ts Is Nothing Then ts.Close
    DetecterEncodage = "ANSI"
End Function

Private Function DetecterFormatage(fichier As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim contenu As String
    Dim posLF As Long
    Dim posCRLF As Long

    On Error GoTo ErreurFormatage

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(fichier, 1, False, -2)

    If Not ts.AtEndOfStream Then
        contenu = ts.Read(10000)
    End If
    ts.Close

    posCRLF = InStr(1, contenu, vbCrLf)
    posLF = InStr(1, contenu, vbLf)

    If posCRLF > 0 Then
        DetecterFormatage = "Windows"
    ElseIf posLF > 0 Then
        DetecterFormatage = "Unix"
    Else
        DetecterFormatage = "Windows"
    End If

    Exit Function

ErreurFormatage:
    DetecterFormatage = "Windows"
End Function

Private Sub btnConvertir_Click()
    Dim encodageCible As String
    Dim formatageCible As String
    Dim extensionCible As String
    Dim fdlgDossier As FileDialog
    Dim dossierSauvegarde As String
    Dim fso As Object
    Dim nomFichierSansExt As String
    Dim fichierSortie As String

    On Error GoTo ErreurConversion

    If optUTF8.Value Then
        encodageCible = "UTF-8"
    Else
        encodageCible = "ANSI"
    End If

    If optWindows.Value Then
        formatageCible = "Windows"
    Else
        formatageCible = "Unix"
    End If

    If optTxt.Value Then
        extensionCible = "txt"
    Else
        extensionCible = "csv"
    End If

    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlgDossier
        .Title = "Sélectionner le dossier de sauvegarde"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
    End With

    If fdlgDossier.Show <> -1 Then
        MsgBox "Sélection annulée.", vbInformation
        Exit Sub
    End If

    dossierSauvegarde = fdlgDossier.SelectedItems(1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    nomFichierSansExt = fso.GetBaseName(CheminFichier)

    If Right(dossierSauvegarde, 1) <> "\" Then
        dossierSauvegarde = dossierSauvegarde & "\"
    End If

    fichierSortie = dossierSauvegarde & nomFichierSansExt & "_converti." & extensionCible

    If Dir(fichierSortie) <> "" Then
        If MsgBox("Le fichier existe déjà :" & vbCrLf & fichierSortie & vbCrLf & vbCrLf & _
                  "Voulez-vous le remplacer ?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    ConvertirFichierComplet CheminFichier, fichierSortie, encodageCible, formatageCible

    MsgBox "Conversion terminée avec succès !" & vbCrLf & vbCrLf & _
           "Fichier créé : " & fichierSortie, vbInformation
      ' Ouvrir le dossier contenant le fichier converti
    Shell "explorer.exe /select,""" & fichierSortie & """", vbNormalFocus

    Unload Me
    Exit Sub

ErreurConversion:
    MsgBox "Erreur lors de la conversion : " & Err.Description, vbCritical
End Sub

Private Sub ConvertirFichierCompletSSSSSSSSSSSSSSSSS(source As String, destination As String, _
                                   encodage As String, formatage As String)
    Dim adoStreamIn As Object
    Dim adoStreamOut As Object
    Dim contenu As String
    Dim charsetIn As String
    Dim charsetOut As String

    On Error GoTo ErreurConversion

    ' Déterminer les charsets
    If DetecterEncodage(source) Like "*UTF-8*" Or DetecterEncodage(source) = "UTF8" Then
        charsetIn = "UTF-8"
    Else
        charsetIn = "Windows-1252" ' ANSI occidental
    End If

    If encodage = "UTF-8" Then
        charsetOut = "UTF-8"
    Else
        charsetOut = "Windows-1252"
    End If

    ' Méthode ADODB.Stream (plus fiable pour UTF-8)
    Set adoStreamIn = CreateObject("ADODB.Stream")
    With adoStreamIn
        .Type = 2 ' adTypeText
        .Charset = charsetIn
        .Open
        .LoadFromFile source
        contenu = .ReadText
        .Close
    End With

    ' Appliquer le formatage
    If formatage = "Windows" Then
        contenu = Replace(contenu, vbCrLf, vbLf)
        contenu = Replace(contenu, vbCr, vbLf)
        contenu = Replace(contenu, vbLf, vbCrLf)
    Else
        contenu = Replace(contenu, vbCrLf, vbLf)
        contenu = Replace(contenu, vbCr, vbLf)
    End If

    ' Écrire le fichier de sortie
    Set adoStreamOut = CreateObject("ADODB.Stream")
    With adoStreamOut
        .Type = 2 ' adTypeText
        .Charset = charsetOut
        .Open
        .WriteText contenu
        .SaveToFile destination, 2 ' adSaveCreateOverWrite
        .Close
    End With

    Set adoStreamIn = Nothing
    Set adoStreamOut = Nothing
    Exit Sub

ErreurConversion:
    MsgBox "Erreur lors de la conversion : " & Err.Description & vbCrLf & _
           "Essai avec méthode alternative...", vbExclamation

    ' Méthode alternative si ADODB échoue
    Dim fso As Object
    Dim tsInput As Object
    Dim tsOutput As Object
    Dim encodageInput As Integer
    Dim encodageOutput As Integer

    On Error Resume Next
    If Not adoStreamIn Is Nothing Then adoStreamIn.Close
    If Not adoStreamOut Is Nothing Then adoStreamOut.Close
    On Error GoTo 0

    Set fso = CreateObject("Scripting.FileSystemObject")

    If charsetIn = "UTF-8" Then encodageInput = -1 Else encodageInput = 0
    If charsetOut = "UTF-8" Then encodageOutput = -1 Else encodageOutput = 0

    Set tsInput = fso.OpenTextFile(source, 1, False, encodageInput)
    contenu = tsInput.ReadAll
    tsInput.Close

    If formatage = "Windows" Then
        contenu = Replace(contenu, vbCrLf, vbLf)
        contenu = Replace(contenu, vbCr, vbLf)
        contenu = Replace(contenu, vbLf, vbCrLf)
    Else
        contenu = Replace(contenu, vbCrLf, vbLf)
        contenu = Replace(contenu, vbCr, vbLf)
    End If

    Set tsOutput = fso.OpenTextFile(destination, 2, True, encodageOutput)
    tsOutput.Write contenu
    tsOutput.Close

    Set fso = Nothing
End Sub

Private Sub ConvertirFichierComplet(source As String, destination As String, _
                                   encodage As String, formatage As String)
    Dim adoStreamIn As Object
    Dim adoStreamOut As Object
    Dim contenu As String
    Dim charsetIn As String
    Dim charsetOut As String

    On Error GoTo ErreurConversion

    ' Déterminer les charsets
    If DetecterEncodage(source) Like "*UTF-8*" Or DetecterEncodage(source) = "UTF8" Then
        charsetIn = "UTF-8"
    Else
        charsetIn = "Windows-1252" ' ANSI occidental
    End If

    If encodage = "UTF-8" Then
        charsetOut = "UTF-8"
    Else
        charsetOut = "Windows-1252"
    End If

    ' Méthode ADODB.Stream (plus fiable pour UTF-8)
    Set adoStreamIn = CreateObject("ADODB.Stream")
    With adoStreamIn
        .Type = 2 ' adTypeText
        .Charset = charsetIn
        .Open
        .LoadFromFile source
        contenu = .ReadText
        .Close
    End With

    ' Appliquer le formatage
    If formatage = "Windows" Then
        contenu = Replace(contenu, vbCrLf, vbLf)
        contenu = Replace(contenu, vbCr, vbLf)
        contenu = Replace(contenu, vbLf, vbCrLf)
    Else
        contenu = Replace(contenu, vbCrLf, vbLf)
        contenu = Replace(contenu, vbCr, vbLf)
    End If

    ' Écrire le fichier de sortie avec UTF-8 SANS BOM
    Set adoStreamOut = CreateObject("ADODB.Stream")
    With adoStreamOut
        .Type = 2 ' adTypeText
        .Charset = charsetOut
        .Open
        .WriteText contenu

        ' Si UTF-8, sauvegarder SANS BOM
        If charsetOut = "UTF-8" Then
            .Position = 0
            .Type = 1 ' adTypeBinary
            .Position = 3 ' Sauter le BOM (3 octets : EF BB BF)

            Dim binaryData As Variant
            binaryData = .Read

            .Close
            .Open
            .Type = 1 ' adTypeBinary
            .Write binaryData
        End If

        .SaveToFile destination, 2 ' adSaveCreateOverWrite
        .Close
    End With

    Set adoStreamIn = Nothing
    Set adoStreamOut = Nothing
    Exit Sub

ErreurConversion:
    MsgBox "Erreur lors de la conversion : " & Err.Description & vbCrLf & _
           "Essai avec méthode alternative...", vbExclamation

    ' Méthode alternative si ADODB échoue
    Dim fso As Object
    Dim tsInput As Object
    Dim tsOutput As Object
    Dim encodageInput As Integer
    Dim encodageOutput As Integer

    On Error Resume Next
    If Not adoStreamIn Is Nothing Then adoStreamIn.Close
    If Not adoStreamOut Is Nothing Then adoStreamOut.Close
    On Error GoTo 0

    Set fso = CreateObject("Scripting.FileSystemObject")

    If charsetIn = "UTF-8" Then encodageInput = -1 Else encodageInput = 0
    If charsetOut = "UTF-8" Then encodageOutput = -1 Else encodageOutput = 0

    Set tsInput = fso.OpenTextFile(source, 1, False, encodageInput)
    contenu = tsInput.ReadAll
    tsInput.Close

    If formatage = "Windows" Then
        contenu = Replace(contenu, vbCrLf, vbLf)
        contenu = Replace(contenu, vbCr, vbLf)
        contenu = Replace(contenu, vbLf, vbCrLf)
    Else
        contenu = Replace(contenu, vbCrLf, vbLf)
        contenu = Replace(contenu, vbCr, vbLf)
    End If

    Set tsOutput = fso.OpenTextFile(destination, 2, True, encodageOutput)
    tsOutput.Write contenu
    tsOutput.Close

    Set fso = Nothing
End Sub

Private Sub btnAnnuler_Click()
    Unload Me
End Sub









'' ========================================
'' USERFORM frmConversion - Code amélioré
'' ========================================
'
'Option Explicit
'
'Public CheminFichier As String
'Private encodageDetecte As String
'Private formatageDetecte As String
'Private extensionDetectee As String
'
'Private Sub UserForm_Initialize()
'    ' Initialisation du formulaire
'    Me.Caption = "Conversion de fichier"
'
'    ' Améliorer l'apparence des frames
'    With Me.FrameEncodage
'        .BackColor = RGB(230, 245, 230)
'        .BorderColor = RGB(0, 120, 100)
'        .SpecialEffect = fmSpecialEffectRaised
'    End With
'
'    With Me.FrameFormatage
'        .BackColor = RGB(255, 250, 220)
'        .BorderColor = RGB(0, 120, 100)
'        .SpecialEffect = fmSpecialEffectRaised
'    End With
'
'    With Me.FrameExtension
'        .BackColor = RGB(255, 240, 230)
'        .BorderColor = RGB(0, 120, 100)
'        .SpecialEffect = fmSpecialEffectRaised
'    End With
'
'    ' Améliorer les boutons radio - les rendre plus visibles
'    AmeliorerBoutonsRadio
'
'    ' IMPORTANT : Définir les GroupName pour que les boutons fonctionnent correctement
'    optUTF8.GroupName = "Encodage"
'    optANSI.GroupName = "Encodage"
'
'    optWindows.GroupName = "Formatage"
'    optUnix.GroupName = "Formatage"
'
'    optTxt.GroupName = "Extension"
'    optCsv.GroupName = "Extension"
'End Sub
'
'Private Sub AmeliorerBoutonsRadio()
'    ' Augmenter la taille et améliorer l'apparence des boutons radio
'    Dim ctrl As Control
'
'    For Each ctrl In Me.Controls
'        If TypeName(ctrl) = "OptionButton" Then
'            With ctrl
'                .Font.Size = 12
'                .Font.Bold = True
'                .Height = 24
'                .Width = 140
'                ' Espacement pour meilleure lisibilité
'                Select Case ctrl.Name
'                    Case "optUTF8"
'                        .Left = 20
'                        .Top = 30
'                    Case "optANSI"
'                        .Left = 180
'                        .Top = 30
'                    Case "optWindows"
'                        .Left = 20
'                        .Top = 30
'                    Case "optUnix"
'                        .Left = 180
'                        .Top = 30
'                    Case "optTxt"
'                        .Left = 20
'                        .Top = 30
'                    Case "optCsv"
'                        .Left = 180
'                        .Top = 30
'                End Select
'            End With
'        End If
'    Next ctrl
'End Sub
'
'Public Sub InitialiserAvecFichier(chemin As String)
'    CheminFichier = chemin
'    DetecterParametresFichier
'End Sub
'
'Public Sub DetecterParametresFichier()
'    Dim fso As Object
'
'    On Error GoTo ErreurDetection
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' Détecter l'extension
'    extensionDetectee = UCase(fso.GetExtensionName(CheminFichier))
'
'    ' Détecter l'encodage (fonction améliorée)
'    encodageDetecte = DetecterEncodageAmeliore(CheminFichier)
'
'    ' Détecter le formatage
'    formatageDetecte = DetecterFormatage(CheminFichier)
'
'    ' Afficher les informations avec style amélioré - Dégradé de bleu
'    With lblFichierActuel
'        .Caption = "Fichier : " & fso.GetFileName(CheminFichier)
'        .Font.Size = 11
'        .Font.Bold = True
'        .ForeColor = RGB(25, 25, 112) ' MidnightBlue - Bleu très foncé
'        .BackColor = RGB(176, 196, 222) ' LightSteelBlue - Bleu moyen foncé
'        .BorderStyle = fmBorderStyleSingle
'        .BorderColor = RGB(70, 130, 180) ' SteelBlue
'        .SpecialEffect = fmSpecialEffectFlat
'        .TextAlign = fmTextAlignLeft
'    End With
'
'    With lblEncodageActuel
'        .Caption = "Encodage actuel : " & encodageDetecte
'        .Font.Size = 10
'        .Font.Bold = False
'        .ForeColor = RGB(25, 25, 112) ' MidnightBlue
'        .BackColor = RGB(176, 224, 230) ' PowderBlue - Bleu moyen
'        .BorderStyle = fmBorderStyleSingle
'        .BorderColor = RGB(70, 130, 180)
'        .SpecialEffect = fmSpecialEffectFlat
'    End With
'
'    With lblFormatageActuel
'        .Caption = "Formatage actuel : " & formatageDetecte
'        .Font.Size = 10
'        .Font.Bold = False
'        .ForeColor = RGB(25, 25, 112) ' MidnightBlue
'        .BackColor = RGB(173, 216, 230) ' LightBlue - Bleu moyen clair
'        .BorderStyle = fmBorderStyleSingle
'        .BorderColor = RGB(70, 130, 180)
'        .SpecialEffect = fmSpecialEffectFlat
'    End With
'
'    With lblExtensionActuelle
'        .Caption = "Extension actuelle : " & extensionDetectee
'        .Font.Size = 10
'        .Font.Bold = False
'        .ForeColor = RGB(25, 25, 112) ' MidnightBlue
'        .BackColor = RGB(224, 255, 255) ' LightCyan - Bleu très clair
'        .BorderStyle = fmBorderStyleSingle
'        .BorderColor = RGB(70, 130, 180)
'        .SpecialEffect = fmSpecialEffectFlat
'    End With
'
'    ' Cocher les options correspondantes
'    Select Case UCase(encodageDetecte)
'        Case "UTF-8", "UTF-8 BOM", "UTF8"
'            optUTF8.Value = True
'        Case "ANSI"
'            optANSI.Value = True
'        Case Else
'            optUTF8.Value = True
'    End Select
'
'    Select Case UCase(formatageDetecte)
'        Case "WINDOWS"
'            optWindows.Value = True
'        Case "UNIX"
'            optUnix.Value = True
'        Case Else
'            optWindows.Value = True
'    End Select
'
'    Select Case UCase(extensionDetectee)
'        Case "TXT"
'            optTxt.Value = True
'        Case "CSV"
'            optCsv.Value = True
'        Case Else
'            optTxt.Value = True
'    End Select
'
'    Exit Sub
'
'ErreurDetection:
'    MsgBox "Erreur lors de la détection : " & Err.Description, vbExclamation
'    optUTF8.Value = True
'    optWindows.Value = True
'    optTxt.Value = True
'End Sub
'
'Private Function DetecterEncodageAmeliore(fichier As String) As String
'    ' Fonction ultra-robuste pour détecter UTF-8
'    Dim adoStream As Object
'    Dim contenu As String
'    Dim fichierNum As Integer
'    Dim bom(1 To 3) As Byte
'    Dim i As Long
'    Dim countUTF8 As Long
'    Dim countErrors As Long
'    Dim b As Long
'
'    On Error GoTo ErreurEncodage
'
'    ' Méthode 1 : Vérifier le BOM
'    fichierNum = FreeFile
'    Open fichier For Binary Access Read As #fichierNum
'
'    If LOF(fichierNum) >= 3 Then
'        Get #fichierNum, , bom(1)
'        Get #fichierNum, , bom(2)
'        Get #fichierNum, , bom(3)
'
'        If bom(1) = &HEF And bom(2) = &HBB And bom(3) = &HBF Then
'            Close #fichierNum
'            DetecterEncodageAmeliore = "UTF-8"
'            Exit Function
'        End If
'    End If
'    Close #fichierNum
'
'    ' Méthode 2 : Utiliser ADODB.Stream pour tester la lecture UTF-8
'    On Error Resume Next
'    Set adoStream = CreateObject("ADODB.Stream")
'    If Err.Number <> 0 Then
'        On Error GoTo ErreurEncodage
'        GoTo MethodeAlternative
'    End If
'    On Error GoTo ErreurEncodage
'
'    With adoStream
'        .Type = 2 ' adTypeText
'        .Charset = "UTF-8"
'        .Open
'        .LoadFromFile fichier
'
'        ' Si on peut lire sans erreur, c'est probablement UTF-8
'        If Not .EOS Then
'            contenu = .ReadText(1000) ' Lire un échantillon
'
'            ' Vérifier si le contenu contient des caractères non-ASCII
'            Dim hasSpecialChars As Boolean
'            hasSpecialChars = False
'
'            For i = 1 To Len(contenu)
'                If AscW(Mid(contenu, i, 1)) > 127 Then
'                    hasSpecialChars = True
'                    Exit For
'                End If
'            Next i
'
'            .Close
'
'            ' Si caractères spéciaux détectés avec succès, c'est UTF-8
'            If hasSpecialChars Then
'                DetecterEncodageAmeliore = "UTF-8"
'                Set adoStream = Nothing
'                Exit Function
'            End If
'        Else
'            .Close
'        End If
'    End With
'    Set adoStream = Nothing
'
'MethodeAlternative:
'    ' Méthode 3 : Analyse heuristique des octets
'    Dim fso As Object
'    Dim ts As Object
'    Dim ligne As String
'    Dim utf8Score As Long
'    Dim ansiScore As Long
'    Dim ligneCount As Long
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' Essayer d'ouvrir en UTF-8
'    On Error Resume Next
'    Set ts = fso.OpenTextFile(fichier, 1, False, -1) ' -1 = UTF-8
'    If Err.Number = 0 Then
'        ligneCount = 0
'        Do While Not ts.AtEndOfStream And ligneCount < 50
'            ligne = ts.ReadLine
'            ' Vérifier la présence de caractères accentués français communs
'            If InStr(1, ligne, "é") > 0 Or InStr(1, ligne, "è") > 0 Or _
'               InStr(1, ligne, "à") > 0 Or InStr(1, ligne, "ç") > 0 Or _
'               InStr(1, ligne, "ê") > 0 Or InStr(1, ligne, "â") > 0 Then
'                utf8Score = utf8Score + 1
'            End If
'            ' Vérifier si pas de caractères bizarres (mojibake)
'            If InStr(1, ligne, "Ã©") = 0 And InStr(1, ligne, "Ã¨") = 0 And _
'               InStr(1, ligne, "Ã ") = 0 Then
'                utf8Score = utf8Score + 1
'            Else
'                ansiScore = ansiScore + 5
'            End If
'            ligneCount = ligneCount + 1
'        Loop
'        ts.Close
'    End If
'    Err.Clear
'    On Error GoTo ErreurEncodage
'
'    ' Essayer d'ouvrir en ANSI
'    Set ts = fso.OpenTextFile(fichier, 1, False, 0) ' 0 = ANSI
'    If Not Err.Number Then
'        ligneCount = 0
'        Do While Not ts.AtEndOfStream And ligneCount < 50
'            ligne = ts.ReadLine
'            ' Vérifier caractères accentués corrects en ANSI
'            If InStr(1, ligne, "é") > 0 Or InStr(1, ligne, "è") > 0 Then
'                ansiScore = ansiScore + 1
'            End If
'            ligneCount = ligneCount + 1
'        Loop
'        ts.Close
'    End If
'
'    ' Décision basée sur les scores
'    If utf8Score > ansiScore Then
'        DetecterEncodageAmeliore = "UTF-8"
'    Else
'        DetecterEncodageAmeliore = "ANSI"
'    End If
'
'    Exit Function
'
'ErreurEncodage:
'    If fichierNum > 0 Then Close #fichierNum
'    On Error Resume Next
'    If Not adoStream Is Nothing Then adoStream.Close
'    If Not ts Is Nothing Then ts.Close
'    DetecterEncodageAmeliore = "ANSI"
'End Function
'
'Private Function DetecterFormatage(fichier As String) As String
'    Dim fso As Object
'    Dim ts As Object
'    Dim contenu As String
'    Dim posLF As Long
'    Dim posCRLF As Long
'
'    On Error GoTo ErreurFormatage
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set ts = fso.OpenTextFile(fichier, 1, False, -2)
'
'    If Not ts.AtEndOfStream Then
'        contenu = ts.Read(10000)
'    End If
'    ts.Close
'
'    posCRLF = InStr(1, contenu, vbCrLf)
'    posLF = InStr(1, contenu, vbLf)
'
'    If posCRLF > 0 Then
'        DetecterFormatage = "Windows"
'    ElseIf posLF > 0 Then
'        DetecterFormatage = "Unix"
'    Else
'        DetecterFormatage = "Windows"
'    End If
'
'    Exit Function
'
'ErreurFormatage:
'    DetecterFormatage = "Windows"
'End Function
'
'Private Sub btnConvertir_Click()
'    Dim encodageCible As String
'    Dim formatageCible As String
'    Dim extensionCible As String
'    Dim fdlgDossier As FileDialog
'    Dim dossierSauvegarde As String
'    Dim fso As Object
'    Dim nomFichierSansExt As String
'    Dim fichierSortie As String
'
'    On Error GoTo ErreurConversion
'
'    If optUTF8.Value Then
'        encodageCible = "UTF-8"
'    Else
'        encodageCible = "ANSI"
'    End If
'
'    If optWindows.Value Then
'        formatageCible = "Windows"
'    Else
'        formatageCible = "Unix"
'    End If
'
'    If optTxt.Value Then
'        extensionCible = "txt"
'    Else
'        extensionCible = "csv"
'    End If
'
'    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
'    With fdlgDossier
'        .Title = "Sélectionner le dossier de sauvegarde"
'        .AllowMultiSelect = False
'        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
'    End With
'
'    If fdlgDossier.Show <> -1 Then
'        MsgBox "Sélection annulée.", vbInformation
'        Exit Sub
'    End If
'
'    dossierSauvegarde = fdlgDossier.SelectedItems(1)
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    nomFichierSansExt = fso.GetBaseName(CheminFichier)
'
'    If Right(dossierSauvegarde, 1) <> "\" Then
'        dossierSauvegarde = dossierSauvegarde & "\"
'    End If
'
'    fichierSortie = dossierSauvegarde & nomFichierSansExt & "_converti." & extensionCible
'
'    If Dir(fichierSortie) <> "" Then
'        If MsgBox("Le fichier existe déjà :" & vbCrLf & fichierSortie & vbCrLf & vbCrLf & _
'                  "Voulez-vous le remplacer ?", vbYesNo + vbQuestion) = vbNo Then
'            Exit Sub
'        End If
'    End If
'
'    ConvertirFichierComplet CheminFichier, fichierSortie, encodageCible, formatageCible
'
'    MsgBox "Conversion terminée avec succès !" & vbCrLf & vbCrLf & _
'           "Fichier créé : " & fichierSortie, vbInformation
'
'    Unload Me
'    Exit Sub
'
'ErreurConversion:
'    MsgBox "Erreur lors de la conversion : " & Err.Description, vbCritical
'End Sub
'
'Private Sub ConvertirFichierComplet(source As String, destination As String, _
'                                   encodage As String, formatage As String)
'    Dim adoStreamIn As Object
'    Dim adoStreamOut As Object
'    Dim contenu As String
'    Dim charsetIn As String
'    Dim charsetOut As String
'
'    On Error GoTo ErreurConversion
'
'    ' Déterminer les charsets
'    If DetecterEncodageAmeliore(source) Like "*UTF-8*" Or DetecterEncodageAmeliore(source) = "UTF8" Then
'        charsetIn = "UTF-8"
'    Else
'        charsetIn = "Windows-1252" ' ANSI occidental
'    End If
'
'    If encodage = "UTF-8" Then
'        charsetOut = "UTF-8"
'    Else
'        charsetOut = "Windows-1252"
'    End If
'
'    ' Méthode ADODB.Stream (plus fiable pour UTF-8)
'    Set adoStreamIn = CreateObject("ADODB.Stream")
'    With adoStreamIn
'        .Type = 2 ' adTypeText
'        .Charset = charsetIn
'        .Open
'        .LoadFromFile source
'        contenu = .ReadText
'        .Close
'    End With
'
'    ' Appliquer le formatage
'    If formatage = "Windows" Then
'        contenu = Replace(contenu, vbCrLf, vbLf)
'        contenu = Replace(contenu, vbCr, vbLf)
'        contenu = Replace(contenu, vbLf, vbCrLf)
'    Else
'        contenu = Replace(contenu, vbCrLf, vbLf)
'        contenu = Replace(contenu, vbCr, vbLf)
'    End If
'
'    ' Écrire le fichier de sortie avec UTF-8 SANS BOM
'    Set adoStreamOut = CreateObject("ADODB.Stream")
'    With adoStreamOut
'        .Type = 2 ' adTypeText
'        .Charset = charsetOut
'        .Open
'        .WriteText contenu
'
'        ' Si UTF-8, sauvegarder SANS BOM
'        If charsetOut = "UTF-8" Then
'            .Position = 0
'            .Type = 1 ' adTypeBinary
'            .Position = 3 ' Sauter le BOM (3 octets : EF BB BF)
'
'            Dim binaryData As Variant
'            binaryData = .Read
'
'            .Close
'            .Open
'            .Type = 1 ' adTypeBinary
'            .Write binaryData
'        End If
'
'        .SaveToFile destination, 2 ' adSaveCreateOverWrite
'        .Close
'    End With
'
'    Set adoStreamIn = Nothing
'    Set adoStreamOut = Nothing
'    Exit Sub
'
'ErreurConversion:
'    MsgBox "Erreur lors de la conversion : " & Err.Description & vbCrLf & _
'           "Essai avec méthode alternative...", vbExclamation
'
'    ' Méthode alternative si ADODB échoue
'    Dim fso As Object
'    Dim tsInput As Object
'    Dim tsOutput As Object
'    Dim encodageInput As Integer
'    Dim encodageOutput As Integer
'
'    On Error Resume Next
'    If Not adoStreamIn Is Nothing Then adoStreamIn.Close
'    If Not adoStreamOut Is Nothing Then adoStreamOut.Close
'    On Error GoTo 0
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    If charsetIn = "UTF-8" Then encodageInput = -1 Else encodageInput = 0
'    If charsetOut = "UTF-8" Then encodageOutput = -1 Else encodageOutput = 0
'
'    Set tsInput = fso.OpenTextFile(source, 1, False, encodageInput)
'    contenu = tsInput.ReadAll
'    tsInput.Close
'
'    If formatage = "Windows" Then
'        contenu = Replace(contenu, vbCrLf, vbLf)
'        contenu = Replace(contenu, vbCr, vbLf)
'        contenu = Replace(contenu, vbLf, vbCrLf)
'    Else
'        contenu = Replace(contenu, vbCrLf, vbLf)
'        contenu = Replace(contenu, vbCr, vbLf)
'    End If
'
'    Set tsOutput = fso.OpenTextFile(destination, 2, True, encodageOutput)
'    tsOutput.Write contenu
'    tsOutput.Close
'
'    Set fso = Nothing
'End Sub
'
'Private Sub btnAnnuler_Click()
'    Unload Me
'End Sub


