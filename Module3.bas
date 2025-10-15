Attribute VB_Name = "Module3"
''Cl
'
'
'
'Sub ConvertirFichier()
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Application.Calculation = xlCalculationManual
'    Application.DisplayAlerts = False
'
'    On Error GoTo GestionErreur
'
'    Dim fdlg As FileDialog
'    Dim inputFile As String
'
'    ' Sélection du fichier à convertir
'    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
'
'    With fdlg
'        .Title = "Sélectionner le fichier à convertir"
'        .Filters.Clear
'        .Filters.Add "Fichiers texte et CSV", "*.txt;*.csv", 1
'        .Filters.Add "Fichiers texte", "*.txt", 2
'        .Filters.Add "Fichiers CSV", "*.csv", 3
'        .Filters.Add "Tous les fichiers", "*.*", 4
'        .AllowMultiSelect = False
'    End With
'
'    If fdlg.Show <> -1 Then
'        MsgBox "Sélection annulée.", vbInformation
'        GoTo Fin
'    End If
'
'    If fdlg.SelectedItems.Count = 0 Then
'        MsgBox "Aucun fichier sélectionné.", vbInformation
'        GoTo Fin
'    End If
'
'    inputFile = fdlg.SelectedItems(1)
'
'    If Dir(inputFile) = "" Then
'        MsgBox "Le fichier n'existe pas : " & inputFile, vbCritical
'        GoTo Fin
'    End If
'
''        Dim uf As New frmConversion
''    uf.CheminFichier = inputFile        ' utilise la Property Let/Get du formulaire
''    uf.DetecterParametresFichier
''    uf.Show
'
'    ' Ouvrir le UserForm avec le fichier sélectionné
'    Load frmConversion
'    frmConversion.CheminFichier = inputFile
'    frmConversion.DetecterParametresFichier
'    frmConversion.Show
'
'    GoTo Fin
'
'GestionErreur:
'    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
'
'Fin:
'    Set fdlg = Nothing
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'    Application.DisplayAlerts = True
'End Sub




' ========================================
' MODULE PRINCIPAL
' ========================================

Sub ConvertirFichier()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    On Error GoTo GestionErreur
    
    Dim fdlg As FileDialog
    Dim inputFile As String
    
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    With fdlg
        .Title = "Sélectionner le fichier à convertir"
        .Filters.Clear
        .Filters.Add "Fichiers texte et CSV", "*.txt;*.csv", 1
        .Filters.Add "Fichiers texte", "*.txt", 2
        .Filters.Add "Fichiers CSV", "*.csv", 3
        .Filters.Add "Tous les fichiers", "*.*", 4
        .AllowMultiSelect = False
    End With
    
    If fdlg.Show <> -1 Then
        MsgBox "Sélection annulée.", vbInformation
        GoTo Fin
    End If
    
    If fdlg.SelectedItems.Count = 0 Then
        MsgBox "Aucun fichier sélectionné.", vbInformation
        GoTo Fin
    End If
    
    inputFile = fdlg.SelectedItems(1)
    
    If Dir(inputFile) = "" Then
        MsgBox "Le fichier n'existe pas : " & inputFile, vbCritical
        GoTo Fin
    End If
    
    ' Initialiser et afficher le formulaire
    frmConversion.InitialiserAvecFichier inputFile
    frmConversion.Show
    
    GoTo Fin
    
GestionErreur:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
    
Fin:
    Set fdlg = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

'' ========================================
'' MODULE PRINCIPAL
'' ========================================
'
'Sub ConvertirFichier()
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Application.Calculation = xlCalculationManual
'    Application.DisplayAlerts = False
'
'    On Error GoTo GestionErreur
'
'    Dim fdlg As FileDialog
'    Dim inputFile As String
'
'    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
'    With fdlg
'        .Title = "Sélectionner le fichier à convertir"
'        .Filters.Clear
'        .Filters.Add "Fichiers texte et CSV", "*.txt;*.csv", 1
'        .Filters.Add "Fichiers texte", "*.txt", 2
'        .Filters.Add "Fichiers CSV", "*.csv", 3
'        .Filters.Add "Tous les fichiers", "*.*", 4
'        .AllowMultiSelect = False
'    End With
'
'    If fdlg.Show <> -1 Then
'        MsgBox "Sélection annulée.", vbInformation
'        GoTo Fin
'    End If
'
'    If fdlg.SelectedItems.Count = 0 Then
'        MsgBox "Aucun fichier sélectionné.", vbInformation
'        GoTo Fin
'    End If
'
'    inputFile = fdlg.SelectedItems(1)
'
'    If Dir(inputFile) = "" Then
'        MsgBox "Le fichier n'existe pas : " & inputFile, vbCritical
'        GoTo Fin
'    End If
'
'    ' Initialiser et afficher le formulaire
'    frmConversion.InitialiserAvecFichier inputFile
'    frmConversion.Show
'
'    GoTo Fin
'
'GestionErreur:
'    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
'
'Fin:
'    Set fdlg = Nothing
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'    Application.DisplayAlerts = True
'End Sub

