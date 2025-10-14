Attribute VB_Name = "UTF8ToANSI"
Sub ConvertUTF8ToANSI()
    Dim fso As Object
    Dim tsInput As Object
    Dim tsOutput As Object
    Dim inputFile As String
    Dim outputFile As String
    Dim fileContent As String

    ' Chemin du fichier source (UTF-8)
    inputFile = "C:\chemin\vers\fichier_utf8.txt"
    ' Chemin du fichier de sortie (ANSI)
    outputFile = "C:\chemin\vers\fichier_ansi.txt"

    ' Créer une instance de FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ouvrir le fichier UTF-8 en lecture
    Set tsInput = fso.OpenTextFile(inputFile, 1, False, -1) ' -1 pour UTF-8
    fileContent = tsInput.ReadAll
    tsInput.Close

    ' Ouvrir le fichier de sortie en écriture (ANSI)
    Set tsOutput = fso.OpenTextFile(outputFile, 2, True, 0) ' 0 pour ANSI
    tsOutput.Write fileContent
    tsOutput.Close

    MsgBox "Conversion terminée !", vbInformation
End Sub

