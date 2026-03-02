' ****************************************************************************************************
' Macro pour la Génération du Dossier de Lancement avec Export dans un fichier ZIP
' de la mise en plan au format PDF, DXF et de la pièce au format STEP
' ****************************************************************************************************
' Auteur : Tristan JACQ
' Date : Mars 2026
' Version : 1.1
' ****************************************************************************************************

Sub main()

    Dim swApp                   As SldWorks.SldWorks
    Dim swModel                 As SldWorks.ModelDoc2
    Dim swDraw                  As SldWorks.DrawingDoc
    Dim swSheet                 As SldWorks.Sheet
    Dim swView                  As SldWorks.View
    Dim swActiveView            As SldWorks.View
    Dim bRet                    As Boolean
    Dim cusPropMgr              As SldWorks.CustomPropertyManager
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Vérification document
    If swModel Is Nothing Then Exit Sub
    If swModel.GetType <> swDocDRAWING Then
        MsgBox "Cette macro ne fonctionne que sur une mise en plan."
        Exit Sub
    End If

    Set swDraw = swModel
    Set swSheet = swDraw.GetCurrentSheet
    Set swActiveView = swDraw.ActiveDrawingView
    
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    Set swRefModel = swView.ReferencedDocument

    ' Extraction nom du fichier du nom de la feuille
    Dim NomFichier As String
    NomFichier = (VBA.Strings.Left(swModel.GetTitle, InStr(swModel.GetTitle, "-") - 2))
    
    ' Demande indice de révision puis ajout date du jour
    Dim Indice_brute As String
    Dim Indice As String
    'Indice = InputBox("Veuillez saisie l'indice du plan ?", "Indice du plan") 'La variable reçoit la valeur entrée dans l'InputBox
    Set cusPropMgr = swModel.Extension.CustomPropertyManager("")
    cusPropMgr.Get5 "Révision", False, Indice_brute, Indice, False
    If Indice = "" Then
       Indice = "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
    Else
       Indice = "-Ind" & Indice & "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
    End If

    'Recherche si le répertoire de destination est créé
    Dim CheminDestination As String
    CheminDestination = "u:\documents\plans"

    Dim NomRepertoire As String
    NomRepertoire = RechercheSsRepCommençantPar(CheminDestination, NomFichier)
    
    'Si le répertoire n'existe pas, création du répertoire avec la désignation de la pièce
    If NomRepertoire = "" Then
        Dim Designation_brute As String
        Dim Designation As String
        Set cusPropMgr = swRefModel.Extension.CustomPropertyManager("")
        cusPropMgr.Get5 "Designation", True, Designation_brute, Designation, False
        'cusPropMgr.Get5 "FieldName", IsConfigSpecific (True pour onglet Spécifiques à la configuration, False pour onglet Personnaliser), Value, ResolvedValue, WasResolved
        MkDir (CheminDestination & "\" & NomFichier & " - " & Designation)
        NomRepertoire = NomFichier & " - " & Designation
    End If
    
    ' Dossier temporaire pour les fichiers avant zip
    Dim CheminTemp As String
    CheminTemp = CheminDestination & "\" & NomRepertoire & "\_temp_export"
    MkDir CheminTemp

    'Sauvegarde sous DXF
    longstatus = swModel.SaveAs3(CheminTemp & "\" & NomFichier & Indice & ".dxf", 0, 0)

    ' Désactiver l'ouverture automatique du PDF
    Dim swExportPDFData As SldWorks.ExportPdfData
    Set swExportPDFData = swApp.GetExportFileData(1)
    If Not swExportPDFData Is Nothing Then
        swExportPDFData.ViewPdfAfterSaving = False
    End If

    'Sauvegarde sous PDF
    'longstatus = swModel.SaveAs3(CheminTemp & "\" & NomFichier & Indice & ".pdf", 0, 0)
    Dim errors As Long
    Dim warnings As Long
    bRet = swModel.Extension.SaveAs(CheminTemp & "\" & NomFichier & Indice & ".pdf", 0, 0, swExportPDFData, errors, warnings)

    'Sauvegarde sous STEP
    longstatus = swRefModel.SaveAs3(CheminTemp & "\" & NomFichier & Indice & ".STEP", 0, 0)

    ' Chemin du fichier ZIP final
    Dim CheminZip As String
    CheminZip = CheminDestination & "\" & NomRepertoire & "\" & NomRepertoire & ".zip"

    ' Création du ZIP
    ZipFiles CheminTemp, CheminZip

    ' Suppression du dossier temporaire
    Dim FSO2 As Object
    Set FSO2 = CreateObject("Scripting.FileSystemObject")
    FSO2.DeleteFolder CheminTemp, True
    
    'Ouverture du dossier contenant
    Shell "EXPLORER /n,/e," & CheminDestination & "\" & NomRepertoire
    
End Sub

Function RechercheSsRepCommençantPar(Chemin As String, Nom As String) As String
    Dim FSO, ListR, sRep, Rep, LesReps, NomRepertoire
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Chemin = "" Then Chemin = "c:\Program files"
    Set ListR = FSO.GetFolder(Chemin)
    Set sRep = ListR.SubFolders
    For Each Rep In sRep
        If VBA.Strings.Left(Rep.Name, Len(Nom)) = Nom Then
            NomRepertoire = Rep.Name
            Exit For
        End If
    Next
    RechercheSsRepCommençantPar = NomRepertoire
End Function

Sub ZipFiles(SourceFolder As String, ZipPath As String)
    ' Crée un ZIP vide
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Initialise le fichier ZIP (en-tête nécessaire)
    Dim iFile As Integer
    iFile = FreeFile
    Open ZipPath For Output As #iFile
    Print #iFile, Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, Chr(0))
    Close #iFile
    
    ' Copie les fichiers dans le ZIP via Shell
    Dim oShell As Object
    Set oShell = CreateObject("Shell.Application")
    
    Dim oZip As Object
    Set oZip = oShell.NameSpace(CVar(ZipPath))
    
    Dim oSource As Object
    Set oSource = oShell.NameSpace(CVar(SourceFolder))
    
    oZip.CopyHere oSource.Items
    
    ' Attendre que le zip soit terminé
    Dim nbFichiers As Integer
    nbFichiers = oSource.Items.Count
    Do While oZip.Items.Count < nbFichiers
        Dim dStart As Double
        dStart = Timer
        Do While Timer < dStart + 0.2  ' Pause de 200ms
            DoEvents
        Loop
    Loop

    Set oSource = Nothing
    Set oZip = Nothing
    Set oShell = Nothing

    Dim t As Single
    t = Timer: Do While Timer < t + 1: DoEvents: Loop
    
End Sub