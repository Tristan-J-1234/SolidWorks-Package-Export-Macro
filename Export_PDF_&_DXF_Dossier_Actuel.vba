' ****************************************************************************************************
' Macro basée sur Generation_Dossier_Lancement_PDF_&_DXF_&_STEP.swp
' Cette macro est ultra simple et permet d'exporter en PDF et DXF la mise en plan active dans son 
' dossier d'origine au sein d'un dossier "pdf" et non dans U:\DOCUMENTS\PLANS.
' Le nom du fichier exporté est basé sur le nom de la mise en plan + indice de révision (s'il existe)
' + date.
' Les anciens fichiers PDF et DXF du même plan sont archivés dans un sous-dossier "Archives" ou
' supprimés s'ils ont le même indice et la même date (on considère que c'est la même version du plan).
' ****************************************************************************************************
' Auteur : Tristan JACQ
' Date : Mars 2026
' Version : 1.0
' ****************************************************************************************************
' Modifications de la version :
'   - Première version basée sur Generation_Dossier_Lancement_PDF_&_DXF_&_STEP.swp version 1.23
' ****************************************************************************************************

Sub main()

    Dim swApp   As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw  As SldWorks.DrawingDoc

    Set swApp   = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Vérification : document actif = mise en plan
    If swModel Is Nothing Then Exit Sub
    If swModel.GetType <> swDocDRAWING Then
        MsgBox "Cette macro ne fonctionne que sur une mise en plan.", vbExclamation
        Exit Sub
    End If

    Set swDraw = swModel

    ' --------------------------------------------------------
    ' 1. Nom de base du fichier (sans extension, sans " - Feuille X")
    ' --------------------------------------------------------
    Dim NomFichier As String
    Dim PosSepar   As Long
    PosSepar = InStr(swModel.GetTitle, " - ")
    If PosSepar > 0 Then
        NomFichier = Left(swModel.GetTitle, PosSepar - 1)
    Else
        NomFichier = swModel.GetTitle
    End If
    NomFichier = Trim(NomFichier)

    ' --------------------------------------------------------
    ' 2. Suffixe : indice de révision + date
    ' --------------------------------------------------------
    Dim cusPropMgr   As SldWorks.CustomPropertyManager
    Dim Indice_brute As String
    Dim Indice       As String
    Set cusPropMgr = swModel.Extension.CustomPropertyManager("")
    cusPropMgr.Get5 "Révision", False, Indice_brute, Indice, False
    Indice = Trim(Indice)

    Dim Suffixe As String
    If Indice = "" Then
        Suffixe = "-" & Format(Date, "YYYYMMDD")
    Else
        Suffixe = "-Ind" & Indice & "-" & Format(Date, "YYYYMMDD")
    End If

    Dim NomFinal As String
    NomFinal = NomFichier & Suffixe   ' ex: 12345-IndA-20260318

    ' --------------------------------------------------------
    ' 3. Dossiers : pdf\ et pdf\Archives\ à côté du SLDDRW
    ' --------------------------------------------------------
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' Dossier du fichier SLDDRW ouvert
    Dim CheminMiseEnPlan As String
    CheminMiseEnPlan = FSO.GetParentFolderName(swModel.GetPathName)

    If CheminMiseEnPlan = "" Then
        MsgBox "La mise en plan doit être enregistrée avant l'export.", vbExclamation
        Exit Sub
    End If

    Dim CheminPDF      As String
    Dim CheminArchives As String
    CheminPDF      = CheminMiseEnPlan & "\pdf"
    CheminArchives = CheminPDF & "\Archives"

    If Not FSO.FolderExists(CheminPDF)      Then MkDir CheminPDF
    If Not FSO.FolderExists(CheminArchives) Then MkDir CheminArchives

    ' --------------------------------------------------------
    ' 4. Archivage des anciens PDF / DXF du même plan
    ' --------------------------------------------------------
    ArchiverAnciensFichiers FSO, CheminPDF, CheminArchives, NomFichier, NomFinal

    ' --------------------------------------------------------
    ' 5. Export PDF
    ' --------------------------------------------------------
    Dim swPDFData As SldWorks.ExportPdfData
    Set swPDFData = swApp.GetExportFileData(1)  ' 1 = PDF
    If Not swPDFData Is Nothing Then swPDFData.ViewPdfAfterSaving = False

    Dim lErr As Long, lWarn As Long
    Dim CheminFichierPDF As String
    CheminFichierPDF = CheminPDF & "\" & NomFinal & ".pdf"
    swDraw.Extension.SaveAs CheminFichierPDF, 0, 0, swPDFData, lErr, lWarn

    ' --------------------------------------------------------
    ' 6. Export DXF
    ' --------------------------------------------------------
    Dim CheminFichierDXF As String
    CheminFichierDXF = CheminPDF & "\" & NomFinal & ".dxf"
    swDraw.SaveAs3 CheminFichierDXF, 0, 0

    ' --------------------------------------------------------
    ' 7. Ouverture du dossier et message final
    ' --------------------------------------------------------
    Shell "EXPLORER /n,/e," & CheminPDF

    MsgBox "Export terminé avec succès !" & vbCrLf & vbCrLf & _
           "Dossier : " & CheminPDF & vbCrLf & _
           "Fichier : " & NomFinal, vbInformation, "Export PDF + DXF"

End Sub

' ============================================================
' Archivage des anciens fichiers PDF et DXF du même plan
'   - Même indice + même date → suppression (sera recréé)
'   - Indice différent        → déplacement dans Archives
' ============================================================
Sub ArchiverAnciensFichiers(FSO As Object, CheminPDF As String, CheminArchives As String, NomFichier As String, NomFinal As String)

    Dim oFile   As Object
    Dim Prefixe As String
    Prefixe = NomFichier & "-"  ' on ne touche que les fichiers de CE plan

    For Each oFile In FSO.GetFolder(CheminPDF).Files

        Dim Ext As String
        Ext = LCase(FSO.GetExtensionName(oFile.Name))

        ' On ne traite que PDF et DXF
        If Ext = "pdf" Or Ext = "dxf" Then

            ' On ne touche que les fichiers qui commencent par NomFichier-
            If Left(oFile.Name, Len(Prefixe)) = Prefixe Then

                ' Nom sans extension pour comparaison
                Dim NomSansExt As String
                NomSansExt = FSO.GetBaseName(oFile.Name)

                If LCase(NomSansExt) = LCase(NomFinal) Then
                    ' Même nom = même indice + même date → on supprime (sera recréé par SolidWorks)
                    FSO.DeleteFile oFile.Path, True
                Else
                    ' Indice ou date différente → déplacement dans Archives
                    Dim Dest As String
                    Dest = CheminArchives & "\" & oFile.Name
                    If FSO.FileExists(Dest) Then FSO.DeleteFile Dest, True
                    FSO.MoveFile oFile.Path, Dest
                End If

            End If
        End If

    Next oFile

End Sub
