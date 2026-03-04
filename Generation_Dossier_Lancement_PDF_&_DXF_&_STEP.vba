' ****************************************************************************************************
' Macro pour la Génération du Dossier de Lancement avec Export dans un fichier ZIP
' de la mise en plan au format PDF, DXF et de la pièce au format STEP
' ****************************************************************************************************
' Auteur : Tristan JACQ
' Date : Mars 2026
' Version : 1.7
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

        ' Affichage de la fenêtre de chargement
        frm_Loading.Show vbModeless  ' vbModeless = non bloquant, la macro continue
        DoEvents

        ' Extraction nom du fichier du nom de la feuille
        Dim NomFichier As String
        NomFichier = (VBA.Strings.Left(swModel.GetTitle, InStr(swModel.GetTitle, "-") - 1))
        NomFichier = VBA.Strings.Trim(NomFichier)

        ' Demande indice de révision puis ajout date du jour
        Dim Indice_brute As String
        Dim Indice As String
        ' Indice = InputBox("Veuillez saisie l'indice du plan ?", "Indice du plan") 'La variable reçoit la valeur entrée dans l'InputBox
        Set cusPropMgr = swModel.Extension.CustomPropertyManager("")
        cusPropMgr.Get5 "Révision", False, Indice_brute, Indice, False
        ' Trim enlève les espaces. Si l'indice contient " " il devient "" ; et on n'ajoute pas "Ind" pour éviter les noms du type "Ind -20260303"
        Indice = VBA.Strings.Trim(Indice)
        If Indice = "" Then
            Indice = "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
        Else
            Indice = "-Ind" & Indice & "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
        End If

        ' Recherche si le répertoire de destination est créé
        Dim CheminDestination As String
        CheminDestination = "u:\documents\plans"

        ' Chemin du dossier contenant les plans (pour la recherche des SLDDRW des composants de la nomenclature)
        Dim CheminPlan As String
        CheminPlan = "I:\Thomas plans"

        Dim NomRepertoire As String
        NomRepertoire = RechercheSsRepCommençantPar(CheminDestination, NomFichier)

        ' Si le répertoire n'existe pas, création du répertoire avec la désignation de la pièce
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

        ' SÉCURITÉ : Nettoyage si le dossier temporaire existe déjà (crash précédent)
        Dim FSO_Init As Object
        Set FSO_Init = CreateObject("Scripting.FileSystemObject")
        If FSO_Init.FolderExists(CheminTemp) Then
            On Error Resume Next ' Au cas où un fichier est verrouillé dans le dossier
            FSO_Init.DeleteFolder CheminTemp, True
            On Error Goto 0
                Wait 500
            End If

            ' Création du dossier temporaire
            MkDir CheminTemp

            ' Sauvegarde sous DXF
            longstatus = swModel.SaveAs3(CheminTemp & "\" & NomFichier & Indice & ".dxf", 0, 0)

            ' Désactiver l'ouverture automatique du PDF
            Dim swExportPDFData As SldWorks.ExportPdfData
            Set swExportPDFData = swApp.GetExportFileData(1)
            If Not swExportPDFData Is Nothing Then
                swExportPDFData.ViewPdfAfterSaving = False
            End If

            ' Sauvegarde sous PDF
            ' longstatus = swModel.SaveAs3(CheminTemp & "\" & NomFichier & Indice & ".pdf", 0, 0)
            Dim errors As Long
            Dim warnings As Long
            bRet = swModel.Extension.SaveAs(CheminTemp & "\" & NomFichier & Indice & ".pdf", 0, 0, swExportPDFData, errors, warnings)

            ' Modification des préférences d'export STEP pour exporter la géométrie complète de l'assemblage
            ResoudreAssemblage swRefModel
            ' AP203 pour ne pas conserver les propriétés personnalisées dans le STEP (couleur, matière, etc.) et ainsi éviter les erreurs d'export sur certains assemblages
            ' AP214 pour conserver les propriétés personnalisées dans le STEP (couleur, matière, etc.)
            'swApp.SetUserPreferenceIntegerValue swStepAP, 214
            swApp.SetUserPreferenceIntegerValue swStepExportPreference, 0
            ' 0 = Export As tessellated geometry (facettes)
            ' 1 = Export As solid/surface geometry (géométrie complète)

            ' Sauvegarde sous STEP
            longstatus = swRefModel.SaveAs3(CheminTemp & "\" & NomFichier & Indice & ".STEP", 0, 0)

            ' Si c'est un assemblage, export des mise en plan des composants de la nomenclature
            If swRefModel.GetType = swDocASSEMBLY Then
                ' MsgBox "Cette pièce est un assemblage, la macro va maintenant analyser la nomenclature pour tenter de localiser les composants." & vbCrLf & vbCrLf

                ' Lecture de la nomenclature via la BOM (Bill Of Materials) de la mise en plan
                LectureBOM swDraw, swRefModel, CheminPlan
            End If

            ' Chemin du fichier ZIP final
            Dim CheminZip As String
            CheminZip = CheminDestination & "\" & NomRepertoire & "\" & NomFichier & Indice & ".zip"

            ' Dossier Archives
            Dim CheminArchives As String
            CheminArchives = CheminDestination & "\" & NomRepertoire & "\Archives"

            ' Archivage des anciens ZIP
            'ArchiverAnciensZip CheminDestination & "\" & NomRepertoire, CheminZip, CheminArchives, NomFichier

            ' Création du ZIP
            'ZipFiles CheminTemp, CheminZip

            ' Suppression du dossier temporaire
            Dim FSO2 As Object
            Set FSO2 = CreateObject("Scripting.FileSystemObject")
            'FSO2.DeleteFolder CheminTemp, True

            ' Ouverture du dossier contenant le ZIP
            'Shell "EXPLORER /n,/e," & CheminDestination & "\" & NomRepertoire

            ' Ouverture du fichier PDF dans le ZIP
            'Shell "explorer.exe """ & CheminZip & "\" & NomFichier & Indice & ".pdf" & """", vbNormalFocus

            ' Fenêtre de fin avec récapitulatif et lien vers le dossier
            Unload frm_Loading
            MsgBox "Dossier de lancement généré avec succès !" & vbCrLf & vbCrLf & _
            NomRepertoire & vbCrLf & _
            NomFichier & Indice & ".zip", _
            vbInformation, "Export terminé"

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
            Wait 200
        Loop

        Set oSource = Nothing
        Set oZip = Nothing
        Set oShell = Nothing

        ' Attendre pour s'assurer que le processus est terminé
        Wait 500

End Sub

Function EstMemePiece(NomZip As String, NomFichier As String) As Boolean
    ' Vérifie que le ZIP appartient exactement à NomFichier sans variante
    ' Format attendu : NomFichier-YYYYMMDD.zip  ou  NomFichier-IndX-YYYYMMDD.zip

    Dim Prefixe As String
    Prefixe = NomFichier & "-"

    ' Le nom doit commencer par NomFichier-
    If VBA.Strings.Left(NomZip, Len(Prefixe)) <> Prefixe Then
        EstMemePiece = False
     Exit Function
    End If

    ' On récupère ce qui suit NomFichier-
    Dim Suite As String
    Suite = Mid(NomZip, Len(Prefixe) + 1)  ' ex: "20260303.zip" ou "IndA-20260303.zip"

    ' Cas 1 : NomFichier-YYYYMMDD.zip  → commence par 8 chiffres
    If EstDate8(VBA.Strings.Left(Suite, 8)) Then
        EstMemePiece = True
     Exit Function
    End If

    ' Cas 2 : NomFichier-IndX-YYYYMMDD.zip  → commence par "Ind"
    If VBA.Strings.Left(Suite, 3) = "Ind" Then
        EstMemePiece = True
     Exit Function
    End If

    ' Sinon : variante (ex: -10-, -GH-) → on ne touche pas
    EstMemePiece = False
End Function

Function EstDate8(s As String) As Boolean
    ' Retourne True si s est une chaîne de 8 chiffres (YYYYMMDD)
    If Len(s) <> 8 Then EstDate8 = False : Exit Function
        Dim i As Integer
        For i = 1 To 8
            If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then
                EstDate8 = False
             Exit Function
            End If
        Next i
        EstDate8 = True
End Function

' Fonction pour archiver les anciens ZIP : si même indice, suppression ; sinon déplacement dans Archives
Sub ArchiverAnciensZip(DossierPiece As String, NouveauZip As String, DossierArchives As String, NomFichier As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    For Each oFile In FSO.GetFolder(DossierPiece).Files
        If LCase(FSO.GetExtensionName(oFile.Name)) = "zip" Then
            ' On ne touche que les ZIP qui appartiennent à cette pièce
            If EstMemePiece(oFile.Name, NomFichier) Then
                If LCase(oFile.Path) = LCase(NouveauZip) Then
                    ' Même nom = même indice + même date → suppression (sera recréé)
                    FSO.DeleteFile oFile.Path, True
                Else
                    ' Indice précédent → déplacement dans Archives
                    If Not FSO.FolderExists(DossierArchives) Then
                        MkDir DossierArchives
                    End If
                    Dim Destination As String
                    Destination = DossierArchives & "\" & oFile.Name
                    If FSO.FileExists(Destination) Then
                        FSO.DeleteFile Destination, True
                    End If
                    FSO.MoveFile oFile.Path, Destination
                End If
            End If
            ' Sinon : ZIP sans rapport avec la pièce → on ne touche pas
        End If
    Next oFile
End Sub

' Fonction Wait en millisecondes
Sub Wait(Millis As Double)
    Dim Fin As Double
    Fin = Timer + Millis / 1000
    Do While Timer < Fin
        DoEvents
    Loop
End Sub

' Fonction pour pouvoir rendre tous les composants d'un assemblage "légers" en "résolus" et ainsi éviter les erreurs d'export STEP
Sub ResoudreAssemblage(Byval swModel As Object)
    If Not swModel Is Nothing Then
        If swModel.GetType = 2 Then ' 2 correspond à swDocASSEMBLY
            Dim swAssy As Object
            Set swAssy = swModel
            ' Force la résolution de tous les composants
            swAssy.ResolveAllLightweightComponents True
        End If
    End If
End Sub

' Fonction pour lire la nomenclature de la mise en plan via la BOM (Bill Of Materials)
Sub LectureBOM(swDraw As SldWorks.DrawingDoc, ByVal swRefModel As Object, CheminPlan As String)
    Dim swFeat      As SldWorks.Feature
    Dim swBomFeat   As SldWorks.BomFeature
    Dim swBomTable  As SldWorks.BomTableAnnotation
    Dim i As Long

    ' Recherche de la BOM dans la mise en plan
    Set swFeat = swDraw.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "BomFeat" Then
            Set swBomFeat = swFeat.GetSpecificFeature2
            Set swBomTable = swBomFeat.GetTableAnnotations(0)
            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

    If swBomTable Is Nothing Then
        MsgBox "Aucune BOM trouvée dans la mise en plan."
        Exit Sub
    End If

    ' Indexation unique de tous les SLDDRW
    Dim IndexDRW As Collection
    Set IndexDRW = IndexerSLDDRW(CheminPlan)

    ' Export CSV
    Dim CheminCSV As String
    CheminCSV = Environ("TEMP") & "\diagnostic_bom.csv"
    Dim iCSV As Integer
    iCSV = FreeFile
    Open CheminCSV For Output As #iCSV
    Print #iCSV, "Num;Numero Plan;Designation;Chemin SLDDRW"

    Dim Introuvables As String
    Introuvables = ""

    For i = 1 To swBomTable.RowCount - 1
        Dim Ligne_NumPart   As String
        Dim Ligne_NumPlan   As String
        Dim Ligne_Design    As String
        Dim CheminDRW       As String
        Ligne_NumPart   = swBomTable.Text(i, 0)
        Ligne_NumPlan   = swBomTable.Text(i, 1)
        Ligne_Design    = swBomTable.Text(i, 2)
        CheminDRW       = TrouverDansIndex(IndexDRW, Ligne_NumPlan)

        ' CSV
        Dim csv_Chemin As String
        If CheminDRW = "" Then
            csv_Chemin = "[INTROUVABLE]"
            Introuvables = Introuvables & "  " & Ligne_NumPart & " | " & Ligne_NumPlan & " | " & Ligne_Design & vbCrLf
        Else
            csv_Chemin = CheminDRW
        End If
        Print #iCSV, Ligne_NumPart & ";" & Ligne_NumPlan & ";" & Ligne_Design & ";" & csv_Chemin
    Next i

    Close #iCSV

    ' Ouvrir le CSV automatiquement
    Shell "explorer.exe """ & CheminCSV & """", vbNormalFocus
    Wait 1000

    ' MsgBox : uniquement les introuvables
    If Introuvables = "" Then
        MsgBox "Tous les fichiers SLDDRW ont été trouvés.", vbInformation, "Résultat BOM"
    Else
        MsgBox "Fichiers SLDDRW introuvables :" & vbCrLf & "----------------------------------------" & vbCrLf & Introuvables, _
               vbExclamation, "Résultat BOM"
    End If

End Sub

' À appeler UNE SEULE FOIS au début pour indexer tous les SLDDRW
Function IndexerSLDDRW(CheminPlan As String) As Collection
    Dim FSO As Object
    Dim oShell As Object
    Dim ResultatFichier As String
    Dim col As New Collection

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set oShell = CreateObject("WScript.Shell")

    ResultatFichier = Environ("TEMP") & "\index_drw.txt"

    If FSO.FileExists(ResultatFichier) Then
        On Error Resume Next
        FSO.DeleteFile ResultatFichier
        On Error Goto 0
        End If

        ' Un seul DIR pour tous les SLDDRW d'un coup
        oShell.Run "cmd /c dir """ & CheminPlan & "\*.SLDDRW"" /s /b > """ & ResultatFichier & """ 2>nul", _
        0, True

        ' Charger tous les chemins en mémoire dans une Collection
        If FSO.FileExists(ResultatFichier) Then
            Dim iFile As Integer
            iFile = FreeFile
                Dim Ligne As String
                Open ResultatFichier For Input As #iFile
                Do While Not EOF(iFile)
                    Line Input #iFile, Ligne
                    Ligne = Trim(Ligne)
                    If Ligne <> "" Then
                        If InStr(1, Ligne, "archive", vbTextCompare) = 0 Then
                            Dim Cle As String
                            Cle = LCase(FSO.GetFileName(Ligne))
                            On Error Resume Next
                            col.Add Ligne, Cle
                            If Err.Number <> 0 Then
                                Err.Clear
                            End If
                            On Error Goto 0
                            End If
                        End If
                    Loop
                    Close #iFile
                    On Error Resume Next
                    FSO.DeleteFile ResultatFichier
                    On Error Goto 0
                    End If

                    Set IndexerSLDDRW = col
End Function

' Cherche dans l'index déjà chargé en mémoire — instantané
Function TrouverDansIndex(Index As Collection, NumeroPlan As String) As String
    Dim NomCible As String
    NomCible = LCase(NumeroPlan & ".slddrw")
    On Error Resume Next
    TrouverDansIndex = Index(NomCible)
    On Error Goto 0
End Function