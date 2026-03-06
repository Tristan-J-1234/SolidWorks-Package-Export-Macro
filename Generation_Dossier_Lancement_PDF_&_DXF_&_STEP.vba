' ****************************************************************************************************
' Macro pour la Génération du Dossier de Lancement avec Export dans un fichier ZIP
' de la mise en plan au format PDF, DXF et de la pièce au format STEP
' ****************************************************************************************************
' Auteur : Tristan JACQ
' Date : Mars 2026
' Version : 1.14
' ****************************************************************************************************
' Modifications de la version :
'   - Mutualisation des exports PDF/DXF/STEP dans la fonction ExporterMiseEnPlan
'   - Export STEP via le modèle référencé par la première vue de la mise en plan
'   - Lecture de l'indice de révision depuis les propriétés de chaque mise en plan composant
'   - Correction du faux positif dans RechercheSsRepCommençantPar (ex: 98765432 → 987654321)
' ****************************************************************************************************

Sub main()

    Dim swApp                   As SldWorks.SldWorks
    Dim swModel                 As SldWorks.ModelDoc2
    Dim swDraw                  As SldWorks.DrawingDoc
    Dim swSheet                 As SldWorks.Sheet
    Dim swView                  As SldWorks.View
    Dim swActiveView            As SldWorks.View
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
        CheminDestination = "U:\DOCUMENTS\PLANS"

        ' Chemin du dossier contenant les plans (pour la recherche des SLDDRW des composants de la nomenclature)
        Dim CheminPlan As String
        CheminPlan = "I:\Thomas plans"
        ' CHeminPlan = "T:\Commun\Transfert\Tristan JACQ\6 - Macro SolidWorks\Fichiers SolidWorks"

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

            ' Export de la mise en plan au format PDF et DXF, ainsi que de la pièce référencée au format STEP
            ExporterMiseEnPlan swApp, swDraw, CheminTemp

            ' Si c'est un assemblage, export des mise en plan des composants de la nomenclature
            If swRefModel.GetType = swDocASSEMBLY Or ContientBOM(swDraw) Then
                ' MsgBox "Cette pièce est un assemblage, la macro va maintenant analyser la nomenclature pour tenter de localiser les composants." & vbCrLf & vbCrLf

                ' Lecture de la nomenclature via la BOM (Bill Of Materials) de la mise en plan
                Dim CheminCSV_BOM   As String
                Dim Introuvables_BOM As String
                Dim BOM_Trouvee As Boolean
                BOM_Trouvee = LectureBOM(swApp, swDraw, swRefModel, CheminPlan, CheminTemp, CheminCSV_BOM, Introuvables_BOM)
            End If

            ' Chemin du fichier ZIP final
            Dim CheminZip As String
            CheminZip = CheminDestination & "\" & NomRepertoire & "\" & NomFichier & Indice & ".zip"

            ' Dossier Archives
            Dim CheminArchives As String
            CheminArchives = CheminDestination & "\" & NomRepertoire & "\Archives"

            ' Archivage des anciens ZIP
            ArchiverAnciensZip CheminDestination & "\" & NomRepertoire, CheminZip, CheminArchives, NomFichier

            ' Création du ZIP
            ZipFiles CheminTemp, CheminZip

            ' Suppression du dossier temporaire
            Dim FSO2 As Object
            Set FSO2 = CreateObject("Scripting.FileSystemObject")
            FSO2.DeleteFolder CheminTemp, True

            ' Ouverture du dossier contenant le ZIP
            Shell "EXPLORER /n,/e," & CheminDestination & "\" & NomRepertoire

            ' Ouverture du fichier PDF dans le ZIP
            'Shell "explorer.exe """ & CheminZip & "\" & NomFichier & Indice & ".pdf" & """", vbNormalFocus

            ' Ouverture du CSV et affichage des résultats BOM en fin de macro
            If BOM_Trouvee Then
                Shell "explorer.exe """ & CheminCSV_BOM & """", vbNormalFocus
                Wait 1000
            End If

            Unload frm_Loading

            ' MsgBox finale groupée
            Dim MsgFinale As String
            MsgFinale = "Dossier de lancement généré avec succès !" & vbCrLf & "----------------------------------------" & vbCrLf & _
                        NomRepertoire & vbCrLf & _
                        NomFichier & Indice & ".zip"

            If BOM_Trouvee Then
                If Introuvables_BOM = "" Then
                    MsgFinale = MsgFinale & vbCrLf & "----------------------------------------" & vbCrLf & "[OK] Tous les fichiers SLDDRW ont été trouvés."
                Else
                    MsgFinale = MsgFinale & vbCrLf & "----------------------------------------" & vbCrLf & "[!] Fichiers SLDDRW introuvables :" & vbCrLf & Introuvables_BOM
                End If
            ElseIf swRefModel.GetType = swDocASSEMBLY Or ContientBOM(swDraw) Then
                MsgFinale = MsgFinale & vbCrLf & "----------------------------------------" & vbCrLf & "[!] Aucune nomenclature trouvée dans la mise en plan."
            End If

            MsgBox MsgFinale, vbInformation, "Export terminé"

End Sub

Function RechercheSsRepCommençantPar(Chemin As String, Nom As String) As String
    Dim FSO, ListR, sRep, Rep, NomRepertoire
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Chemin = "" Then Chemin = "C:\Program Files"
    Set ListR = FSO.GetFolder(Chemin)
    Set sRep = ListR.SubFolders
    For Each Rep In sRep
        If VBA.Strings.Left(Rep.Name, Len(Nom)) = Nom Then
            ' Vérifier que le caractère suivant est un espace ou tiret (pas un autre chiffre/lettre)
            Dim Suite As String
            Suite = Mid(Rep.Name, Len(Nom) + 1)
            If Suite = "" Or Left(Suite, 1) = " " Or Left(Suite, 1) = "-" Then
                NomRepertoire = Rep.Name
                Exit For
            End If
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
Function LectureBOM(swApp As SldWorks.SldWorks, swDraw As SldWorks.DrawingDoc, ByVal swRefModel As Object, CheminPlan As String, CheminTemp As String, ByRef CheminCSV_Out As String, ByRef Introuvables_Out As String) As Boolean

    ' Indexation unique de tous les SLDDRW
    Dim IndexDRW As Collection
    Set IndexDRW = IndexerSLDDRW(CheminPlan)

    ' Préparation CSV
    Dim CheminCSV As String
    CheminCSV = Environ("TEMP") & "\diagnostic_bom_" & VBA.Strings.Format(Now, "YYYY_MM_DD_HH_MM_SS") & ".csv"
    Dim iCSV As Integer
    iCSV = FreeFile
    Open CheminCSV For Output As #iCSV
    Print #iCSV, "Num;Numero Plan;Designation;Chemin SLDDRW"

    Dim Introuvables As String
    Introuvables = ""

    ' Collection pour éviter les doublons/boucles infinies
    Dim DejaTraites As New Collection

    ' Obtenir la table racine
    Dim oTableRacine As Object
    Set oTableRacine = ObtenirTable(swDraw)

    If oTableRacine Is Nothing Then
        Close #iCSV
        LectureBOM = False
        Exit Function
    End If

    ' Traitement récursif
    TraiterLignesTable swApp, oTableRacine, IndexDRW, iCSV, DejaTraites, Introuvables, CheminTemp

    Close #iCSV
    CheminCSV_Out = CheminCSV
    Introuvables_Out = Introuvables
    LectureBOM = True
End Function

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

' Fonction pour vérifier si la mise en plan contient une BOM (nomenclature) et ainsi décider d'exporter ou non les plans des composants
' Utilisée pour éviter d'exporter les pièces possedant une nomenclature telles que les pièces soudées qui sont des pièces mais avec une BOM dans la mise en plan
Function ContientBOM(swDraw As SldWorks.DrawingDoc) As Boolean
    Dim swFeat As SldWorks.Feature
    Set swFeat = swDraw.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "BomFeat" Or swFeat.GetTypeName2 = "WeldmentTableFeat" Then
            ContientBOM = True
         Exit Function
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    ContientBOM = False
End Function

' Fonction récursive principale — appelée par LectureBOM et par elle-même
Sub TraiterLignesTable(swApp As SldWorks.SldWorks, _
                       oTable As Object, _
                       IndexDRW As Collection, _
                       iCSV As Integer, _
                       DejaTraites As Collection, _
                       ByRef Introuvables As String, _
                       CheminTemp As String)

    ' Vérification défensive avant d'utiliser la table
    If oTable Is Nothing Then Exit Sub
    
    Dim NbLignes As Long
    On Error Resume Next
    NbLignes = oTable.RowCount
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub  ' L'objet ne supporte pas RowCount → ce n'est pas une table valide
    End If
    On Error GoTo 0

    Dim r As Long
    For r = 1 To NbLignes - 1
        Dim NumPart  As String
        Dim NumPlan  As String
        Dim Design   As String
        NumPart = Trim(oTable.Text(r, 0))
        NumPlan = Trim(oTable.Text(r, 1))
        Design  = Trim(oTable.Text(r, 2))

        ' Nettoyer les retours à la ligne dans les cellules (désignations sur plusieurs lignes)
        NumPlan = Join(Split(NumPlan, vbCrLf), " ")
        NumPlan = Join(Split(NumPlan, vbLf), " ")
        Design  = Join(Split(Design, vbCrLf), " ")
        Design  = Join(Split(Design, vbLf), " ")

        ' Ignorer les lignes vides ou invalides (cartouche, en-têtes parasites)
        If NumPlan = "" Then GoTo Suite
        If Not EstNumeroPlanValide(NumPlan) Then GoTo Suite

        ' Ignorer si déjà traité (évite les boucles infinies)
        Dim DejaVu As Boolean
        DejaVu = False
        On Error Resume Next
        Dim test As String
        test = DejaTraites(LCase(NumPlan))
        If Err.Number = 0 Then DejaVu = True
        Err.Clear
        On Error GoTo 0
        If DejaVu Then GoTo Suite

        ' Marquer comme traité
        DejaTraites.Add NumPlan, LCase(NumPlan)

        ' Chercher le SLDDRW
        Dim CheminDRW As String
        CheminDRW = TrouverDansIndex(IndexDRW, NumPlan)

        Dim csv_Chemin As String
        If CheminDRW = "" Then
            csv_Chemin = "[INTROUVABLE]"
            Introuvables = Introuvables & "  " & NumPart & " | " & NumPlan & " | " & Design & vbCrLf
        Else
            csv_Chemin = CheminDRW
        End If
        
        Print #iCSV, NumPart & ";" & Chr(34) & NumPlan & Chr(34) & ";" & Design & ";" & csv_Chemin

        ' Si le SLDDRW existe, l'ouvrir et vérifier s'il contient une BOM
        If CheminDRW <> "" Then

            Dim lErr As Long, lWarn As Long
            Dim swDrawSub As SldWorks.DrawingDoc
            Set swDrawSub = swApp.OpenDoc6(CheminDRW, swDocDRAWING, swOpenDocOptions_Silent, "", lErr, lWarn)

            If Not swDrawSub Is Nothing Then

                ' Export de la mise en plan au format PDF et DXF, ainsi que de la pièce référencée au format STEP
                ExporterMiseEnPlan swApp, swDrawSub, CheminTemp

                ' Chercher une BOM standard
                Dim swFeatSub   As SldWorks.Feature
                Dim swBomSub    As SldWorks.BomFeature
                Dim swTableSub  As SldWorks.BomTableAnnotation
                Set swFeatSub = swDrawSub.FirstFeature
                Do While Not swFeatSub Is Nothing
                    If swFeatSub.GetTypeName2 = "BomFeat" Then
                        Set swBomSub = swFeatSub.GetSpecificFeature2
                        Set swTableSub = swBomSub.GetTableAnnotations(0)
                        Exit Do
                    End If
                    Set swFeatSub = swFeatSub.GetNextFeature
                Loop

                If Not swTableSub Is Nothing Then
                    ' Récursion sur la BOM standard
                    Dim oTableStd As Object
                    Set oTableStd = swTableSub
                    TraiterLignesTable swApp, oTableStd, IndexDRW, iCSV, DejaTraites, Introuvables, CheminTemp
                Else
                    ' Chercher une table weldment via les vues
                    Dim swViewSub As SldWorks.View
                    Dim vAnnotsSub As Variant
                    Dim oTableWeld As Object
                    Set swViewSub = swDrawSub.GetFirstView
                    Do While Not swViewSub Is Nothing
                        vAnnotsSub = swViewSub.GetTableAnnotations
                        If Not IsEmpty(vAnnotsSub) And Not IsNull(vAnnotsSub) Then
                            Dim j As Long
                            For j = 0 To UBound(vAnnotsSub)
                                Dim oTSub As Object
                                Set oTSub = vAnnotsSub(j)
                                Dim tTypeSub As Long
                                On Error Resume Next
                                tTypeSub = oTSub.TableType
                                On Error GoTo 0
                                If tTypeSub = 3 Then  ' Weldment uniquement
                                    Set oTableWeld = oTSub
                                    TraiterLignesTable swApp, oTableWeld, IndexDRW, iCSV, DejaTraites, Introuvables, CheminTemp
                                    Exit For
                                End If
                            Next j
                        End If
                        Set swViewSub = swViewSub.GetNextView
                    Loop
                End If

                ' Fermer le SLDDRW silencieusement
                swApp.CloseDoc swDrawSub.GetPathName
            End If
        End If

Suite:
    Next r
End Sub

' Extrait la TableAnnotation d'un DrawingDoc (BOM standard ou weldment)
Function ObtenirTable(swDraw As SldWorks.DrawingDoc) As Object
    ' Chercher BOM standard en priorité
    Dim swFeat As SldWorks.Feature
    Set swFeat = swDraw.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "BomFeat" Then
            Dim swBom As SldWorks.BomFeature
            Set swBom = swFeat.GetSpecificFeature2
            Set ObtenirTable = swBom.GetTableAnnotations(0)
            Exit Function
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

    ' Sinon chercher weldment uniquement (type 10)
    ' swTableAnnotation_e : 0=General, 1=BOM, 2=Revision, 3=Weldment, 4=HoleChart, etc.
    Dim swView As SldWorks.View
    Dim vAnnots As Variant
    Set swView = swDraw.GetFirstView
    Do While Not swView Is Nothing
        vAnnots = swView.GetTableAnnotations
        If Not IsEmpty(vAnnots) And Not IsNull(vAnnots) Then
            Dim i As Long
            For i = 0 To UBound(vAnnots)
                Dim oT As Object
                Set oT = vAnnots(i)
                ' TableType : 0=General, 1=BOM, 2=Revision, 3=Weldment cutlist, 4=HoleChart
                Dim tType As Long
                On Error Resume Next
                tType = oT.TableType
                On Error GoTo 0
                ' On accepte uniquement Weldment (3) — pas Revision (2) ni HoleChart (4)
                If tType = 3 Then
                    Set ObtenirTable = oT
                    Exit Function
                End If
            Next i
        End If
        Set swView = swView.GetNextView
    Loop

    Set ObtenirTable = Nothing
End Function

Function EstNumeroPlanValide(NumPlan As String) As Boolean
    ' Rejette les chaînes trop courtes ou qui contiennent des mots clés de cartouche
    If Len(NumPlan) < 2 Then EstNumeroPlanValide = False : Exit Function
    If InStr(1, NumPlan, "REV", vbTextCompare) > 0 Then EstNumeroPlanValide = False : Exit Function
    If InStr(1, NumPlan, "DATE", vbTextCompare) > 0 Then EstNumeroPlanValide = False : Exit Function
    If InStr(1, NumPlan, "DESCRIPTION", vbTextCompare) > 0 Then EstNumeroPlanValide = False : Exit Function
    If InStr(1, NumPlan, "Mise a jour", vbTextCompare) > 0 Then EstNumeroPlanValide = False : Exit Function
    If Len(NumPlan) = 1 And NumPlan >= "A" And NumPlan <= "Z" Then EstNumeroPlanValide = False : Exit Function
    EstNumeroPlanValide = True
End Function

' Fonction d'export de la mise en plan au format PDF et DXF, ainsi que de la pièce référencée au format STEP
Sub ExporterMiseEnPlan(swApp As SldWorks.SldWorks, swDrawExp As SldWorks.DrawingDoc, CheminTemp As String)

    Dim cusPropMgr   As SldWorks.CustomPropertyManager
    Dim Indice_brute As String
    Dim Indice_resolved As String

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim NomExp As String
    NomExp = FSO.GetBaseName(swDrawExp.GetPathName)
    If NomExp = "" Then NomExp = FSO.GetBaseName(swDrawExp.GetTitle)

    ' Lire l'indice
    Set cusPropMgr = swDrawExp.Extension.CustomPropertyManager("")
    cusPropMgr.Get5 "Révision", False, Indice_brute, Indice_resolved, False
    Indice_resolved = VBA.Strings.Trim(Indice_resolved)

    Dim Suffixe As String
    If Indice_resolved = "" Then
        Suffixe = "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
    Else
        Suffixe = "-Ind" & Indice_resolved & "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
    End If

    Dim NomFinal As String
    NomFinal = NomExp & Suffixe

    ' Export PDF
    Dim swPDFData As SldWorks.ExportPdfData
    Set swPDFData = swApp.GetExportFileData(1)
    If Not swPDFData Is Nothing Then swPDFData.ViewPdfAfterSaving = False
    Dim lE As Long, lW As Long
    swDrawExp.Extension.SaveAs CheminTemp & "\" & NomFinal & ".pdf", 0, 0, swPDFData, lE, lW

    ' Export DXF
    swDrawExp.SaveAs3 CheminTemp & "\" & NomFinal & ".dxf", 0, 0

    ' Export STEP via le modèle référencé par la première vue
    Dim swView As SldWorks.View
    Set swView = swDrawExp.GetFirstView
    Set swView = swView.GetNextView  ' La première vue réelle (pas la vue feuille)
    If Not swView Is Nothing Then
        Dim swRefModel As Object
        Set swRefModel = swView.ReferencedDocument
        If Not swRefModel Is Nothing Then
            ResoudreAssemblage swRefModel
            swApp.SetUserPreferenceIntegerValue swStepExportPreference, 0
            swRefModel.SaveAs3 CheminTemp & "\" & NomFinal & ".STEP", 0, 0
        End If
    End If

End Sub