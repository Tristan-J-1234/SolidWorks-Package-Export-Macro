' ****************************************************************************************************
' Macro pour la Génération du Dossier de Lancement avec Export dans un fichier ZIP
' de la mise en plan au format PDF, DXF et de la pièce au format STEP
' ****************************************************************************************************
' Auteur : Tristan JACQ
' Date : Avril 2026
' Version : 1.24
' ****************************************************************************************************
' Modifications de la version :
'   - [à remplir]
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

    ' Vérification que le document actif est une mise en plan
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
        Dim PosSepar As Long
        PosSepar = InStr(swModel.GetTitle, " - ")
        If PosSepar > 0 Then
            NomFichier = VBA.Strings.Left(swModel.GetTitle, PosSepar - 1)
        Else
            NomFichier = swModel.GetTitle  ' Pas de " - " trouvé donc on prend tout le nom du fichier
        End If
        NomFichier = VBA.Strings.Trim(NomFichier)
        NomFichier = VBA.Strings.Trim(NomFichier)

        ' Lecture de l'indice de révision pour le nommage du ZIP
        Dim Indice_brute As String
        Dim Indice As String
        Set cusPropMgr = swModel.Extension.CustomPropertyManager("")
        cusPropMgr.Get5 "Révision", False, Indice_brute, Indice, False
        ' Trim supprime les espaces
        Indice = VBA.Strings.Trim(Indice)
        ' Si l'indice est vide, on met juste la date ; sinon on ajoute "Ind" devant puis la date
        If Indice = "" Then
            Indice = "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
        Else
            Indice = "-Ind" & Indice & "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
        End If

        ' Recherche si le répertoire de destination est créé
        Dim CheminDestination As String
        ' CheminDestination = "U:\DOCUMENTS\PLANS"
        CheminDestination = "T:\Commun\Transfert\Tristan JACQ\6 - Macro SolidWorks\Fichiers SolidWorks\test export"

        ' Chemin du dossier contenant les plans (pour la recherche des SLDDRW des composants de la nomenclature)
        Dim CheminPlan As String
        ' CheminPlan = "I:\Thomas plans"
        CheminPlan = "T:\Commun\Transfert\Tristan JACQ\6 - Macro SolidWorks\Fichiers SolidWorks"

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

        ' Dossier temporaire pour les fichiers avant la compression en ZIP
        Dim CheminTemp As String
        CheminTemp = CheminDestination & "\" & NomRepertoire & "\_temp_export"

        ' Nettoyage du dossier temporaire s'il existe déjà (sécurité au cas où une exécution précédente aurait été interrompue avant de le supprimer)
        Dim FSO_Init As Object
        Set FSO_Init = CreateObject("Scripting.FileSystemObject")
        If FSO_Init.FolderExists(CheminTemp) Then
            On Error Resume Next ' Au cas où un fichier serait ouvert ou protégé
            FSO_Init.DeleteFolder CheminTemp, True
            On Error Goto 0
                Wait 500
            End If

            ' Création du dossier temporaire
            MkDir CheminTemp

            ' Export de la mise en plan au format PDF et DXF, ainsi que de la pièce référencée au format STEP
            ExporterMiseEnPlan swApp, swDraw, CheminTemp

            ' Si c'est un assemblage ou une pièce avec BOM, tenter de localiser les plans des composants via la nomenclature pour les inclure dans le ZIP
            If swRefModel.GetType = swDocASSEMBLY Or ContientBOM(swDraw) Then
                ' Lecture de la nomenclature via la BOM (Bill Of Materials) de la mise en plan
                Dim CheminCSV_BOM   As String
                Dim Introuvables_BOM As String
                Dim BOM_Trouvee As Boolean
                BOM_Trouvee = LectureBOM(swApp, swDraw, swRefModel, CheminPlan, CheminTemp, CheminDestination, CheminCSV_BOM, Introuvables_BOM)
            End If

            ' Chemin du fichier ZIP final
            Dim CheminZip As String
            CheminZip = CheminDestination & "\" & NomRepertoire & "\" & NomFichier & Indice & ".zip"

            ' Dossier Archives pour sauvegarder les anciens ZIP de la même pièce avec un indice différent (ex: IndA-20260301.zip) sans les écraser,
            ' et supprimer uniquement le ZIP avec le même indice (ex: IndA-20260303.zip) pour le remplacer par le nouveau
            Dim CheminArchives As String
            CheminArchives = CheminDestination & "\" & NomRepertoire & "\Archives"

            ' Archivage des anciens ZIP
            ArchiverAnciensZip CheminDestination & "\" & NomRepertoire, CheminZip, CheminArchives, NomFichier

            ' Création du ZIP
            ZipFiles CheminTemp, CheminZip

            ' Suppression du dossier temporaire une fois le ZIP créé
            Dim FSO2 As Object
            Set FSO2 = CreateObject("Scripting.FileSystemObject")
            FSO2.DeleteFolder CheminTemp, True

            ' Ouverture du dossier contenant le ZIP dans l'explorateur Windows pour l'utilisateur
            Shell "EXPLORER /n,/e," & CheminDestination & "\" & NomRepertoire

            ' Note : ne fonctionne plus car le PDF est dans un ZIP
            ' Ouverture du fichier PDF dans le ZIP
            'Shell "explorer.exe """ & CheminZip & "\" & NomFichier & Indice & ".pdf" & """", vbNormalFocus

            ' Ouverture du CSV de diagnostic de la BOM pour vérifier les fichiers trouvés et introuvables
            If BOM_Trouvee Then
                Shell "explorer.exe """ & CheminCSV_BOM & """", vbNormalFocus
                Wait 1000
            End If

            Unload frm_Loading

            ' MsgBox finale : succès ou non + chemin du ZIP + type de document (assemblage, pièce avec BOM, pièce simple) + diagnostic sur la BOM si applicable
            Dim MsgFinale As String
            MsgFinale = "Dossier de lancement généré avec succès !" & vbCrLf & "----------------------------------------" & vbCrLf & _
                        NomRepertoire & vbCrLf & _
                        NomFichier & Indice & ".zip"

            ' Type de document détecté (assemblage, pièce avec BOM, pièce simple)
            Dim TypeDoc As String
            If swRefModel.GetType = swDocASSEMBLY Then
                TypeDoc = "Assemblage"
            ElseIf ContientBOM(swDraw) Then
                TypeDoc = "Pièce avec nomenclature (ex: soudure)"
            Else
                TypeDoc = "Pièce simple"
            End If
            MsgFinale = MsgFinale & vbCrLf & "----------------------------------------" & vbCrLf & "Type détecté : " & TypeDoc

            ' Diagnostic sur la BOM : si BOM trouvée, indiquer les fichiers SLDDRW trouvés et introuvables ;
            ' sinon si c'est un assemblage ou une pièce avec BOM mais que rien n'a été trouvé, indiquer qu'aucune nomenclature n'a été trouvée dans la mise en plan
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

' Fonction pour rechercher un sous-répertoire dans un chemin donné qui commence par un nom spécifique (ex: "12345" ou "12345 - Designation") et retourner son nom complet
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

' Fonction pour compresser en ZIP les fichiers d'un dossier source dans un fichier ZIP à un chemin donné
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

        ' Attendre que le ZIP soit terminé
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

' Vérifie que le ZIP appartient exactement à NomFichier sans variante
' Format attendu : NomFichier-YYYYMMDD.zip  ou  NomFichier-IndX-YYYYMMDD.zip
Function EstMemePiece(NomZip As String, NomFichier As String) As Boolean
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
Function LectureBOM(swApp As SldWorks.SldWorks, swDraw As SldWorks.DrawingDoc, ByVal swRefModel As Object, CheminPlan As String, CheminTemp As String, CheminDestination As String, ByRef CheminCSV_Out As String, ByRef Introuvables_Out As String) As Boolean

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
    TraiterLignesTable swApp, oTableRacine, IndexDRW, iCSV, DejaTraites, Introuvables, CheminTemp, CheminDestination

    Close #iCSV
    CheminCSV_Out = CheminCSV
    Introuvables_Out = Introuvables
    LectureBOM = True
End Function

' Fonction pour indexer tous les fichiers SLDDRW en mémoire pour accélérer les recherches ultérieures
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

' Fonction pour trouver le chemin d'un SLDDRW dans l'index en fonction du numéro de plan
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

' Foncion pour traiter les lignes d'une table de nomenclature de manière récursive (pour gérer les sous-assemblages) et remplir le CSV de diagnostic
Sub TraiterLignesTable(swApp As SldWorks.SldWorks, _
                       oTable As Object, _
                       IndexDRW As Collection, _
                       iCSV As Integer, _
                       DejaTraites As Collection, _
                       ByRef Introuvables As String, _
                       CheminTemp As String, _
                       CheminDestination As String)

    ' Vérification défensive avant d'utiliser la table
    If oTable Is Nothing Then Exit Sub

    ' Obtenir le nombre de lignes de la table, avec gestion d'erreur au cas où ce n'est pas une table valide
    Dim NbLignes As Long
    On Error Resume Next
    NbLignes = oTable.RowCount
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Parcours de chaque ligne de la table (en partant de 1 pour sauter l'en-tête)
    Dim r As Long
    For r = 1 To NbLignes - 1
        Dim NumPart  As String
        Dim NumPlan  As String
        Dim Design   As String
        NumPart = Trim(oTable.Text(r, 0))
        NumPlan = Trim(oTable.Text(r, 1))
        Design  = Trim(oTable.Text(r, 2))

        NumPlan = Join(Split(NumPlan, vbCrLf), " ")
        NumPlan = Join(Split(NumPlan, vbLf), " ")
        Design  = Join(Split(Design, vbCrLf), " ")
        Design  = Join(Split(Design, vbLf), " ")

        ' Filtrer les lignes sans numéro de plan ou avec un numéro de plan invalide pour éviter les recherches inutiles dans l'index et les erreurs
        If NumPlan = "" Then GoTo Suite
        If Not EstNumeroPlanValide(NumPlan) Then GoTo Suite

        ' Vérifier si ce numéro de plan a déjà été traité pour éviter les doublons et les boucles infinies en cas de références croisées dans la nomenclature
        Dim DejaVu As Boolean
        DejaVu = False
        On Error Resume Next
        Dim test As String
        test = DejaTraites(LCase(NumPlan))
        If Err.Number = 0 Then DejaVu = True
        Err.Clear
        On Error GoTo 0
        If DejaVu Then GoTo Suite

        ' Marquer ce numéro de plan comme traité
        DejaTraites.Add NumPlan, LCase(NumPlan)

        ' Trouver le chemin du SLDDRW correspondant à ce numéro de plan dans l'index
        Dim CheminDRW As String
        CheminDRW = TrouverDansIndex(IndexDRW, NumPlan)

        ' Ajouter une ligne dans le CSV de diagnostic avec le numéro de pièce, numéro de plan, désignation et chemin du SLDDRW trouvé ou [INTROUVABLE] si non trouvé
        Dim csv_Chemin As String
        If CheminDRW = "" Then
            csv_Chemin = "[INTROUVABLE]"
            Introuvables = Introuvables & "  " & NumPart & " | " & NumPlan & " | " & Design & vbCrLf
        Else
            csv_Chemin = CheminDRW
        End If

        Print #iCSV, NumPart & ";" & Chr(34) & NumPlan & Chr(34) & ";" & Design & ";" & csv_Chemin

        ' Si le SLDDRW est trouvé, l'ouvrir pour traiter la mise en plan du composant et faire la récursion sur sa propre nomenclature s'il y en a une, puis exporter PDF/DXF/STEP et inclure son ZIP dans celui du parent
        If CheminDRW <> "" Then

            Dim lErr As Long, lWarn As Long
            Dim swDrawSub As SldWorks.DrawingDoc
            Set swDrawSub = swApp.OpenDoc6(CheminDRW, swDocDRAWING, swOpenDocOptions_Silent, "", lErr, lWarn)

            ' Si l'ouverture du SLDDRW a réussi, on traite la mise en plan du composant ; sinon on passe à la ligne suivante de la table
            If Not swDrawSub Is Nothing Then

                ' Debug : lister les features de la mise en plan du composant pour vérifier que la BOM est bien lue
                Debug.Print "--- Features de : " & NumPlan
                Dim swFeatDbg As SldWorks.Feature
                Set swFeatDbg = swDrawSub.FirstFeature
                Do While Not swFeatDbg Is Nothing
                    Debug.Print "  Feature : " & swFeatDbg.GetTypeName2
                    Set swFeatDbg = swFeatDbg.GetNextFeature
                Loop
                Debug.Print "--- Fin features"
                ' Fin debug

                ' Chercher ou créer le dossier du composant dans CheminDestination
                Dim DossierComp As String
                DossierComp = RechercheSsRepCommençantPar(CheminDestination, NumPlan)
                
                ' Si le dossier du composant n'existe pas, le créer avec la désignation du composant (ex: "12345 - Designation") ; sinon on réutilise le même dossier
                If DossierComp = "" Then
                    Dim swViewComp  As SldWorks.View
                    Dim swRefComp   As Object
                    Dim cusPropComp As SldWorks.CustomPropertyManager
                    Dim DesigComp_brute As String
                    Dim DesigComp As String
                    Set swViewComp = swDrawSub.GetFirstView
                    Set swViewComp = swViewComp.GetNextView
                    If Not swViewComp Is Nothing Then
                        Set swRefComp = swViewComp.ReferencedDocument
                        If Not swRefComp Is Nothing Then
                            Set cusPropComp = swRefComp.Extension.CustomPropertyManager("")
                            cusPropComp.Get5 "Designation", True, DesigComp_brute, DesigComp, False
                        End If
                    End If
                    DesigComp = VBA.Strings.Trim(DesigComp)
                    If DesigComp = "" Then
                        DossierComp = NumPlan
                    Else
                        DossierComp = NumPlan & " - " & DesigComp
                    End If
                    On Error Resume Next
                    MkDir CheminDestination & "\" & DossierComp
                    If Err.Number <> 0 Then
                        ' Debug : afficher l'erreur de création du dossier si échec (ex: caractères interdits dans le nom de dossier, permissions, etc.)
                        Debug.Print "MkDir ECHEC : " & CheminDestination & "\" & DossierComp & " erreur=" & Err.Number & " " & Err.Description
                        Err.Clear
                        On Error GoTo 0
                        GoTo FermerDoc
                    End If
                    On Error GoTo 0
                End If

                ' Lire l'indice pour nommer le ZIP
                Dim cusPropZip      As SldWorks.CustomPropertyManager
                Dim IndiceZip_brute As String
                Dim IndiceZip       As String
                Set cusPropZip = swDrawSub.Extension.CustomPropertyManager("")
                cusPropZip.Get5 "Révision", False, IndiceZip_brute, IndiceZip, False
                IndiceZip = VBA.Strings.Trim(IndiceZip)

                Dim SuffixeZip As String
                If IndiceZip = "" Then
                    SuffixeZip = "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
                Else
                    SuffixeZip = "-Ind" & IndiceZip & "-" & VBA.Strings.Format(VBA.DateTime.Date, "YYYYMMDD")
                End If

                Dim CheminZipComp      As String
                Dim CheminArchivesComp As String
                Dim CheminTempComp     As String
                Dim FSO_Comp           As Object
                CheminZipComp      = CheminDestination & "\" & DossierComp & "\" & NumPlan & SuffixeZip & ".zip"
                CheminArchivesComp = CheminDestination & "\" & DossierComp & "\Archives"
                CheminTempComp     = CheminDestination & "\" & DossierComp & "\_temp_export"
                Set FSO_Comp = CreateObject("Scripting.FileSystemObject")

                ' Nettoyer et créer le dossier temporaire du composant (nécessaire dans tous les cas pour la récursion)
                If FSO_Comp.FolderExists(CheminTempComp) Then
                    On Error Resume Next
                    FSO_Comp.DeleteFolder CheminTempComp, True
                    On Error GoTo 0
                    Wait 500
                End If
                MkDir CheminTempComp

                ' Vérifier si un ZIP avec le même indice existe déjà
                Dim ZipExistant As String
                ZipExistant = TrouverZipMemeIndice(CheminDestination & "\" & DossierComp, NumPlan, IndiceZip)

                ' Chercher la BOM via ObtenirTable (BOM standard ou weldment)
                Dim oTableSub As Object
                Set oTableSub = ObtenirTable(swDrawSub)

                ' Debug : vérifier que la table est trouvée pour le composant
                Debug.Print "BOM trouvée pour " & NumPlan & " : " & Not (oTableSub Is Nothing)

                ' Si la BOM est trouvée et qu'il n'y a pas déjà un ZIP avec le même indice, on traite la BOM du composant de manière récursive pour inclure les plans des sous-composants dans le ZIP du composant, puis on exporte le PDF/DXF/STEP du composant lui-même et on crée son ZIP
                If Not oTableSub Is Nothing Then
                    TraiterLignesTable swApp, oTableSub, IndexDRW, iCSV, DejaTraites, Introuvables, CheminTempComp, CheminDestination
                End If

                ' 1. Exporter PDF/DXF/STEP du composant courant dans son dossier _temp_export
                ExporterMiseEnPlan swApp, swDrawSub, CheminTempComp

                ' 2. Archiver les anciens ZIP du composant
                ArchiverAnciensZip CheminDestination & "\" & DossierComp, CheminZipComp, CheminArchivesComp, NumPlan

                ' 3. Créer le ZIP du composant (contient PDF/DXF/STEP + ZIP des enfants)
                ZipFiles CheminTempComp, CheminZipComp

                ' 4. Copier ce ZIP dans le _temp_export du PARENT pour qu'il l'inclue dans son propre ZIP
                FSO_Comp.CopyFile CheminZipComp, CheminTemp & "\" & FSO_Comp.GetFileName(CheminZipComp), True

                ' 5. Nettoyer le dossier temporaire du composant
                FSO_Comp.DeleteFolder CheminTempComp, True

                ' Debug : afficher le composant traité et le ZIP copié dans le parent
                Debug.Print "Composant traité et ZIP copié dans parent : " & NumPlan

                ' Fermer le SLDDRW du composant avant de passer au suivant
                FermerDoc:
                swApp.CloseDoc swDrawSub.GetPathName
            End If
        End If

Suite:
    Next r
End Sub

' Fonction qui extrait la TableAnnotation d'un DrawingDoc (BOM standard ou weldment)
Function ObtenirTable(swDraw As SldWorks.DrawingDoc) As Object

    Dim swFeatTest As SldWorks.Feature
    Set swFeatTest = swDraw.FirstFeature
    Do While Not swFeatTest Is Nothing
        Debug.Print "FEAT : " & swFeatTest.GetTypeName2
        Set swFeatTest = swFeatTest.GetNextFeature
    Loop

    ' Debug : lister toutes les tables de toutes les vues
    Dim swViewDbg As SldWorks.View
    Set swViewDbg = swDraw.GetFirstView
    Dim vDbg As Integer
    vDbg = 0
    Do While Not swViewDbg Is Nothing
        Debug.Print "  Vue " & vDbg & " : " & swViewDbg.Name
        Dim vAnnotsDbg As Variant
        vAnnotsDbg = swViewDbg.GetTableAnnotations
        If IsEmpty(vAnnotsDbg) Or IsNull(vAnnotsDbg) Then
            Debug.Print "    -> GetTableAnnotations : vide/null"
        Else
            Dim kDbg As Long
            For kDbg = 0 To UBound(vAnnotsDbg)
                Dim oTDbg As Object
                Set oTDbg = vAnnotsDbg(kDbg)
                Dim tTypeDbg As Long
                On Error Resume Next
                tTypeDbg = oTDbg.TableType
                On Error GoTo 0
                Debug.Print "    -> Table type : " & tTypeDbg
            Next kDbg
        End If
        vDbg = vDbg + 1
        Set swViewDbg = swViewDbg.GetNextView
    Loop
    ' Fin debug

    ' Chercher BomFeat en priorité via les features
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

    ' Chercher WeldmentTableFeat : identifier la feature cible, puis retrouver
    ' la TableAnnotation correspondante dans les vues via GetFeature
    ' Chercher WeldmentTableFeat via les features
    Dim swFeatWeld As SldWorks.Feature
    Set swFeatWeld = swDraw.FirstFeature
    Do While Not swFeatWeld Is Nothing
        If swFeatWeld.GetTypeName2 = "WeldmentTableFeat" Then
            Dim swViewRoot As SldWorks.View
            Set swViewRoot = swDraw.GetFirstView
            Dim vAnnotsRoot As Variant
            vAnnotsRoot = swViewRoot.GetTableAnnotations
            If Not IsEmpty(vAnnotsRoot) And Not IsNull(vAnnotsRoot) Then
                Dim kW As Long
                Dim oBest As Object
                Dim nBestRows As Long
                nBestRows = 0
                For kW = 0 To UBound(vAnnotsRoot)
                    Dim oTW As Object
                    Set oTW = vAnnotsRoot(kW)
                    Dim nR As Long
                    On Error Resume Next
                    nR = oTW.RowCount
                    On Error GoTo 0
                    If nR > nBestRows Then
                        nBestRows = nR
                        Set oBest = oTW
                    End If
                Next kW
                If Not oBest Is Nothing Then
                    Set ObtenirTable = oBest
                    Exit Function
                End If
            End If
        End If
        Set swFeatWeld = swFeatWeld.GetNextFeature
    Loop

    Set ObtenirTable = Nothing
End Function

' Function pour vérifier que le numéro de plan extrait de la BOM est valide pour éviter les recherches inutiles dans l'index et les erreurs
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

' Function qui retourne le chemin du ZIP existant si même indice, "" sinon
Function TrouverZipMemeIndice(DossierComp As String, NumPlan As String, IndiceZip As String) As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    
    ' Construire le préfixe attendu selon l'indice
    Dim PrefixeRecherche As String
    If IndiceZip = "" Then
        PrefixeRecherche = LCase(NumPlan & "-")  ' NomFichier-YYYYMMDD.zip
    Else
        PrefixeRecherche = LCase(NumPlan & "-Ind" & IndiceZip & "-")  ' NomFichier-IndX-YYYYMMDD.zip
    End If

    For Each oFile In FSO.GetFolder(DossierComp).Files
        If LCase(FSO.GetExtensionName(oFile.Name)) = "zip" Then
            If Left(LCase(oFile.Name), Len(PrefixeRecherche)) = PrefixeRecherche Then
                TrouverZipMemeIndice = oFile.Path
                Exit Function
            End If
        End If
    Next oFile
    TrouverZipMemeIndice = ""
End Function