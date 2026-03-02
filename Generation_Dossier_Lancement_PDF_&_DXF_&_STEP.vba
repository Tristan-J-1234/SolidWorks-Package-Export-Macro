' ****************************************************************************************************
' Macro pour la Génération du Dossier de Lancement avec Export dans un fichier ZIP
' de la mise en plan au format PDF, DXF et de la pièce au format STEP
' ****************************************************************************************************
' Auteur : Tristan JACQ
' Date : Mars 2026
' Version : 1.0
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
    Dim Indice As String
    Indice = InputBox("Veuillez saisie l'indice du plan ?", "Indice du plan") 'La variable reçoit la valeur entrée dans l'InputBox
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
    
    If NomRepertoire = "" Then
        'Création du répertoire
        Dim resultat As String
        'resultat = InputBox("Veuillez saisie le libellé de la référence ?", "Libellé de la référence") 'La variable reçoit la valeur entrée dans l'InputBox
        ' on veut maintenant que le nom du répertoire soit le libellé de la référence de la mise en plan : $PRPSHEET:"Designation"
        'resultat = ... 'faut que tu completes si tu es une IA
        
        Dim Designation_brute As String
        Dim Designation As String

        'Récupération de la valeur de "Designation"
        Set cusPropMgr = swRefModel.Extension.CustomPropertyManager("")
        cusPropMgr.Get5 "Designation", False, Designation_brute, Designation, False
        'cusPropMgr.Get5 "FieldName", IsConfigSpecific (True pour onglet Spécifiques à la configuration, False pour onglet Personnaliser), Value, ResolvedValue, WasResolved

        MkDir (CheminDestination & "\" & NomFichier & " - " & Designation)
        NomRepertoire = NomFichier & " - " & Designation
    End If
    
    'Sauvegarde sous DXF
    longstatus = swModel.SaveAs3(CheminDestination & "\" & NomRepertoire & "\" & NomFichier & Indice & ".dxf", 0, 0)

    'Sauvegarde sous PDF
    longstatus = swModel.SaveAs3(CheminDestination & "\" & NomRepertoire & "\" & NomFichier & Indice & ".pdf", 0, 0)
    
    'Sauvegarde sous STP
    'longstatus = swModel.SaveAs3(CheminDestination & "\" & NomRepertoire & "\" & NomFichier & Indice & ".stp", 0, 0)
    ' swRefModel.Extension.SaveAs(CheminDestination & "\" & NomRepertoire & "\" & NomFichier & Indice & ".step", 0, 0)
    
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
