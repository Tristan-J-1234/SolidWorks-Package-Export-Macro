Private Sub UserForm_Initialize()
    Me.Caption = "Génération en cours..."
    Me.Width = 320
    Me.Height = 120
    Me.StartUpPosition = 1 ' Centré sur le propriétaire
    
    With Me.Controls.Add("Forms.Label.1", "lblMessage")
        .Caption = "Export en cours, veuillez patienter..." & vbCrLf & vbCrLf & "Ne fermez pas SolidWorks."
        .Left = 10
        .Top = 10
        .Width = 280
        .Height = 60
        .TextAlign = 2 ' Centre
        .Font.Size = 10
    End With
End Sub
