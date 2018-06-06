Attribute VB_Name = "NavigationControls"
Sub ShowPageOne()
pageone.Visible = xlSheetVisible
instructions.Visible = xlSheetVeryHidden
output_sheet.Visible = xlSheetVeryHidden
pagetwo.Visible = xlSheetVeryHidden
pagethree.Visible = xlSheetVeryHidden
background_data.Visible = xlSheetVeryHidden
MasterController.Visible = xlSheetVeryHidden


pageone.Activate
pageone.Range("C4").Select
MasterController.Range("B3") = "True"
End Sub

Sub ShowPageTwo()
pagetwo.Visible = xlSheetVisible
instructions.Visible = xlSheetVeryHidden
output_sheet.Visible = xlSheetVeryHidden
pageone.Visible = xlSheetVeryHidden
pagethree.Visible = xlSheetVeryHidden
background_data.Visible = xlSheetVeryHidden
MasterController.Visible = xlSheetVeryHidden


pagetwo.Activate
pagetwo.Range("C6").Select
MasterController.Range("B4") = "True"
End Sub

Sub ShowPageThree()
pagethree.Visible = xlSheetVisible
instructions.Visible = xlSheetVeryHidden
output_sheet.Visible = xlSheetVeryHidden
pagetwo.Visible = xlSheetVeryHidden
pageone.Visible = xlSheetVeryHidden
background_data.Visible = xlSheetVeryHidden
MasterController.Visible = xlSheetVeryHidden


pagethree.Activate
pagethree.Range("F6").Select
MasterController.Range("B5") = "True"
End Sub

Sub ShowResults()
output_sheet.Visible = xlSheetVisible
instructions.Visible = xlSheetVeryHidden
pageone.Visible = xlSheetVeryHidden
pagetwo.Visible = xlSheetVeryHidden
pagethree.Visible = xlSheetVeryHidden
background_data.Visible = xlSheetVeryHidden
MasterController.Visible = xlSheetVeryHidden


output_sheet.Activate
output_sheet.Range("A1").Select
MasterController.Range("B6") = "True"
End Sub

Sub ShowAll()

Dim Password As String

    Password = InputBox("Please enter password below", "Password", "????")
        If Password <> "Pa55word" Then
			MsgBox "Incorrect Password"
			Exit Sub
			Else
		End If

instructions.Visible = xlSheetVisible
pageone.Visible = xlSheetVisible
pagetwo.Visible = xlSheetVisible
pagethree.Visible = xlSheetVisible
background_data.Visible = xlSheetVisible
MasterController.Visible = xlSheetVisible
output_sheet.Visible = xlSheetVisible
MasterController.Range("B3") = "False"
MasterController.Range("B4") = "False"
MasterController.Range("B5") = "False"
MasterController.Range("B6") = "False"
End Sub

Sub GetStarted()
If instructions.NameBox.Value = "" Then
    MsgBox "Please enter your name to begin", vbOKOnly, "Error"
    Exit Sub
Else
    Application.Run "ShowPageOne"
End If
End Sub

