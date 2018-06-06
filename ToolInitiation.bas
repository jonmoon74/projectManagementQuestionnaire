Attribute VB_Name = "ToolInitiation"
'Open workbook to correct page for progression
Sub ToolInitiation()
Dim c As Integer

c = 0
MasterController.Visible = xlVeryHidden
background_data.Visible = xlVeryHidden

Application.ScreenUpdating = False

If MasterController.Range("B3") = "TRUE" Then
        c = c + 1
        If MasterController.Range("B4") = "TRUE" Then
                c = c + 1
                If MasterController.Range("B5") = "TRUE" Then
                        c = c + 1
                        If MasterController.Range("B6") = "TRUE" Then
                                c = c + 1
                        End If
                End If
        End If
End If

If c = 0 Then
        instructions.Visible = xlSheetVisible
        instructions.Activate
        instructions.Range("A1").Select
        output_sheet.Visible = xlVeryHidden
        pageone.Visible = xlVeryHidden
        pagetwo.Visible = xlVeryHidden
        pagethree.Visible = xlVeryHidden
End If
If c = 1 Then
        instructions.Visible = xlVeryHidden
        output_sheet.Visible = xlVeryHidden
        pageone.Visible = xlSheetVisible
        pageone.Activate
        pageone.Range("C4").Select
        pagetwo.Visible = xlVeryHidden
        pagethree.Visible = xlVeryHidden
End If
If c = 2 Then
        instructions.Visible = xlVeryHidden
        output_sheet.Visible = xlVeryHidden
        pageone.Visible = xlVeryHidden
        pagetwo.Visible = xlSheetVisible
        pagetwo.Activate
        pagetwo.Range("C6").Select
        pagethree.Visible = xlVeryHidden
End If
If c = 3 Then
        instructions.Visible = xlVeryHidden
        output_sheet.Visible = xlVeryHidden
        pageone.Visible = xlVeryHidden
        pagetwo.Visible = xlVeryHidden
        pagethree.Visible = xlSheetVisible
        pagethree.Activate
        pagethree.Range("F6").Select
End If
If c = 4 Then
        instructions.Visible = xlVeryHidden
        output_sheet.Visible = xlSheetVisible
        output_sheet.Activate
        output_sheet.Range("A1").Select
        pageone.Visible = xlVeryHidden
        pagetwo.Visible = xlVeryHidden
        pagethree.Visible = xlVeryHidden
End If

Application.ScreenUpdating = True

End Sub
Function PMFCleanUp()
'Function to wipe any test data from all sheets and reset to initiation state
instructions.NameBox = ""
pageone.Range("C4") = ""
pageone.Range("C5") = ""
pageone.Range("C6") = ""
pageone.Range("C7") = ""
pageone.Range("C8") = ""
pageone.Range("C9") = ""
pageone.Range("C10") = ""
pageone.Range("C14") = ""
pageone.Range("C15") = ""
pageone.Range("C16") = ""
pageone.Range("C17") = ""
pageone.Range("C18") = ""
pageone.Range("C19") = ""
pageone.Range("C20") = ""
pageone.Range("C24") = ""
pageone.Range("C25") = ""
pageone.Range("C26") = ""
pageone.Range("C27") = ""
pageone.Range("C28") = ""
pagetwo.Range("C6") = ""
pagetwo.Range("C7") = ""
pagetwo.Range("C8") = ""
pagetwo.twooneone = False
pagetwo.twoonetwo = False
pagetwo.twoonethree = False
pagetwo.twoonefour = False
pagetwo.twotwoone = False
pagetwo.twotwotwo = False
pagetwo.twotwothree = False
pagetwo.twotwofour = False
pagetwo.twothreeone = False
pagetwo.twothreetwo = False
pagetwo.twothreethree = False
pagetwo.twothreefour = False
pagetwo.twofourone = False
pagetwo.twofourtwo = False
pagetwo.twofourthree = False
pagetwo.twofourfour = False
pagetwo.twofiveone = False
pagetwo.twofivetwo = False
pagetwo.twofivethree = False
pagetwo.twofivefour = False
pagetwo.twosixone = False
pagetwo.twosixtwo = False
pagetwo.twosixthree = False
pagetwo.twosixfour = False
pagetwo.twosevenone = False
pagetwo.twoseventwo = False
pagetwo.twoseventhree = False
pagetwo.twosevenfour = False
pagetwo.twoeightone = False
pagetwo.twoeighttwo = False
pagetwo.twoeightthree = False
pagetwo.twoeightfour = False
pagetwo.twonineone = False
pagetwo.twoninetwo = False
pagetwo.twoninethree = False
pagetwo.twoninefour = False
pagetwo.twotenone = False
pagetwo.twotentwo = False
pagetwo.twotenthree = False
pagetwo.twotenfour = False
pagethree.Range("F6") = ""
pagethree.Range("F7") = ""
pagethree.Range("F8") = ""
pagethree.Range("F9") = ""
pagethree.Range("F10") = ""
pagethree.Range("F11") = ""
pagethree.Range("F12") = ""
pagethree.Range("F13") = ""
pagethree.Range("F14") = ""
pagethree.Range("F15") = ""
pagethree.threeoneone = False
pagethree.threeonetwo = False
pagethree.threeonethree = False
pagethree.threeonefour = False
pagethree.threetwoone = False
pagethree.threetwotwo = False
pagethree.threetwothree = False
pagethree.threetwofour = False
pagethree.threethreeone = False
pagethree.threethreetwo = False
pagethree.threethreethree = False
pagethree.threethreefour = False
pagethree.threefourone = False
pagethree.threefourtwo = False
pagethree.threefourthree = False
pagethree.threefourfour = False
pagethree.threefiveone = False
pagethree.threefivetwo = False
pagethree.threefivethree = False
pagethree.threefivefour = False
pagethree.threesixone = False
pagethree.threesixtwo = False
pagethree.threesixthree = False
pagethree.threesixfour = False
pagethree.threesevenone = False
pagethree.threeseventwo = False
pagethree.threeseventhree = False
pagethree.threesevenfour = False
pagethree.threeeightone = False
pagethree.threeeighttwo = False
pagethree.threeeightthree = False
pagethree.threeeightfour = False
pagethree.threenineone = False
pagethree.threeninetwo = False
pagethree.threeninethree = False
pagethree.threeninefour = False
MasterController.Range("B3") = "False"
MasterController.Range("B4") = "False"
MasterController.Range("B5") = "False"
MasterController.Range("B6") = "False"
MasterController.Range("H101") = ""
MasterController.Range("H102") = ""
MasterController.Range("H103") = ""
MasterController.Range("H104") = ""
MasterController.Range("H105") = ""
MasterController.Range("H106") = ""
MasterController.Range("H107") = ""
MasterController.Range("H108") = ""
MasterController.Range("H109") = ""
MasterController.Range("H110") = ""
MasterController.Range("H111") = ""
MasterController.Range("H112") = ""
MasterController.Range("H113") = ""
MasterController.Range("H114") = ""
End Function
