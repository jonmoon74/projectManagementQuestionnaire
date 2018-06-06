Attribute VB_Name = "ResultCompiling"
Sub ResultCompilation()
Call M1Result
Call M2Result
Call M3Result
Call M4Result
Call M5Result
Call M6Result
Call M7Result
Call M8Result
Call M9Result
Call M10Result
Call M11Result
Call M12Result
Call M13Result
Call M14Result

End Sub


Function M1Result()
Dim resultLevel1 As String, result1 As Integer
'Project controls
result1 = MasterController.Range("C101").Value
Select Case result1
    Case Is < 0
        resultLevel1 = "Exempt"
    Case 0 To 25
        resultLevel1 = "Basic"
    Case 26 To 45
        resultLevel1 = "Intermediate"
    Case Is > 45
        resultLevel1 = "Advanced"
End Select
MasterController.Range("H101") = resultLevel1
End Function

Function M2Result()
Dim resultLevel2 As String, result2 As Integer
'Scope management
result2 = MasterController.Range("C102").Value
Select Case result2
    Case Is < 0
        resultLevel2 = "Exempt"
    Case 0 To 6
        resultLevel2 = "Basic"
    Case 7 To 15
        resultLevel2 = "Intermediate"
    Case Is > 15
        resultLevel2 = "Advanced"
End Select
MasterController.Range("H102") = resultLevel2
End Function

Function M3Result()
Dim resultLevel3 As String, result3 As Integer
'Time management
result3 = MasterController.Range("C103").Value
Select Case result3
    Case Is < 0
        resultLevel3 = "Exempt"
    Case 0 To 9
        resultLevel3 = "Basic"
    Case 10 To 19
        resultLevel3 = "Intermediate"
    Case Is > 19
        resultLevel3 = "Advanced"
End Select
MasterController.Range("H103") = resultLevel3
End Function

Function M4Result()
Dim resultLevel4 As String, result4 As Integer
'Cost management
result4 = MasterController.Range("C104").Value
Select Case result4
    Case Is < 0
        resultLevel4 = "Exempt"
    Case 0 To 9
        resultLevel4 = "Basic"
    Case 10 To 19
        resultLevel4 = "Intermediate"
    Case Is > 19
        resultLevel4 = "Advanced"
End Select
MasterController.Range("H104") = resultLevel4
End Function

Function M5Result()
Dim resultLevel5 As String, result5 As Integer
'Quality management
result5 = MasterController.Range("C105").Value
Select Case result5
    Case Is < 0
        resultLevel5 = "Exempt"
    Case 0 To 8
        resultLevel5 = "Basic"
    Case 9 To 19
        resultLevel5 = "Intermediate"
    Case Is > 19
        resultLevel5 = "Advanced"
End Select
MasterController.Range("H105") = resultLevel5
End Function

Function M6Result()
Dim resultLevel6 As String, result6 As Integer
'resource management
result6 = MasterController.Range("C106").Value
Select Case result6
    Case Is < 0
        resultLevel6 = "Exempt"
    Case 0 To 8
        resultLevel6 = "Basic"
    Case 9 To 19
        resultLevel6 = "Intermediate"
    Case Is > 19
        resultLevel6 = "Advanced"
End Select
MasterController.Range("H106") = resultLevel6
End Function

Function M7Result()
Dim resultLevel7 As String, result7 As Integer
'communications management
result7 = MasterController.Range("C107").Value
Select Case result7
    Case Is < 0
        resultLevel7 = "Exempt"
    Case 0 To 8
        resultLevel7 = "Basic"
    Case 9 To 19
        resultLevel7 = "Intermediate"
    Case Is > 19
        resultLevel7 = "Advanced"
End Select
MasterController.Range("H107") = resultLevel7
End Function

Function M8Result()
Dim resultLevel8 As String, result8 As Integer
'risk management
result8 = MasterController.Range("C108").Value
Select Case result8
    Case Is < 0
        resultLevel8 = "Exempt"
    Case 0 To 8
        resultLevel8 = "Basic"
    Case 9 To 19
        resultLevel8 = "Intermediate"
    Case Is > 19
        resultLevel8 = "Advanced"
End Select
MasterController.Range("H108") = resultLevel8
End Function

Function M9Result()
Dim resultLevel9 As String, result9 As Integer
'procurement management
result9 = MasterController.Range("C109").Value
Select Case result9
    Case Is < 0
        resultLevel9 = "Exempt"
    Case 0 To 9
        resultLevel9 = "Basic"
    Case 10 To 19
        resultLevel9 = "Intermediate"
    Case Is > 19
        resultLevel9 = "Advanced"
End Select
MasterController.Range("H109") = resultLevel9
End Function

Function M10Result()
Dim resultLevel10 As String, result10 As Integer
'governance and stakeholder management
result10 = MasterController.Range("C110").Value
Select Case result10
    Case Is < 0
        resultLevel10 = "Exempt"
    Case 0 To 10
        resultLevel10 = "Basic"
    Case 11 To 20
        resultLevel10 = "Intermediate"
    Case Is > 20
        resultLevel10 = "Advanced"
End Select
MasterController.Range("H110") = resultLevel10
End Function

Function M11Result()
Dim resultLevel11 As String, result11 As Integer
'awareness of general PM methodologies
result11 = MasterController.Range("C111").Value
Select Case result11
    Case Is < 0
        resultLevel11 = "Exempt"
    Case 0 To 41
        resultLevel11 = "Basic"
    Case 42 To 60
        resultLevel11 = "Intermediate"
    Case Is > 60
        resultLevel11 = "Advanced"
End Select
MasterController.Range("H111") = resultLevel11
End Function

Function M12Result()
Dim resultLevel12 As String, result12 As Integer
'tools
result12 = MasterController.Range("C112").Value
Select Case result12
    Case Is < 0
        resultLevel12 = "Exempt"
    Case 0 To 45
        resultLevel12 = "Basic"
    Case 46 To 56
        resultLevel12 = "Intermediate"
    Case Is > 56
        resultLevel12 = "Advanced"
End Select
MasterController.Range("H112") = resultLevel12
End Function

Function M13Result()
Dim resultLevel13 As String, result13 As Integer
'soft skills
result13 = MasterController.Range("C113").Value
Select Case result13
    Case Is < 0
        resultLevel13 = "Exempt"
    Case 0 To 10
        resultLevel13 = "Basic"
    Case 11 To 20
        resultLevel13 = "Intermediate"
    Case Is > 20
        resultLevel13 = "Advanced"
End Select
MasterController.Range("H113") = resultLevel13
End Function

Function M14Result()
Dim resultLevel14 As String, result14 As Integer
'programme and portfolio management
result14 = MasterController.Range("C114").Value
Select Case result14
    Case Is < 0
        resultLevel14 = "Exempt"
    Case 0 To 16
        resultLevel14 = "Basic"
    Case 17 To 45
        resultLevel14 = "Intermediate"
    Case Is > 45
        resultLevel14 = "Advanced"
End Select
MasterController.Range("H114") = resultLevel14
End Function
