Attribute VB_Name = "Page2Code"
'Suite of functions to handle entries on page two of assessment
Function PastStakeholders() As Integer
If pagetwo.twooneone = True Then PastStakeholders = background_data.Range("E4").Value
If pagetwo.twoonetwo = True Then PastStakeholders = background_data.Range("E5").Value
If pagetwo.twoonethree = True Then PastStakeholders = background_data.Range("E6").Value
If pagetwo.twoonefour = True Then PastStakeholders = background_data.Range("E7").Value
If PastStakeholders < 2 Then PastStakeholders = 1
MasterController.Range("C40") = PastStakeholders
End Function

Function PastDependancies() As Integer
If pagetwo.twotwoone = True Then PastDependancies = background_data.Range("E4").Value
If pagetwo.twotwotwo = True Then PastDependancies = background_data.Range("E5").Value
If pagetwo.twotwothree = True Then PastDependancies = background_data.Range("E6").Value
If pagetwo.twotwofour = True Then PastDependancies = background_data.Range("E7").Value
If PastDependancies < 2 Then PastDependancies = 1
MasterController.Range("C41") = PastDependancies
End Function

Function FutureStakeholders() As Integer
If pagetwo.twothreeone = True Then FutureStakeholders = background_data.Range("E4").Value
If pagetwo.twothreetwo = True Then FutureStakeholders = background_data.Range("E5").Value
If pagetwo.twothreethree = True Then FutureStakeholders = background_data.Range("E6").Value
If pagetwo.twothreefour = True Then FutureStakeholders = background_data.Range("E7").Value
If FutureStakeholders < 2 Then FutureStakeholders = 1
MasterController.Range("C42") = FutureStakeholders
End Function

Function FutureDependancies() As Integer
If pagetwo.twofourone = True Then FutureDependancies = background_data.Range("E4").Value
If pagetwo.twofourtwo = True Then FutureDependancies = background_data.Range("E5").Value
If pagetwo.twofourthree = True Then FutureDependancies = background_data.Range("E6").Value
If pagetwo.twofourfour = True Then FutureDependancies = background_data.Range("E7").Value
If FutureDependancies < 2 Then FutureDependancies = 1
MasterController.Range("C43") = FutureDependancies
End Function

Function FormalProcesses() As Integer
If pagetwo.twofiveone = True Then FormalProcesses = background_data.Range("E4").Value
If pagetwo.twofivetwo = True Then FormalProcesses = background_data.Range("E5").Value
If pagetwo.twofivethree = True Then FormalProcesses = background_data.Range("E6").Value
If pagetwo.twofivefour = True Then FormalProcesses = background_data.Range("E7").Value
If FormalProcesses < 2 Then FormalProcesses = 1
MasterController.Range("C45") = FormalProcesses
End Function

Function StableProcesses() As Integer
If pagetwo.twosixone = True Then StableProcesses = background_data.Range("E4").Value
If pagetwo.twosixtwo = True Then StableProcesses = background_data.Range("E5").Value
If pagetwo.twosixthree = True Then StableProcesses = background_data.Range("E6").Value
If pagetwo.twosixfour = True Then StableProcesses = background_data.Range("E7").Value
If StableProcesses < 2 Then StableProcesses = 1
MasterController.Range("C47") = StableProcesses
End Function

Function WideProcesses() As Integer
If pagetwo.twosevenone = True Then WideProcesses = background_data.Range("E4").Value
If pagetwo.twoseventwo = True Then WideProcesses = background_data.Range("E5").Value
If pagetwo.twoseventhree = True Then WideProcesses = background_data.Range("E6").Value
If pagetwo.twosevenfour = True Then WideProcesses = background_data.Range("E7").Value
If WideProcesses < 2 Then WideProcesses = 1
MasterController.Range("C49") = WideProcesses
End Function

Function TypicalTools() As Integer
If pagetwo.twoeightone = True Then TypicalTools = background_data.Range("E4").Value
If pagetwo.twoeighttwo = True Then TypicalTools = background_data.Range("E5").Value
If pagetwo.twoeightthree = True Then TypicalTools = background_data.Range("E6").Value
If pagetwo.twoeightfour = True Then TypicalTools = background_data.Range("E7").Value
If TypicalTools < 2 Then TypicalTools = 1
MasterController.Range("C51") = TypicalTools
End Function

Function StandardAppetite() As Integer
If pagetwo.twonineone = True Then StandardAppetite = background_data.Range("E4").Value
If pagetwo.twoninetwo = True Then StandardAppetite = background_data.Range("E5").Value
If pagetwo.twoninethree = True Then StandardAppetite = background_data.Range("E6").Value
If pagetwo.twoninefour = True Then StandardAppetite = background_data.Range("E7").Value
If StandardAppetite < 2 Then StandardAppetite = 1
MasterController.Range("C53") = StandardAppetite
End Function

Function CurrentTraining() As Integer
If pagetwo.twotenone = True Then CurrentTraining = background_data.Range("E4").Value
If pagetwo.twotentwo = True Then CurrentTraining = background_data.Range("E5").Value
If pagetwo.twotenthree = True Then CurrentTraining = background_data.Range("E6").Value
If pagetwo.twotenfour = True Then CurrentTraining = background_data.Range("E7").Value
If CurrentTraining < 2 Then CurrentTraining = 1
MasterController.Range("C55") = CurrentTraining
End Function
