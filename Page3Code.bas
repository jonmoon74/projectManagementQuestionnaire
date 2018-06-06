Attribute VB_Name = "Page3Code"
'Suite of functions to handle entries on page three of assessment
Function RiskManagement() As Integer
If pagethree.threeoneone = True Then RiskManagement = background_data.Range("E4").Value
If pagethree.threeonetwo = True Then RiskManagement = background_data.Range("E5").Value
If pagethree.threeonethree = True Then RiskManagement = background_data.Range("E6").Value
If pagethree.threeonefour = True Then RiskManagement = background_data.Range("E7").Value
If RiskManagement < 2 Then RiskManagement = 1
MasterController.Range("C70") = RiskManagement
End Function

Function SchedulingAbility() As Integer
If pagethree.threetwoone = True Then SchedulingAbility = background_data.Range("E4").Value
If pagethree.threetwotwo = True Then SchedulingAbility = background_data.Range("E5").Value
If pagethree.threetwothree = True Then SchedulingAbility = background_data.Range("E6").Value
If pagethree.threetwofour = True Then SchedulingAbility = background_data.Range("E7").Value
If SchedulingAbility < 2 Then SchedulingAbility = 1
MasterController.Range("C71") = SchedulingAbility
End Function

Function BenefitId() As Integer
If pagethree.threethreeone = True Then BenefitId = background_data.Range("E4").Value
If pagethree.threethreetwo = True Then BenefitId = background_data.Range("E5").Value
If pagethree.threethreethree = True Then BenefitId = background_data.Range("E6").Value
If pagethree.threethreefour = True Then BenefitId = background_data.Range("E7").Value
If BenefitId < 2 Then BenefitId = 1
MasterController.Range("C72") = BenefitId
End Function

Function ResourceManagement() As Integer
If pagethree.threefourone = True Then ResourceManagement = background_data.Range("E4").Value
If pagethree.threefourtwo = True Then ResourceManagement = background_data.Range("E5").Value
If pagethree.threefourthree = True Then ResourceManagement = background_data.Range("E6").Value
If pagethree.threefourfour = True Then ResourceManagement = background_data.Range("E7").Value
If ResourceManagement < 2 Then ResourceManagement = 1
MasterController.Range("C73") = ResourceManagement
End Function

Function ReportingAbility() As Integer
If pagethree.threefiveone = True Then ReportingAbility = background_data.Range("E4").Value
If pagethree.threefivetwo = True Then ReportingAbility = background_data.Range("E5").Value
If pagethree.threefivethree = True Then ReportingAbility = background_data.Range("E6").Value
If pagethree.threefivefour = True Then ReportingAbility = background_data.Range("E7").Value
If ReportingAbility < 2 Then ReportingAbility = 1
MasterController.Range("C74") = ReportingAbility
End Function

Function ChangeManagement() As Integer
If pagethree.threesixone = True Then ChangeManagement = background_data.Range("E4").Value
If pagethree.threesixtwo = True Then ChangeManagement = background_data.Range("E5").Value
If pagethree.threesixthree = True Then ChangeManagement = background_data.Range("E6").Value
If pagethree.threesixfour = True Then ChangeManagement = background_data.Range("E7").Value
If ChangeManagement < 2 Then ChangeManagement = 1
MasterController.Range("C75") = ChangeManagement
End Function

Function ProcurementAbility() As Integer
If pagethree.threesevenone = True Then ProcurementAbility = background_data.Range("E4").Value
If pagethree.threeseventwo = True Then ProcurementAbility = background_data.Range("E5").Value
If pagethree.threeseventhree = True Then ProcurementAbility = background_data.Range("E6").Value
If pagethree.threesevenfour = True Then ProcurementAbility = background_data.Range("E7").Value
If ProcurementAbility < 2 Then ProcurementAbility = 1
MasterController.Range("C76") = ProcurementAbility
End Function

Function StakeholderManagement() As Integer
If pagethree.threeeightone = True Then StakeholderManagement = background_data.Range("E4").Value
If pagethree.threeeighttwo = True Then StakeholderManagement = background_data.Range("E5").Value
If pagethree.threeeightthree = True Then StakeholderManagement = background_data.Range("E6").Value
If pagethree.threeeightfour = True Then StakeholderManagement = background_data.Range("E7").Value
If StakeholderManagement < 2 Then StakeholderManagement = 1
MasterController.Range("C77") = StakeholderManagement
End Function

Function PandPManagement() As Integer
If pagethree.threenineone = True Then PandPManagement = background_data.Range("E4").Value
If pagethree.threeninetwo = True Then PandPManagement = background_data.Range("E5").Value
If pagethree.threeninethree = True Then PandPManagement = background_data.Range("E6").Value
If pagethree.threeninefour = True Then PandPManagement = background_data.Range("E7").Value
If PandPManagement < 2 Then PandPManagement = 1
MasterController.Range("C78") = PandPManagement
End Function
