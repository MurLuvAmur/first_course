Attribute VB_Name = "Module2"
Function ������(x As Double) As Double
    ������ = ((Sin(Cos(Application.WorksheetFunction.Pi() * x)) / x) - Sin(Application.WorksheetFunction.Pi() * x))
End Function
