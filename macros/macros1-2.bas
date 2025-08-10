Attribute VB_Name = "Module2"
Function цсяебю(x As Double) As Double
    цсяебю = ((Sin(Cos(Application.WorksheetFunction.Pi() * x)) / x) - Sin(Application.WorksheetFunction.Pi() * x))
End Function
