Option Explicit

' ### enter your Settings here  ###

Function RetrieveVar(VarName as String)
    Select Case VarName
        Case "Field"
            RetrieveVar = 1
        Case "Table"
            RetrieveVar = "Table1"
        Case "Filter"
            RetrieveVar = "[Name] = '"
        Case "TrgtRange"
            RetrieveVar= "A1"
    End Select
End Function
