Attribute VB_Name = "FromQuiita"
Option Explicit

Private Const MODULE_NAME As String = "FromQuiita"

Public Function C() As String

    On Error GoTo ErrorHandler

    
    Dim result As String
    result = B
    C = "Bでの結果は" & result & "です"

    GoTo LogicEnd
ErrorHandler:
    Call sendError(Err, MODULE_NAME, "C", extraInfo:="Cで何も渡していない")
LogicEnd:

End Function

Public Function B() As String

    On Error GoTo ErrorHandler
    
    B = ""
    Dim divided As Integer
    Dim divisor  As Integer
    Dim result As Integer
    divided = 2
    divisor = 0
    result = 0
        
    result = A(divided, divisor)
    B = "計算結果は" & result & "です"
    GoTo LogicEnd
ErrorHandler:

    Call sendError(Err, MODULE_NAME, "B", extraInfo:="Bでのdivided=" & divided & ", divisor:" & divisor)


LogicEnd:
    
    
End Function



Public Function A(ByVal divided As Integer, ByVal divisor As Integer) As Integer
    On Error GoTo ErrorHandler

    A = 0 ' 初期値
    A = divided / divisor
    
    GoTo LogicEnd
ErrorHandler:

    Call sendError(Err, MODULE_NAME, "A")

LogicEnd:


End Function


Public Sub sendError(ByRef preError As ErrObject, ByVal moduleName As String, procedureName As String, Optional extraInfo As String = "")

    Dim myerr As MyErrObject
    Set myerr = New MyErrObject
    myerr.setProperties
    preError.Clear
    Call myerr.raiseError(moduleName:=moduleName, procedureName:=procedureName, extraInfo:=extraInfo)
End Sub
