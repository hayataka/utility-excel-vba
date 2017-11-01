VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyErrObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    ' エラー退避用オブジェクト
    
    Public Source As String
    Public Description As String
    Public Number As Integer
    Public preHelpContext As Integer
    Public preHelpFile As String
    Public HelpContext As String
    Public HelpFile As String
    

Public Sub setProperties()
    Me.Source = Err.Source
    Me.Description = Err.Description
    Me.HelpContext = Err.HelpContext
    Me.HelpFile = Err.HelpFile
    Me.Number = Err.Number
End Sub

Public Sub raiseError(ByVal moduleName As String, procedureName As String, Optional extraInfo As String = "")
 
    Dim desc As String
    If extraInfo <> "" Then
        desc = Me.Description & extraInfo & ","
    Else
        desc = Me.Description & ","
    End If
 
    Call Err.Raise(Number:=Me.Number, Source:=moduleName & "." & procedureName & "|" & Me.Source, _
           Description:=desc, _
           HelpContext:=Me.HelpContext, _
           HelpFile:=Me.HelpFile)


End Sub
