VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_isDebug As Boolean

Private Sub Class_Initialize()
    m_isDebug = True
End Sub

Private Sub Class_Terminate()
End Sub

Public Property Let DebugMode(ByVal aIsDebug As Boolean)
    'this‚ª—˜—p‚Å‚«‚È‚¢
    m_isDebug = aIsDebug
End Property

Public Sub trace(ByVal message As String)
   
    If m_isDebug = True Then
        Debug.Print message
    End If
End Sub

Public Sub info(ByVal message As String)
    Debug.Print message
End Sub

