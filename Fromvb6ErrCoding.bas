Attribute VB_Name = "Fromvb6ErrCoding"
Option Explicit

'「VB6エラーコーディング」という書籍からの案を紹介する用途

'エラーメッセージの泡(StackTraceですよね)

' 最終形を (StackTrace相当とするために)
'errMessage = "エラーメッセージA|エラーメッセージB|エラーメッセージC"
' とする
'例：  C→B→Aと関数呼び出しを行い、Aでエラーが発生した場合に
'エラーメッセージA＝"ルーチンA：0で除算しました"
'エラーメッセージB＝"ルーチンB：不正な値が入力されました。値：0、１"
'エラーメッセージC＝"ルーチンC：入力された値は無効です。"


Public Function C(ByRef ret As String) As String

    C = ""  '初期値
    
    Dim errMessage As String
    errMessage = B(ret)
    If errMessage <> "" Then
        C = errMessage & "| ルーチンC：入力された値は無効です"
            
    End If


End Function

Private Function B(ByRef ret As String) As String

    On Error Resume Next
    B = ""

    Dim divided As Integer
    Dim divisor  As Integer
    Dim result As Integer
    divided = 2
    divisor = 0
    result = 0
        
    Dim errMessage As String
    errMessage = A(divided, divisor, result)
    If errMessage <> "" Then
             B = errMessage & "| ルーチンB：不正な値が入力されました divided:" & divided & ",  divisor:" & divisor
    End If
    ret = "計算結果は" & result & "です"
    
End Function



Private Function A(ByVal divided As Integer, ByVal divisor As Integer, ByRef result As Integer) As String
    On Error Resume Next

    A = "" ' 初期値
    result = divided / divisor
    If Err.Number <> 0 Then
        A = "ルーチンA：" & Err.Description
    End If
    On Error GoTo 0

End Function

