VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_log As LogUtil

'コンストラクタ(引数渡し不可能)
Private Sub Class_Initialize()
    Set m_log = Nothing '明示的な初期化
End Sub

'デストラクタ
Private Sub Class_Terminate()
    ' Set m_log = Nothing 自分でnew したわけではないので呼び出し元側で制御
End Sub

'プロパティ
Public Property Let LogObj(ByRef aLogUtil As LogUtil)
    Set m_log = aLogUtil
End Property




'excelの選択セル（複数のセル）に入っているテキスト要素を一次元の可変長文字列Listにして返却します。
Public Function SelectionToCollection(ByRef selectedCells As Object) As collection

    Dim allRange As Range
    Dim oneCell As Range
    
    Dim result As collection
    Set result = New collection
    
    Set allRange = selectedCells
    For Each oneCell In allRange
        result.Add oneCell.Cells.value
    Next oneCell

    Call m_log.trace("----------------------------")
    Call m_log.trace("選択セルを一次元配列化")
    Call m_log.trace(toString(result))
    

    Set SelectionToCollection = result


End Function

'コレクション要素内の各項目に対して、splitWordをかけて分割した後の一次元の可変長Listにして返却します
Public Function SplitBy(ByRef aCollection As collection, ByVal splitWord As String) As collection
    Dim result As collection
    Set result = New collection

    Dim value As Variant
    For Each value In aCollection
        Dim arr As Variant
        Dim i As Integer
        arr = Split(value, splitWord)
        For i = LBound(arr) To UBound(arr)
            ' 複数続けて スプリット単語がある場合の対応
            If arr(i) <> Space$(0) Then
                result.Add arr(i)
            End If
        Next i
    Next value
    
    Call m_log.trace("----------------------------")
    Call m_log.trace("コレクション内の各要素をsplitして一次元List化:" & splitWord)
    Call m_log.trace(toString(result))
    
    
    Set SplitBy = result
End Function

'コレクション要素内の各項目に対して、oldCharをnewCharに変更した後に一次元の可変長Listにして返却します
Public Function Sanitize(ByRef aCollection As collection, ByVal oldChar As String, ByVal newChar As String) As collection

    Dim result As collection
    Set result = New collection
    Dim value As Variant
    For Each value In aCollection
        Dim tmp As String
        tmp = Replace(value, oldChar, newChar)
        result.Add tmp
    Next value
    
    Call m_log.trace("replace処理  oldChar:" & oldChar & ", newChar=" & newChar)
    Call m_log.trace(toString(result))
    
    Set Sanitize = result
End Function

Public Function toString(ByRef aCollection As collection) As String

    Dim result As String
    result = ""

    If Not aCollection Is Nothing Then
        Dim value As Variant
        For Each value In aCollection
            result = result & value & vbCrLf
        Next value
    End If

    toString = result

End Function

