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

'�R���X�g���N�^(�����n���s�\)
Private Sub Class_Initialize()
    Set m_log = Nothing '�����I�ȏ�����
End Sub

'�f�X�g���N�^
Private Sub Class_Terminate()
    ' Set m_log = Nothing ������new �����킯�ł͂Ȃ��̂ŌĂяo�������Ő���
End Sub

'�v���p�e�B
Public Property Let LogObj(ByRef aLogUtil As LogUtil)
    Set m_log = aLogUtil
End Property




'excel�̑I���Z���i�����̃Z���j�ɓ����Ă���e�L�X�g�v�f���ꎟ���̉ϒ�������List�ɂ��ĕԋp���܂��B
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
    Call m_log.trace("�I���Z�����ꎟ���z��")
    Call m_log.trace(toString(result))
    

    Set SelectionToCollection = result


End Function

'�R���N�V�����v�f���̊e���ڂɑ΂��āAsplitWord�������ĕ���������̈ꎟ���̉ϒ�List�ɂ��ĕԋp���܂�
Public Function SplitBy(ByRef aCollection As collection, ByVal splitWord As String) As collection
    Dim result As collection
    Set result = New collection

    Dim value As Variant
    For Each value In aCollection
        Dim arr As Variant
        Dim i As Integer
        arr = Split(value, splitWord)
        For i = LBound(arr) To UBound(arr)
            ' ���������� �X�v���b�g�P�ꂪ����ꍇ�̑Ή�
            If arr(i) <> Space$(0) Then
                result.Add arr(i)
            End If
        Next i
    Next value
    
    Call m_log.trace("----------------------------")
    Call m_log.trace("�R���N�V�������̊e�v�f��split���Ĉꎟ��List��:" & splitWord)
    Call m_log.trace(toString(result))
    
    
    Set SplitBy = result
End Function

'�R���N�V�����v�f���̊e���ڂɑ΂��āAoldChar��newChar�ɕύX������Ɉꎟ���̉ϒ�List�ɂ��ĕԋp���܂�
Public Function Sanitize(ByRef aCollection As collection, ByVal oldChar As String, ByVal newChar As String) As collection

    Dim result As collection
    Set result = New collection
    Dim value As Variant
    For Each value In aCollection
        Dim tmp As String
        tmp = Replace(value, oldChar, newChar)
        result.Add tmp
    Next value
    
    Call m_log.trace("replace����  oldChar:" & oldChar & ", newChar=" & newChar)
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
