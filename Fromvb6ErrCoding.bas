Attribute VB_Name = "Fromvb6ErrCoding"
Option Explicit

'�uVB6�G���[�R�[�f�B���O�v�Ƃ������Ђ���̈Ă��Љ��p�r

'�G���[���b�Z�[�W�̖A(StackTrace�ł����)

' �ŏI�`�� (StackTrace�����Ƃ��邽�߂�)
'errMessage = "�G���[���b�Z�[�WA|�G���[���b�Z�[�WB|�G���[���b�Z�[�WC"
' �Ƃ���
'��F  C��B��A�Ɗ֐��Ăяo�����s���AA�ŃG���[�����������ꍇ��
'�G���[���b�Z�[�WA��"���[�`��A�F0�ŏ��Z���܂���"
'�G���[���b�Z�[�WB��"���[�`��B�F�s���Ȓl�����͂���܂����B�l�F0�A�P"
'�G���[���b�Z�[�WC��"���[�`��C�F���͂��ꂽ�l�͖����ł��B"


Public Function C(ByRef ret As String) As String

    C = ""  '�����l
    
    Dim errMessage As String
    errMessage = B(ret)
    If errMessage <> "" Then
        C = errMessage & "| ���[�`��C�F���͂��ꂽ�l�͖����ł�"
            
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
             B = errMessage & "| ���[�`��B�F�s���Ȓl�����͂���܂��� divided:" & divided & ",  divisor:" & divisor
    End If
    ret = "�v�Z���ʂ�" & result & "�ł�"
    
End Function



Private Function A(ByVal divided As Integer, ByVal divisor As Integer, ByRef result As Integer) As String
    On Error Resume Next

    A = "" ' �����l
    result = divided / divisor
    If Err.Number <> 0 Then
        A = "���[�`��A�F" & Err.Description
    End If
    On Error GoTo 0

End Function

