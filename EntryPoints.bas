Attribute VB_Name = "EntryPoints"
'�K�{�@�錾���Ă��Ȃ��ϐ����R���p�C���G���[�����Ƃ���
Option Explicit

'�����W���[���ɂ́A�}�N���Ƃ��ČĂяo���\��Public Sub�̂݋L�ڂ���


Public Sub �����I���Z���ɂ��钆�g���p�[�X����()

    'java���ۂ��@ �N���X���ƕϐ����𓯂��ɂ��Ȃ��ق��������@�i �ϐ��������D�悳��āA�Ȃ����^���������ɂȂ��Ă��܂�)
    Dim log As LogUtil
    Set log = New LogUtil
    log.DebugMode = True
    
    log.info message:="----------------------------------"
    log.info message:="�I���Z���̎������m�F�p�p�[�X�����J�n"
    
    Dim cu As CollectionUtil
    Set cu = New CollectionUtil
    cu.LogObj = log

    '������
    Dim result As collection
    Set result = cu.SelectionToCollection(Application.selection)
    Set result = cu.Sanitize(result, "��", "=")
    Set result = cu.Sanitize(result, "�D", ".")
    Set result = cu.Sanitize(result, "�C", ",")
    Set result = cu.SplitBy(result, Space$(1))
    Set result = cu.SplitBy(result, ",")
    Set result = cu.SplitBy(result, "=")
    Call log.info("���ŏI���ʁ�")
    Call log.info(cu.toString(result))

    GoTo LogicEnd
ErrorHandler:
    Call MsgBox(Err.Description, vbCritical + vbOKOnly)

LogicEnd:
    Set result = Nothing
    Set cu = Nothing
    Set log = Nothing

End Sub


' vb6�G���[�R�[�f�B���O�{�ɂ���p�^�[��
' function�̖߂�l���G���[����ŌŒ肷��p�^�[��
' �֐��̖߂�����ׂāA������byref �Q�Ɠn���ɂ���K�v������A��肪�����ɂȂ�
' �����I�ɂ�C����̏����ɋ߂�
Public Sub �G���[�̖A�p�^�[��1()
    
    Dim ret As String
    ret = ""
    
    Dim errMessage As String
    
    '���O��Ԃ��쐬���邱�Ƃ��ł���i�Ƃ������A���W���[�����������������ɁA�ǂ��ɂ��邩��T���₷������
    '(Shift + F2�Œ�`�ɔ�ׂ΂����j
    '�������O��public�֐����쐬���Ă����v�Ȃ悤�ɂ���
    errMessage = Fromvb6ErrCoding.C(ret)
    
    Debug.Print errMessage
    Debug.Print ret


End Sub

Public Sub �G���[�̖A�p�^�[��2()
On Error GoTo ErrorHandler
    
    Dim result As Integer
    result = FromQuiita.C

    Debug.Print "�I�����܂���"
    GoTo LogicEnd
ErrorHandler:
    Debug.Print Err.Source & vbCrLf & Err.Description
LogicEnd:
    
End Sub
