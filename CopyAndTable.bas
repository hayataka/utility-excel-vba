Attribute VB_Name = "CopyAndTable"
Option Explicit

'�c�[���@�Q�Ɛݒ�@Microsoft Scripting.Runtime

Private Const NOTHING_POS As Integer = -1

Private Const SEPARATE_KEY As String = "&="

Public Sub �}�g���b�N�X����\�ɂ���()
    

    
    
    If MsgBox(prompt:="�p�����[�^�l�ƃe�X�g�P�[�X�́Z�܂ł�I�����Ă�����s���Ă��������B", Title:="NO�ŃL�����Z���\", Buttons:=vbYesNo + vbInformation) = vbNo Then
        Exit Sub
    End If
    
    ' ��{�I�ȕ����I���Z�������擾����
    Dim info As ClsOriginInfoDto
    Set info = toInfo
   
   '�e�p�����[�^�ɐݒ肷��l�̏�񂪂������擾����
    info.Prm_Value_ColPos = paramValuePos
    If info.Prm_Value_ColPos = NOTHING_POS Then
        Exit Sub
    End If
    
    
    'testCaseID���񋓂���Ă���Z�����擾����
    Dim testCaseRanges As Range
    Set testCaseRanges = inputCases
    If testCaseRanges Is Nothing Then
        Exit Sub
    End If
    
    ' TestCase�ӏ��݂̂̑I�������擾����
    Dim caseTitleInfo As ClsOriginInfoDto
    Set caseTitleInfo = toInfo
    
     '��ƑΏۃV�[�g����������
    Dim wkSt As Excel.Worksheet
    Set wkSt = copyWkIn
    
    'wk���V�[�g�́A�p�����[�^���ӏ��ŋ󔒃Z���𖄂߂�
    Call putParamName(info, wkSt)
    ' test�P�[�X�ӏ��̃}�g���b�N�X�̒l�������ɁA�J�������{�l��ݒ肷��
    Call putCaseValue(wkSt, info, testCaseRanges)
    
    Dim items As Dictionary
    Set items = caseDictionary(wkSt, info, testCaseRanges)
    
'    Dim k As Variant
'    For Each k In items
'        Debug.Print k
'        Dim oneCase As Dictionary
'        Set oneCase = items.Item(k)
'        Dim inK As Variant
'        For Each inK In oneCase
'            Debug.Print oneCase.Item(inK)
'        Next inK
'    Next k
    
    Dim outSt As Excel.Worksheet
    Set outSt = Worksheets.Add
    outSt.Name = "out_" & Hour(Now) & Minute(Now) & Second(Now)
    
    
End Sub

Private Function toInfo() As ClsOriginInfoDto

    Dim LoopArea As Range
    Set LoopArea = Selection
    
    
    Dim inf As ClsOriginInfoDto
    Set inf = New ClsOriginInfoDto

    inf.RowCnt = LoopArea.Row
    inf.ColCnt = LoopArea.Column


    inf.StartRowPos = LoopArea.Cells(1).Row
    inf.StartColumnPos = LoopArea.Cells(1).Column
    inf.Prm_Name_ColPos = inf.StartColumnPos
    inf.Max_RowPos = LoopArea.Cells(LoopArea.Count).Row
    inf.Max_ColumnPos = LoopArea.Cells(LoopArea.Count).Column
        
    Set toInfo = inf

End Function

Private Function copyWkIn() As Excel.Worksheet

    Dim originSt As Excel.Worksheet
    Set originSt = ActiveSheet
   
    Dim wkSt As Excel.Worksheet
    Call Sheets(originSt.Name).Copy(after:=ActiveWorkbook.Worksheets(originSt.Name))
    Set wkSt = ActiveSheet
    Dim newNm As String
    newNm = Hour(Now) & Minute(Now) & Second(Now)
    wkSt.Name = "in_" & newNm
    
    Set copyWkIn = wkSt
End Function

Private Sub putParamName(ByRef inf As ClsOriginInfoDto, ByRef st As Excel.Worksheet)

    Dim i As Integer
    Dim befCellVal As String
    For i = inf.StartRowPos To inf.Max_RowPos
        Dim nowCellVal As String
        nowCellVal = st.Cells(i, inf.StartColumnPos).value
        If nowCellVal <> "" Then
            befCellVal = nowCellVal
        Else
            st.Cells(i, inf.StartColumnPos).value = befCellVal
        End If
        Debug.Print nowCellVal
    Next i
End Sub

Private Function inputCases() As Range
    Const TYPE_IS_RANGE As Integer = 8
    Dim buf As Range
    ' �L�����Z������ƁA��O����
    On Error Resume Next
        Set buf = Application.InputBox(prompt:="�e�X�g�P�[�XID���L�ڂ���Ă���Z���𕡐��I�����Ă��������B", Type:=TYPE_IS_RANGE, Title:="�L�����Z���ŏ������I�����܂�")
    On Error GoTo 0
    If buf Is Nothing Then
        Set inputCases = Nothing
    Else
        Set inputCases = buf
    End If

End Function

Private Function paramValuePos() As Integer

    Const TYPE_IS_RANGE As Integer = 8
    Dim buf As Range
    ' �L�����Z������ƁA��O����
    On Error Resume Next
        Set buf = Application.InputBox(prompt:="parameter�̒l������Z����1�I�����Ă��������B(����Ƃ��)", Type:=TYPE_IS_RANGE, Title:="�L�����Z���ŏ������I�����܂�")
    On Error GoTo 0
    If buf Is Nothing Then
        paramValuePos = NOTHING_POS
    Else
        paramValuePos = buf.Column
    End If

End Function


Private Sub putCaseValue(ByRef st As Excel.Worksheet, ByRef info As ClsOriginInfoDto, ByRef caseTitleInfo As Range)

    'testcase�̃}�g���b�N�X�́Z�����Ă���ʒu�ɁA�p�����[�^��&=�p�����[�^�l��ݒ肷�邱�Ƃ�ړI�Ƃ���
    Dim startCol As Integer
    startCol = caseTitleInfo.Column
    Dim endCol As Integer
    endCol = info.Max_ColumnPos
    
    Dim col As Integer
    Dim r As Integer
    For col = startCol To endCol
        Dim startRow As Integer
        Dim endRow As Integer
        startRow = info.StartRowPos
        endRow = info.Max_RowPos
        For r = startRow To endRow
            Dim v As String
            v = st.Cells(r, col).value
            If v <> Space$(0) Then
                Dim pName As String
                pName = st.Cells(r, info.Prm_Name_ColPos).value
                Dim pVal As String
                pVal = st.Cells(r, info.Prm_Value_ColPos).value
                st.Cells(r, col).value = pName & SEPARATE_KEY & pVal
            End If
        Next r
    Next col

End Sub

Private Function caseDictionary(ByRef st As Excel.Worksheet, ByRef info As ClsOriginInfoDto, ByRef testCaseRanges As Range) As Dictionary

    Dim result As Dictionary
    Set result = New Dictionary
    
    Dim hasError As Boolean
    hasError = False

    Dim startCol As Integer
    startCol = testCaseRanges.Column
    Dim caseTitleRow As Integer
    caseTitleRow = testCaseRanges.Row
    Dim endCol As Integer
    endCol = info.Max_ColumnPos
    
    Dim col As Integer
    Dim r As Integer
    For col = startCol To endCol
        Dim startRow As Integer
        Dim endRow As Integer
        startRow = info.StartRowPos
        endRow = info.Max_RowPos
        
        Dim oneCase As Dictionary
        Set oneCase = New Dictionary
        For r = startRow To endRow
            Dim v As String
            v = st.Cells(r, col).value
            If v <> Space$(0) Then
                Dim tmp As Variant
                tmp = Split(v, delimiter:=SEPARATE_KEY)
                Dim pName As String
                pName = tmp(0)
                Dim pVal As String
                pVal = tmp(1)
                On Error Resume Next
                pVal = Split(pVal, "(")(0)
                If Err.Number <> 0 Then
                    If hasError = False Then
                        MsgBox v & "��()�ň͂��Ă��܂���"
                        Debug.Print v & "�̂���s:" & r
                        hasError = True
                    Else
                        Debug.Print v & "�̂���s:" & r
                    End If
                End If
                On Error GoTo 0
                oneCase.Add Key:=pName, Item:=pVal
            End If
        Next r
        result.Add Key:=st.Cells(caseTitleRow, col).value, Item:=oneCase
    Next col


    Set caseDictionary = result

End Function

