Attribute VB_Name = "basInputCheck"
' @(h) basInputCheck.bas  ver 1.00 ( 2000/08/29 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   �v���W�F�N�g��  : L&F��ۼު��
'   ���W���[����    : basInputCheck
'   �t�@�C����      : basInputCheck.bas
'   Version        : 1.00
'   �@�\����       �F ���������Ɋւ��鋤�ʊ֐�
'   �쐬��         �F T.Fukutani
'   �쐬��         �F 2000/12/01
'   �C�������@�@�@  �F 2001/05/15  GF_Com_KeyPress�ɉp���������p�啶����ǉ�
'   �@�@�@�@�@�@�@     2001/12/11  GF_Com_KeyPress14,15��ǉ� <T.Matsui>
'   �@�@�@�@�@�@�@     2002/01/08  GF_Com_KeyPress���g�p����GF_Com_CheckString���쐬 <T.Matsui>
'   �@�@�@�@�@�@�@     2002/01/10  GF_ChangeQuateSing���쐬 <N.Kigaku>
'   �@�@�@�@�@�@�@     2002/01/17  GF_ReplaceAmper���쐬 <N.Kigaku>
'   �@�@�@�@�@�@�@     2002/04/09  GF_Com_KeyPress16��ǉ� <N.Kigaku>
'   �@�@�@�@�@�@�@     2002/07/11  GS_Com_TxtGotFocus��'chk'��ǉ� <N.Kigaku>
'                      2002/08/27  GF_MinusCheck��ǉ� <N.Kigaku>
'                      2002/10/23  GF_FileNameRestrinction��ǉ� <N.Kigaku>
'                      2005/01/11  GF_DateConv�ǉ� <N.Kigaku>
'                      2005/12/28  GF_THJCMBXMR_CHK�ǉ� <N.Kigaku>
'                      2006/01/06  GF_CheckNumber2,GF_ChkDeci�ǉ� <N.Kigaku>
'                      2006/01/19  GF_Com_KeyPress�ɋ��������"17"�����������׸ނ�ǉ�,
'                                  GF_OptFormatChk,�萔[��߼�݌���,���޵�߼�݌���]�ǉ� <N.Kigaku>
'                      2006/01/31  GF_DateConv�ɑS�p�����ǉ� <N.Kigaku>
'                      2006/07/07  GF_UndoAmper��ǉ� <N.Kigaku>
'                      2006/12/05  �׸�8.1.7 Nocache�Ή� �������AReadOnly����Nocache�ɕύX <N.Kigaku >
'                      2006/12/08  ���s���������֐�[GF_CheckLinefeed]�ǉ� <N.Kigaku>
'                      2006/12/11  GF_CheckEngNumMark�ǉ� <N.Kigaku>
'                      2008/02/21  GF_CharPermitChek�̋��������12��ʲ�������C��
'                      2008/05/27  GF_ChangeQuateDouble�֐��ǉ� <N.Kigaku>
'                      2009/03/16  GF_Com_KeyPress�ɋ��������"18"��ǉ� <N.Kigaku>
'                      2011/05/20  GF_CharPermitChek�ɋ��������"15","16"��ǉ� <N.Kigaku>
'                      2016/11/14  GF_Com_KeyPress�ɐ��亰�ޗL������݂�ǉ��AGF_Com_CheckString�����������׸ނ�ǉ�
'                      2017/01/11  M.Tanaka K545 CS�v���Z�X���P GF_CheckStartToEnd�ǉ�
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ���錾
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
'  �萔�錾
'------------------------------------------------------------------------------
Private Const mlngOptLength = 4         '��߼�݌���
Private Const mlngSizeOptLength = 8     '���޵�߼�݌���

' 2017/01/11 �� M.Tanaka K545 CS�v���Z�X���P  ADD
Public Enum CSTE_ChkKbn '�`�F�b�N�敪
    CSTE_Year = 1           '�N
    CSTE_Month = 2        '��
    CSTE_Date = 3           '��
End Enum
' 2017/01/11 �� M.Tanaka K545 CS�v���Z�X���P  ADD

Public Function GF_Com_CheckString(intPatan As Integer, _
                                   strCharCode As String, _
                          Optional bolAsterisk As Boolean = False, _
                          Optional bolEnterChkFlg As Boolean = True) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ������R�[�h�`�F�b�N
' �@�\�@ : �����ꂽ�R�[�h�ȊO��������Ɋ܂܂�Ă��Ȃ����`�F�b�N���s��
' �����@ : intPatan As Integer        ''���������
' �@�@�@                                (GF_Com_KeyPress�ɏ���)
' �@�@�@   strCharCode As String      ''������(�`�F�b�NOK�̏ꍇ�A�啶�����ȂǕϊ���̒l��Ԃ�)
' �@�@�@   bolAsterisk As Boolean     ''���ؽ������׸�(�ȗ���False=�s��)
'          bolEnterChkFlg As Boolean  ''���������׸�(�ȗ���True=�L)
' �߂�l : Boolean                    ''True����`�F�b�NOK     False����`�F�b�NNG
' ���l�@ : �����GF_Com_KeyPress�ɏ������Ă��Ȃ��ƈӖ����Ȃ�(�����Ȃǂ�)
'------------------------------------------------------------------------------
    Dim l                           As Long
    Dim intRet                      As Integer
    Dim intAscii                    As Integer
    Dim sString                     As String
    
    GF_Com_CheckString = True
    
    '1����1�����A������̒������`�F�b�N
    sString = ""
    For l = 1 To Len(strCharCode)
        
        '�A�X�L�[�R�[�h�ɕϊ�
        intAscii = Asc(Mid$(strCharCode, l, 1))
        
        '���̓L�[�R�[�h�`�F�b�N
'2016/11/14 LF1667_TIK�S�����[���A�h���X���̌��g�� START <<<
'        intRet = GF_Com_KeyPress(intPatan, intAscii, bolAsterisk)
        intRet = GF_Com_KeyPress(intPatan, intAscii, bolAsterisk, bolEnterChkFlg)
'2016/11/14 LF1667_TIK�S�����[���A�h���X���̌��g�� END >>>
        If (intRet = 0) Then
            '�����ȕ��������������ꍇ
            GF_Com_CheckString = False
            Exit Function
        End If
        
        '�����ɖ߂��Ċi�[
        sString = sString & Chr$(intAscii)
    Next l
    
    '�啶�����ȂǕϊ��㕶�����Ԃ�
    strCharCode = sString
    
End Function

Public Function GF_Com_KeyPress(intPatan As Integer, _
                                intKeyAscii As Integer, _
                       Optional bolAsterisk As Boolean = False, _
                       Optional bolEnterChkFlg As Boolean = True, _
                       Optional intCtrlPatan As Integer = 0) As Integer
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���̓L�[�R�[�h�`�F�b�N
' �@�\   : ���͋����ꂽ�L�[�R�[�h�ȊO�̂��̂��͂���
' ����   : intPatan As Integer     ''���������
'                0  - Keypress Code Non Check
'                1  - ����  Code Non Check  "0,1,2,�`9"
'                2  - �����{��ص�� Code Non Check   "0,1,2,�`9,.,"
'                3  - �����{��ص�ށ{ϲŽ Code Non Check   "0,1,2,�`9,.,-"
'                4  - �����{�p�� Code Non Check   "0,1,2,�`9,A�`Z"
'                5  - �J�i����
'                6  - �����{ϲŽ Code Non Check   "0,1,2,�`9,-"
'                7  - �����{�p�� Code Non Check   "0,1,2,�`9,A�`Z,a�`z,"
'                8  - '!' �` '}' �܂�OK    (���ނ��� 33 �` 125�܂�)
'                9  - �p�����{��׽�{ϲŽ�{"*" Code Non Check   "0,1,2,�`9,A�`Z,a�`z,+,-,*"
'                10 - �p�啶�� Code Non Check   "A�`Z"
'                11 - �����{ʲ�݁{�ׯ�� Code Non Check   "0,1,2,�`9,-,/"
'                12 - �����{ʲ�݁{��� Code Non Check   "0,1,2,�`9,-,(,)"
'                13 - �p������ �� �p�啶��      "a�`z""
'                14 - �����{�p���{ʲ�� Code Non Check   "0,1,2,�`9,A�`Z","-"
'                15 - �����{�p���{���ݸ Code Non Check   "0,1,2,�`9,A�`Z"," "
'                16 - �����{���ݸ Code Non Check   "0,1,2,�`9," "
'                17 - ASCII����(0�`127)�̐��䕶��(��-�,ײ�̨-��ޏ���)�ȊO�S�ċ���
'                18 - ASCII����(0�`127)+�g��ASCII����(128�`255)�̐��䕶���ȊO�S�ċ���
'          intKeyAscii As Integer      ''��������
'          bolAsterisk As Boolean      ''���ؽ������׸�(�ȗ���False=�s��)
'          bolEnterChkFlg              ''���������׸�(�ȗ���True=�L)
'          intCtrlPatan As Integer     ''���亰�ޗL�������
'                 0 - ��������
'                 1 - Ctrl+C,Ctrl+V
'                 2 - Ctrl+C,Ctrl+V,Ctrl+X
'                 3 - Ctrl+C,Ctrl+V,Ctrl+X,Ctrl+Z
' �߂�l : Integer          ''0����������ނ͖��� �@ 1����������ނ͗L��
' ���l   :
'------------------------------------------------------------------------------
    GF_Com_KeyPress = 0    '�����ނ͖���

    ''���(9)  �ޯ����-�(8) �- ����   ===>���䖳��
    If (intKeyAscii = 9) Then
        Exit Function
    ElseIf (intKeyAscii = 8) Then
        GF_Com_KeyPress = 1   ''�����ނ͗L��
        Exit Function
    End If
    ''��-�(13) ײ�̨-���(10) �- ����   ===>���䖳��
    If (bolEnterChkFlg = True) And (intKeyAscii = 13 Or intKeyAscii = 10) Then
        intKeyAscii = 0     '�������Ȃ���BEEP�����邽��
        Exit Function
    End If
    
    '���ؽ����͋����A*(42)������
    If bolAsterisk = True And intKeyAscii = 42 Then
        GF_Com_KeyPress = 1   ''�����ނ͗L��
        Exit Function
    End If

'2016/11/14 LF1667_TIK�S�����[���A�h���X���̌��g�� START <<<
    '����R�[�h�ňȉ��̂��̂͗L���Ƃ���
    Select Case intCtrlPatan
        Case 0
        Case 1
            'Ctrl+C(3), Ctrl+V(22)
            If (intKeyAscii = 3) Or (intKeyAscii = 22) Then
                GF_Com_KeyPress = 1   ''�����ނ͗L��
                Exit Function
            End If
        
        Case 2
            'Ctrl+C(3), Ctrl+V(22), Ctrl+X(24)
            If (intKeyAscii = 3) Or (intKeyAscii = 22) Or (intKeyAscii = 24) Then
                GF_Com_KeyPress = 1   ''�����ނ͗L��
                Exit Function
            End If
        
        Case 3
            'Ctrl+C(3), Ctrl+V(22), Ctrl+X(24), Ctrl+Z(26)
            If (intKeyAscii = 3) Or (intKeyAscii = 22) Or (intKeyAscii = 24) Or (intKeyAscii = 26) Then
                GF_Com_KeyPress = 1   ''�����ނ͗L��
                Exit Function
            End If
    End Select
'2016/11/14 LF1667_TIK�S�����[���A�h���X���̌��g�� END >>>

    '��������� ����
    Select Case intPatan
        Case 0          '' Check Non
        
        Case 1          '' 0-9
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                intKeyAscii = 0
            End If
            
        Case 2          '' 0-9 or .(46)
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                If intKeyAscii <> 46 Then
                   intKeyAscii = 0
                End If
            End If

        Case 3          '' 0-9 or -(45) or .(46)
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                If intKeyAscii <> 45 And intKeyAscii <> 46 Then
                   intKeyAscii = 0
                End If
            End If
            
        Case 4          '' 0-9 or A-Z
            Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''�������ˑ啶���ɕϊ�
            End Select
                    
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
               ((intKeyAscii < 65) Or (intKeyAscii > 90)) Then
               intKeyAscii = 0
            End If
                
        Case 5          '' ������ Non Check
            
        Case 6          '' 0-9 or -(45)
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And (intKeyAscii <> 45) Then
                   intKeyAscii = 0
            End If
            
        Case 7          '' 0-9 or A-Z or a-z
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
               ((intKeyAscii < 65) Or (intKeyAscii > 90)) And _
               ((intKeyAscii < 97) Or (intKeyAscii > 122)) Then
               intKeyAscii = 0
            End If
            
        Case 8          ''  "!" �` "}"
            If intKeyAscii < 33 Or intKeyAscii > 125 Then
               intKeyAscii = 0
            End If
            
        Case 9          '' 0-9 or A-Z or a-z or +(43) or -(45) or *(42)
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
                    ((intKeyAscii < 65) Or (intKeyAscii > 90)) And _
                    ((intKeyAscii < 97) Or (intKeyAscii > 122)) Then
                If intKeyAscii = 43 Or intKeyAscii = 45 Or intKeyAscii = 42 Then
                
                Else
                    intKeyAscii = 0
                End If
            End If
        Case 10          '' A-Z
            Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''�������ˑ啶���ɕϊ�
            End Select
                    
            If intKeyAscii < 65 Or intKeyAscii > 90 Then
               intKeyAscii = 0
            End If
        Case 11          '' 0-9 or -(45) or /(47)
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                If intKeyAscii <> 45 And intKeyAscii <> 47 Then
                   intKeyAscii = 0
                End If
            End If
        
        Case 12          '' 0-9 or -(45) or ((40) or )(41)
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                If intKeyAscii <> 45 And intKeyAscii <> 40 And intKeyAscii <> 41 Then
                   intKeyAscii = 0
                End If
            End If
        Case 13
           Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''�������ˑ啶���ɕϊ�
            End Select
            
        Case 14          '' 0-9 or A-Z or "-"
            
            'ʲ̫ݓ��͋���
            If (intKeyAscii = 45) Then
                GF_Com_KeyPress = 1   ''�����ނ͗L��
                Exit Function
            End If
            
            Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''�������ˑ啶���ɕϊ�
            End Select
                    
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
               ((intKeyAscii < 65) Or (intKeyAscii > 90)) Then
               intKeyAscii = 0
            End If
            
        Case 15          '' 0-9 or A-Z or " "
            
            '���ݸ���͋���
            If (intKeyAscii = 32) Then
                GF_Com_KeyPress = 1   ''�����ނ͗L��
                Exit Function
            End If
            
            Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''�������ˑ啶���ɕϊ�
            End Select
                    
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
               ((intKeyAscii < 65) Or (intKeyAscii > 90)) Then
               intKeyAscii = 0
            End If
            
        Case 16          '' 0-9 or A-Z or " "
            
            '���ݸ���͋���
            If (intKeyAscii = 32) Then
                GF_Com_KeyPress = 1   ''�����ނ͗L��
                Exit Function
            End If
            
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) Then
               intKeyAscii = 0
            End If
            
        Case 17
            
            'ASCII����(0�`127)�̐��䕶��(�ꕔ����)�ȊO�S�ċ���
            ''�����䕶�� (��-݁Aײ�̨-���)
            If (bolEnterChkFlg = False) And (intKeyAscii = 13 Or intKeyAscii = 10) Then
                GF_Com_KeyPress = 1   ''�����ނ͗L��
                Exit Function
            End If
            If (intKeyAscii < 32) Or (intKeyAscii > 126) Then
                intKeyAscii = 0
            End If

'2009/03/16 Added by N.Kigaku Start --------------------------------------------------------
        Case 18
            'ASCII����(0�`127)+�g��ASCII����(128�`255)�̐��䕶��(�ꕔ����)�ȊO�S�ċ���
            ''�����䕶�� (��-݁Aײ�̨-���)
            If (bolEnterChkFlg = False) And (intKeyAscii = 13 Or intKeyAscii = 10) Then
                GF_Com_KeyPress = 1   ''�����ނ͗L��
                Exit Function
            End If
            If ((intKeyAscii < 32) Or (intKeyAscii > 126)) And _
               ((intKeyAscii < 160) Or (intKeyAscii > 223)) Then
                intKeyAscii = 0
            End If
'2009/03/16 End ----------------------------------------------------------------------------

    End Select

    If intKeyAscii = 0 Then
        GF_Com_KeyPress = 0    '�����ނ͖���
    Else
        GF_Com_KeyPress = 1    '�����ނ͗L��
    End If
    
End Function
Public Function GF_Com_CutNumber(strIn_txt As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���p�����̐؂�o��
' �@�\   : �����񒆂̔��p�����݂̂�؂�o��
' ����   : strIn_txt As String  '' ������
' �߂�l : String          �؂�o���ꂽ���p����
' ���l   :
'------------------------------------------------------------------------------
    Dim intLoop_c As Integer    'ٰ�߶���
    Dim strChk_chr As String    '���������p
    Dim strOut_chr As String    '������쐬�p
    '1�����ڂ��珇���������A���p�����̏ꍇ�͐؂�o��
    For intLoop_c = 1 To Len(strIn_txt)
        strChk_chr = Mid(strIn_txt, intLoop_c, 1)
        Select Case strChk_chr
            Case "0" To "9"
                strOut_chr = strOut_chr & strChk_chr
        End Select
    Next intLoop_c
    GF_Com_CutNumber = strOut_chr
End Function

Public Function GF_Com_CutAlfNum(strIn_txt As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���p�p�����̐؂�o��
' �@�\   : �����񒆂̔��p�p�����݂̂�؂�o��
' ����   : strIn_txt As String  '' ������
' �߂�l : String          �؂�o���ꂽ���p�p����
' ���l   :
'------------------------------------------------------------------------------
    Dim intLoop_c As Integer    'ٰ�߶���
    Dim strChk_chr As String    '���������p
    Dim strOut_chr As String    '������쐬�p
    '1�����ڂ��珇���������A���p�p�����̏ꍇ�͐؂�o��
    For intLoop_c = 1 To Len(strIn_txt)
        strChk_chr = Mid(strIn_txt, intLoop_c, 1)
        Select Case strChk_chr
            Case "0" To "9", "A" To "Z", "a" To "z"
                strOut_chr = strOut_chr & strChk_chr
        End Select
    Next intLoop_c
    GF_Com_CutAlfNum = strOut_chr
End Function

Public Sub GS_TextSelect(txtControl As TextBox)
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �e�L�X�g�S�I��
' �@�\   : ÷��BOX�̃e�L�X�g��S�I������
' ����   : txtControl As TextBox   '÷��BOX
' ���l   :
'------------------------------------------------------------------------------
    
    ''̫������擾������S�I����Ԃɂ���
    txtControl.SelStart = 0
    txtControl.SelLength = Len(txtControl)
    
End Sub

Public Function GF_CheckHostStr(strTemp As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : νĎ�M�\�����ϊ�
' �@�\   : νĂɂēo�^�s�\������o�^�\�����ɕϊ�����B
' ����   : strTemp   As String      ''�Ώە�����
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    Dim nCount    As Integer
    Dim intLength As Integer
    
    GF_CheckHostStr = True
    
    intLength = Len(strTemp)
    
    For nCount = 1 To intLength

        Select Case Mid(strTemp, nCount, 1)
        Case "�" To "�"
        Case "�"
        Case "�"
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "�" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "�" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "�" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "�" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "�" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "�" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "�" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "�" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "�" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "�"
            strTemp = Left(strTemp, nCount - 1) & "-" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case Chr(34)    '����ٸ��ð���
            strTemp = Left(strTemp, nCount - 1) & Chr(39) & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "\"
        Case "�"
        Case Chr(32) To Chr(90)
        Case Else
            strTemp = Left(strTemp, nCount - 1) & "?" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        End Select
     
    Next nCount
    
End Function

Public Function GF_ConvertWide(strChar As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ������̑S�p�ϊ�
' �@�\   : �����񒆂̔��p������S�p�ɕϊ�����
' ����   : strChar As String   '������
' �߂�l : String (�ϊ���̕�����)
' ���l   :
'------------------------------------------------------------------------------
    GF_ConvertWide = strConv(strChar, vbWide)
    
End Function

Public Function GF_ConvertNarrow(strChar As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ������̔��p�ϊ�
' �@�\   : �����񒆂̑S�p�����𔼊p�ɕϊ�����
' ����   : strChar As String   '������
' �߂�l : String (�ϊ���̕�����)
' ���l   :
'------------------------------------------------------------------------------
    GF_ConvertNarrow = strConv(strChar, vbNarrow)
    
End Function

Public Function GF_CutWide(strIn_txt As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �S�p�����̐؂�o��
' �@�\   : �����񒆂̑S�p�����݂̂�؂�o��
' ����   : strIn_txt As String  '' ������
' �߂�l : String          �؂�o���ꂽ�S�p������
' ���l   :
'------------------------------------------------------------------------------
    Dim intLoop_c As Integer    'ٰ�߶���
    Dim strChk_chr As String    '���������p
    Dim strOut_chr As String    '������쐬�p
    strOut_chr = ""
    '1�����ڂ��珇���������A���p�p�����̏ꍇ�͐؂�o��
    For intLoop_c = 1 To Len(strIn_txt)
        strChk_chr = Mid(strIn_txt, intLoop_c, 1)
        If LenB(strConv(strChk_chr, vbFromUnicode)) = 2 Then
            strOut_chr = strOut_chr & strChk_chr
        End If
    Next intLoop_c
    GF_CutWide = strOut_chr
End Function

Public Function GF_CutHalf(strIn_txt As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���p�ŉp�����̐؂�o��
' �@�\   : �����񒆂̔��p�ŉp�����݂̂�؂�o��
' ����   : strIn_txt As String  '' ������
' �߂�l : String          �؂�o���ꂽ���p�ŉp����
' ���l   :
'------------------------------------------------------------------------------
    Dim intLoop_c As Integer    'ٰ�߶���
    Dim strChk_chr As String    '���������p
    Dim strOut_chr As String    '������쐬�p
    Dim intAsc     As Integer
    strOut_chr = ""
    '1�����ڂ��珇���������A���p�p�����̏ꍇ�͐؂�o��
    For intLoop_c = 1 To Len(strIn_txt)
        strChk_chr = Mid(strIn_txt, intLoop_c, 1)
        intAsc = Asc(strChk_chr)
        If ((intAsc >= 48) And (intAsc <= 57)) Or _
           ((intAsc >= 65) And (intAsc <= 90)) Or _
           ((intAsc >= 97) And (intAsc <= 122)) Or _
           ((intAsc >= 166) And (intAsc <= 223)) Then
            strOut_chr = strOut_chr & strChk_chr
        End If
    Next intLoop_c
    GF_CutHalf = strOut_chr
End Function

Public Sub GS_DecimalPointCheck(intKeyAscii As Integer, strString As String)
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �����_���͐���
' �@�\   : ����'�_'�̓��͐���
' ����   : intKeyAscii As Integer   ''��������
'          strString   As String    ''�ҏW���̕�����
' ���l   :
'------------------------------------------------------------------------------
    
    If intKeyAscii = Asc(".") Then
        ''1�����ڂ�"."��2�߂̏����_�Ȃ�͂���
        If strString = "" Or InStr(strString, ".") > 0 Then intKeyAscii = 0
    End If
    
End Sub

Public Function GF_MinusCheck(strString As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �}�C�i�X�݂̂̓��̓`�F�b�N
' �@�\   : �}�C�i�X�݂̂̓��̓`�F�b�N
' ����   :  strString   As String    ''������
' �߂�l : True = �}�C�i�X�ȊO / False = �}�C�i�X�̂�
' ���l   :
'------------------------------------------------------------------------------
    GF_MinusCheck = False
    
    If Trim(strString) <> "-" Then
        GF_MinusCheck = True
    End If
    
End Function

Public Function GF_StrCheck(strString As String, strConvert As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���p��������
' �@�\   : ������𔼊p�����ϊ����A�S�p�������c���Ă���δװ��Ԃ�
' ����   :
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    
    Dim intStrLength    As Integer  '������̒���(������)
    Dim intCount        As Integer  'ٰ�ߕϐ�
    
    On Error GoTo Err_GF_StrCheck
    
    GF_StrCheck = False
    
    strConvert = ""
    
    ''���p�ɕϊ��ł��镶���͕ϊ�����
    strConvert = strConv(strString, vbNarrow)
    '�������擾
    intStrLength = Len(strConvert)
    '�ꕶ��������������
    For intCount = 1 To intStrLength
        '1�޲ĈȊO������δװ�Ƃ��ĕԂ�
        If LenB(strConv(Mid(strConvert, intCount, 1), vbFromUnicode)) _
                                                    <> 1 Then Exit Function
    Next
        
    GF_StrCheck = True
    
    Exit Function
    
Err_GF_StrCheck:
    
    Call GS_ErrorHandler("GF_StrCheck", "")
    
End Function

Public Function GF_HalfOrFullSizeCheck(ByVal strString As String, ByVal intCheck As Integer) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���p�S�p�����`�F�b�N
' �@�\   : �����񂪔��p(�܂��͑S�p)�����`�F�b�N����
' ����   : strString As String  ������
'          intCheck As Integer  �t���O    1�F���p�`�F�b�N�A2�F�S�p
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    
    Dim intStrLength    As Integer  '������̒���(������)
    Dim intCount        As Integer  'ٰ�ߕϐ�
    
    On Error GoTo Err_GF_HalfOrFullSizeCheck
    
    GF_HalfOrFullSizeCheck = False
    
    '�������擾
    intStrLength = Len(strString)
    '�ꕶ��������������
    For intCount = 1 To intStrLength
        '1�޲ĈȊO������δװ�Ƃ��ĕԂ�
        If LenB(strConv(Mid(strString, intCount, 1), vbFromUnicode)) _
                                                    <> intCheck Then Exit Function
    Next
        
    GF_HalfOrFullSizeCheck = True
    
    Exit Function
    
Err_GF_HalfOrFullSizeCheck:
    
    Call GS_ErrorHandler("GF_HalfOrFullSizeCheck", "")
    
End Function

Public Function GF_Com_KeyPressNum(KeyAscii As Integer, crtControl As Control, strTxtData As String, intPatan As Integer, intMaxlen As Integer, Optional dblMaxVal As Double = 9999999999999#) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���l�i���z�E�����j�̃`�F�b�N�֐�
' �@�\   : �L���͈͓��ł���� strTxtData �ɐݒ�l��ۑ�����
' ����   : KeyAscii As Integer  13(���ݺ��ށj ���w�肷��Ǝ��̃t�B�[���h�Ƀt�H�[�J�X���ړ�����
'                               0 ���w�肷��ƃt�H�[�J�X�̈ړ��Ȃ��Ń`�F�b�N���s����
'                               13(���ݺ��ށj�y�� 0 �̏ꍇ�A�G���[�������͓��͈�Ƀt�H�[�J�X�����킹��
'          crtControl As Control  ���۰�
'          strTxtData  As String   ÷��
'          intPatan As Integer
'               0 - ���z (#,##9)        ������ intMaxlen �Ɉˑ�����  �L���͈͂̉����l �w0 �ȏ�x�@����l�� dblMaxVal��
'               1 - ���z (#,##9)        ������ intMaxlen �Ɉˑ�����  �L���͈͂̉����l �w1 �ȏ�x�@����l�� dblMaxVal��
'               2 - ���� (#9.99999999)  ������ �Q���D�W���̌Œ�      �L���͈͂� �w0 ���� 100 �����x�̌Œ�
'               3 - ���z (#,##9.99)     ������ intMaxlen �Ɉˑ�����  �L���͈͂̉����l �w0    �ȏ�x�@����l�� dblMaxVal��
'               4 - ���z (#,##9.99)     ������ intMaxlen �Ɉˑ�����  �L���͈͂̉����l �w0.01 �ȏ�x�@����l�� dblMaxVal��
'               5 - ���� (#9.99999)     ������ �Q���D�T���̌Œ�      �L���͈͂� �w0 ���� 100 �����x�̌Œ�
'               6 - ���� (#9.999)       ������ �Q���D�R���̌Œ�      �L���͈͂� �w0 �ȏ�@ 100 �����x�̌Œ�
'               7 - ���� (#9.99999999)  ������ �Q���D�W���̌Œ�      �L���͈͂� �w0 �ȏ�@ 100 �����x�̌Œ�
'          intMaxlen As Integer ���l�̓��͉\�����i�����_�܂܂��j
'          dblMaxVal As Double  �L���͈͂̏���l�@��intPatan �� 2 , 5 , 6 �̏ꍇ������
' �߂�l : �G���[���b�Z�[�W
' ���l   :
'------------------------------------------------------------------------------
    Dim strMsg       As String
    Dim strCvtTxt    As String

    strMsg = ""
    If (KeyAscii = 13) Or (KeyAscii = 0) Then           '���ݺ��ނ��m�F�p���ނȂ�����
        If IsNumeric(crtControl.Text) = True Then
            If (intPatan = 0) Or (intPatan = 1) Then
                strCvtTxt = Format(crtControl.Text, "#,##0")
                crtControl.Text = strCvtTxt
            ElseIf (intPatan = 2) Or (intPatan = 7) Then
                strCvtTxt = GF_Com_DecCut((crtControl.Text), 8)
                strCvtTxt = Format(strCvtTxt, "#0.00000000")
            ElseIf (intPatan = 3) Or (intPatan = 4) Then
                strCvtTxt = Format(crtControl.Text, "#,##0.00")
            ElseIf intPatan = 6 Then
                strCvtTxt = GF_Com_DecCut((crtControl.Text), 3)
                strCvtTxt = Format(strCvtTxt, "#0.000")
            ElseIf intPatan = 5 Then
                strCvtTxt = GF_Com_DecCut((crtControl.Text), 5)
                strCvtTxt = Format(strCvtTxt, "#0.00000")
            End If
            
            If ((intPatan = 0) And (CDbl(strCvtTxt) < 0)) Or ((intPatan = 1) And (CDbl(strCvtTxt) < 1)) Or _
               ((intPatan = 3) And (CDbl(strCvtTxt) < 0)) Or ((intPatan = 4) And (CDbl(strCvtTxt) < 0.01)) Then
                strMsg = GF_GetMsg("WTG020")
                Call GS_Com_TxtGotFocus(crtControl)
            ElseIf ((intPatan = 0) Or (intPatan = 1) Or (intPatan = 3) Or (intPatan = 4)) And (CDbl(strCvtTxt) > dblMaxVal) Then
                strMsg = GF_GetMsg("WTG019")
                Call GS_Com_TxtGotFocus(crtControl)
            ElseIf (intPatan = 7) And ((CDbl(strCvtTxt) < 0) Or (CDbl(strCvtTxt) >= 100)) Then
                strMsg = GF_GetMsg("WTG021")
                Call GS_Com_TxtGotFocus(crtControl)
            ElseIf ((intPatan = 2) Or (intPatan = 6) Or (intPatan = 5)) And ((CDbl(strCvtTxt) <= 0) Or (CDbl(strCvtTxt) >= 100)) Then
                strMsg = GF_GetMsg("WTG021")
                Call GS_Com_TxtGotFocus(crtControl)
            Else
                crtControl.Text = strCvtTxt
                strTxtData = GF_Com_CutNumber(strCvtTxt)
                If (KeyAscii = 13) Then            '���ݺ��ނȂ�����
                    If (Len(GF_Com_CutNumber(crtControl.Text)) - Len(GF_Com_CutNumber(crtControl.SelText))) <= intMaxlen Then
                        Call GF_Com_KeyPress(0, KeyAscii)
                        Call GS_Com_NextCntl(crtControl)
                    ElseIf Not (KeyAscii = 9 Or KeyAscii = 8) Then
                        KeyAscii = 0
                    End If
                End If
            End If
        Else
            If (intPatan = 0) Or (intPatan = 1) Or (intPatan = 3) Or (intPatan = 4) Then
                strMsg = GF_GetMsg("ITH004") & GF_GetMsg("ITH001")
            ElseIf (intPatan = 2) Or (intPatan = 7) Or (intPatan = 6) Or (intPatan = 5) Then
                strMsg = GF_GetMsg("ITH003") & GF_GetMsg("ITH001")
            End If
            If crtControl.Enabled = True Then
                Call GS_Com_TxtGotFocus(crtControl)
            Else
'                MsgBox "�����ɖ��͂���܂��񂪒����˗������ĉ�����" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "���W���[���� : GF_Com_KeypressNum" & Chr(13) & Chr(10) & "�v���O�������� : " & G_APL_Job1 & G_APL_Job2 & Chr(13) & Chr(10) & "�R���g���[�� : " & crtControl.Name & Chr(13) & Chr(10) & "�⑫ : " & strMsg & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "OK�{�^�����N���b�N�������𑱍s���ĉ�����"
            End If
        End If
    Else
        If (Len(GF_Com_CutNumber(crtControl.Text)) - Len(GF_Com_CutNumber(crtControl.SelText))) <= intMaxlen Then
            Select Case intPatan
            '���l�̂�
            Case 0, 1
                Call GF_Com_KeyPress(1, KeyAscii)
            '���l�E�����_
            Case 2, 7, 3, 4, 6, 5
                Call GF_Com_KeyPress(2, KeyAscii)
                Call GS_DecimalPointCheck(KeyAscii, crtControl.Text)
            '���l�E�����_�E�}�C�i�X
            Case Else
                Call GF_Com_KeyPress(3, KeyAscii)
                Call GS_DecimalPointCheck(KeyAscii, crtControl.Text)
            End Select
        ElseIf Not (KeyAscii = 9 Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End If
    
    GF_Com_KeyPressNum = strMsg

End Function

Public Function GF_Com_DecCut(strTxt As String, intCut As Integer) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ��������؂�̂�
' �@�\   : �w�肳�ꂽ�ȍ~�̏�������؂�̂Ă�
' ����   : strTxt As String     ���͒l
'          intCut As Integer    �����扽�ʂ��w��(�����܂ŗL���j
' �߂�l : String   �Z�o�����l
' ���l   :
'------------------------------------------------------------------------------
    Dim intCheckChar As Integer
    Dim strOutChar As String
    
    strTxt = GF_Com_CnvSisu(strTxt)
    
    intCheckChar = InStr(1, strTxt, ".")
    
    If intCheckChar <> 0 Then
        strOutChar = Mid(strTxt, 1, (intCheckChar + intCut))
    Else
        strOutChar = strTxt
    End If
    
    GF_Com_DecCut = strOutChar
End Function

Public Function GF_Com_CnvSisu(strTxt As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �w���\���̐��l������ϊ�
' �@�\   : �w�肳�ꂽ�w���\���̐��l������i�����_�܂ށj���w���\�����g��Ȃ�������ɕϊ�����
' ����   : strTxt As String     ���͒l
' �߂�l : String   ���l������
' ���l   :
'------------------------------------------------------------------------------
    Dim intEPoint    As Integer  '�w���\�� 'E'�̈ʒu
    Dim intTenPoint  As Integer  '�����_�ʒu
    Dim intAddZerosu As Integer
    Dim strFugoChar  As String   '�w�����̕���
    Dim strFugoTxt   As String   '������擪�̕����i-�j
    Dim strHenTxt    As String
    Dim strZeroTxt   As String
    Dim intSisu      As Integer  '�w�����i�[
    Dim strTenMaeTxt As String   '�����_���O�̐���
    Dim strTenAtoTxt As String   '�����_����̐���
    Dim intKugirisu  As Integer
    
    '"E"�̈ʒu�Z�o
    intEPoint = InStr(1, strTxt, "E")
    '�����_"."�̈ʒu�Z�o
    intTenPoint = InStr(1, strTxt, ".")
    
    '�w�����\�����܂�ł��Ȃ����̂܂��͏����_���܂܂Ȃ����̂́A���͒l��߂�l�ɂ��Ĕ�����
    If (intEPoint = 0) Or (intTenPoint = 0) Then
        GF_Com_CnvSisu = strTxt
        Exit Function
    End If
    '�w�����i�[
    intSisu = CInt(Mid(strTxt, intEPoint + 2, 2))
    
    '�w�����ȑO�̕�����y�сA��̕�����𔲂��o��
    strHenTxt = Mid(strTxt, 1, intEPoint - 1)
    strTenMaeTxt = Mid(strHenTxt, 1, intTenPoint - 1)
    '�w�����ȑO�̕����񂪃}�C�i�X���܂�ł��邩
    If (InStr(1, strTenMaeTxt, "-")) <> 0 Then
        strTenMaeTxt = Mid(strTenMaeTxt, 2)
        strFugoTxt = "-"
    Else
        strFugoTxt = ""
    End If
    strTenAtoTxt = Mid(strHenTxt, intTenPoint + 1)
    
    '�w�����ȍ~�̕����ɂ���ď����U�蕪��
    strFugoChar = Mid(strTxt, intEPoint + 1, 1)
    '�}�C�i�X�̏ꍇ
    If strFugoChar = "-" Then
        '�����_�����炷�����̏ꍇ
        If intSisu < Len(strTenMaeTxt) Then
            intKugirisu = Len(strTenMaeTxt) - intSisu
            strHenTxt = Mid(strTenMaeTxt, 1, intKugirisu) & "." & Mid(strTenMaeTxt, intKugirisu + 1)
            strHenTxt = strHenTxt & strTenAtoTxt
        Else
        '0.0�d������ꍇ
            intAddZerosu = intSisu - 1
            strZeroTxt = "0."
            Do While intAddZerosu > 0
                strZeroTxt = strZeroTxt & "0"
                intAddZerosu = intAddZerosu - 1
            Loop
            strHenTxt = strZeroTxt & strTenMaeTxt & strTenAtoTxt
        End If
    '�v���X�̏ꍇ
    Else
        '�����_�����炷�����̏ꍇ
        If intSisu < Len(strTenAtoTxt) Then
            strHenTxt = Mid(strTenAtoTxt, 1, intSisu) & "." & Mid(strTenAtoTxt, intSisu + 1)
            strHenTxt = strTenMaeTxt & strHenTxt
        Else
        '00�d������ꍇ
            intAddZerosu = intSisu - Len(strTenAtoTxt)
            strZeroTxt = ""
            Do While intAddZerosu > 0
                strZeroTxt = strZeroTxt & "0"
                intAddZerosu = intAddZerosu - 1
            Loop
            strHenTxt = strTenMaeTxt & strTenAtoTxt & strZeroTxt
        End If
    End If
    
    GF_Com_CnvSisu = strFugoTxt & strHenTxt
End Function

Public Sub GS_Com_TxtGotFocus(crtControl As Control)
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �e�L�X�g�R���g���[���̓��͓��e��I����ԁi���]�\���j�ɂ���
' �@�\   :
' ����   : crtControl As Control  �ΏۂƂȂ�e�L�X�g
'                                 �R���{�{�b�N�X
'                                 �`�F�b�N�{�b�N�X
' ���l   :
'------------------------------------------------------------------------------
    Screen.ActiveForm.Enabled = True
    If (crtControl.Visible = True) And (crtControl.Enabled = True) Then
        Select Case LCase(Left(crtControl.Name, 3))
            Case "txt"
                crtControl.SetFocus
                crtControl.SelStart = 0
                crtControl.SelLength = Len(crtControl.Text)
            Case "cbo"
                crtControl.SetFocus
            Case "chk"
                crtControl.SetFocus
        End Select
    End If
End Sub

Public Function GF_Com_KeyPressDate(KeyAscii As Integer, cntControl As Control, strTxtDat As String, _
                                            intInCheck As Integer, Optional cntControl2 As Control = Nothing, Optional bolDispFlg As Integer = 0) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���t�̃`�F�b�N�֐�
' �@�\   : �Y������� strTxtDat �ɐݒ�l��ۑ�����
' ����   : KeyAscii As Integer  13(���ݺ��ށj ���w�肷��Ǝ��̃t�B�[���h�Ƀt�H�[�J�X���ړ�����
'                               0 ���w�肷��ƃt�H�[�J�X�̈ړ��Ȃ��Ń`�F�b�N���s����
'                               13(���ݺ��ށj�y�� 0 �̏ꍇ�A�G���[�������͓��͈�Ƀt�H�[�J�X�����킹��
'          crtControl As Control  ���۰�
'          strTxtData  As String  ÷��
'          intInCheck As Integer�@ 1 - �c�Ɠ�����
'                                  2 - ���e�N�Z������
'                                  4 - �{���c�Ɠ����z���Ă�������
'                                  8 - �{���c�Ɠ��ȍ~����
'                                 16 - �������t�i�{���L���j����
'                                 32 - �������t�i�{�������j
'                                 64 - �O���c�Ɠ�
'          crtControl2 As Control  ���۰�   �j���\���p���۰�(��c�Ɠ��͐Ԃ�����)
'          bolDispFlg  As Integer  �\���׸ށ@0:�N�����A1:�N���A2:�����A3:��
' �߂�l : String          �G���[���b�Z�[�W
' ���l   : �w�肳�ꂽ�`�F�b�N����(intInCheck)�Ɋ֌W�����c�Ɠ��`�F�b�N���j���̐F��ݒ肷��
'------------------------------------------------------------------------------
    Dim strDate     As String
    Dim strMsg      As String
    Dim strYoubiTbl() As Variant      '�j���e�[�u��
    
    
    strYoubiTbl = Array("(��)", "(��)", "(��)", "(��)", "(��)", "(��)", "(�y)")
    If Not (cntControl2 Is Nothing) Then
        cntControl2 = ""
        cntControl2.ForeColor = &H80000008              '�f�t�H���g�F�\��
    End If
    
    strMsg = ""
    strDate = GS_Com_TxtCvtYmd(strTxtDat)     'YYYY/MM/DD �ɕϊ�
    If (KeyAscii = 13) Or (KeyAscii = 0) Then           '���ݺ��ނ��m�F�p���ނȂ�����
        Select Case strDate
            Case ""
                strMsg = GF_GetMsg("ITH002") & GF_GetMsg("ITH001")   '������
                Call GS_Com_TxtGotFocus(cntControl)
            Case "Not Date"
                strMsg = GF_GetMsg("WTG022")       '�����ȓ��t
                Call GS_Com_TxtGotFocus(cntControl)
            Case "Less Than"
                strMsg = GF_GetMsg("WTG022")       '�W������
                Call GS_Com_TxtGotFocus(cntControl)
            Case Else
                '�ް��\��
                Select Case bolDispFlg
                Case 0
                '�N����
                    cntControl.Text = strDate
                Case 1
                '�N��
                    cntControl.Text = Format(strDate, "YYYY/MM")
                Case 2
                '����
                    cntControl.Text = Format(strDate, "MM/DD")
                Case 3
                '��
                    cntControl.Text = Format(strDate, "DD")
                End Select
                
                If ((intInCheck And 1) = 1) And (GF_Com_CheckOpen(strDate) = False) Then    '�c�Ɠ��`�F�b�N
                    strMsg = GF_GetMsg("ITG004")                   '�c�Ɠ��łȂ�
                    Call GS_Com_TxtGotFocus(cntControl)
'                ElseIf ((intInCheck And 2) = 2) And (CDbl(GF_Com_CutNumber(strDate)) < CDbl(G_rs_okkymd)) Then     '���e�N�Z���`�F�b�N
'                    strMsg = GF_GetMsg("B249")                   '���e�N�Z�����z���Ă���i������Ă���j
'                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 4) = 4) And (CDbl(GF_Com_CutNumber(strDate)) > CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then        '�ߋ����t�`�F�b�N�i�{���L���j
                    strMsg = GF_GetMsg("ITG005")                   '�{���c�Ɠ����z���Ă���
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 8) = 8) And (CDbl(GF_Com_CutNumber(strDate)) >= CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then     '�ߋ����t�`�F�b�N�i�{�������j
                    strMsg = GF_GetMsg("ITG006")                   '�{���c�Ɠ��ȍ~�ł�
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 16) = 16) And (CDbl(GF_Com_CutNumber(strDate)) < CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then     '�������t�`�F�b�N�i�{���L���j
                    strMsg = GF_GetMsg("ITG007")                   '�������t�i�{���L���j
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 32) = 32) And (CDbl(GF_Com_CutNumber(strDate)) <= CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then     '�������t�`�F�b�N�i�{�������j
                    strMsg = GF_GetMsg("ITG008")                   '�������t�i�{�������j
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 64) = 64) And (CDbl(GF_Com_CutNumber(strDate)) <> CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then      '�O���c�Ɠ��`�F�b�N
                    strMsg = GF_GetMsg("ITG009")                   '�O���c�Ɠ��Ɠ���
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 128) = 128) And CDbl(Left(GF_Com_CutNumber(strDate), 6) < CDbl(Left(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate), 6))) Then  '�O���ȍ~����
                    strMsg = GF_GetMsg("ITG010")                   '�O���ȍ~
                    Call GS_Com_TxtGotFocus(cntControl)
                Else
                    strTxtDat = GF_Com_CutNumber(strDate)
                    If (KeyAscii = 13) Then            '���ݺ��ނȂ�����
                        If (Len(GF_Com_CutNumber(strTxtDat)) - Len(GF_Com_CutNumber(strTxtDat))) < 8 Then
                            Call GF_Com_KeyPress(11, KeyAscii)
                        ElseIf Not (KeyAscii = 9 Or KeyAscii = 8) Then
                            KeyAscii = 0
                        End If
                        Call GS_Com_NextCntl(cntControl)
                    End If
                End If
                
                If Not (cntControl2 Is Nothing) Then
                    '�`�F�b�N����(intInCheck)�Ɋ֌W�����c�Ɠ��`�F�b�N���j���̐F��ݒ肷��
                    If (GF_Com_CheckOpen(strDate) = False) Then
                        cntControl2.ForeColor = &HFF&              '�ԐF�\��
                    End If
                    cntControl2 = strYoubiTbl(Weekday(strDate) - 1)    '�j���ݒ�
                End If
        End Select
    ElseIf (Len(GF_Com_CutNumber(strDate)) - Len(GF_Com_CutNumber(strDate))) < 8 Then
        Call GF_Com_KeyPress(11, KeyAscii)
    ElseIf Not (KeyAscii = 9 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    GF_Com_KeyPressDate = strMsg
End Function

Public Function GS_Com_TxtCvtYmd(strDate As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ������ҏW
' �@�\   : ���͂��ꂽ������� YYYY/MM/DD �ɕҏW����
' ����   : strDate As String     ���͓��t
' �߂�l : String   �ҏW�㕶����
'                       "�h          �@���͖���
'                       "Not Date�h�@�@���݂��Ȃ��N���̏ꍇ
'                       "Less Than"    8�������A����P�O�O�O�N�����̏ꍇ
' ���l   :
'------------------------------------------------------------------------------
    Dim strInData As String    ' �R���g���[���e�L�X�g���e
    Dim strFmtData As String   ' �ҏW����̓f�[�^
    
    GS_Com_TxtCvtYmd = ""
    
    If IsDate(strDate) = True Then
        strFmtData = Format(strDate, "YYYY/MM/DD")
    Else
        ' �X���b�V���̃J�b�g
        strInData = GF_Com_CutNumber(strDate)
        If Len(strInData) = 0 Then
            Exit Function
        ElseIf Len(strInData) <> 8 Then
            GS_Com_TxtCvtYmd = "Less Than"
            Exit Function
        End If
        
        strFmtData = Format(strInData, "####/##/##")
    End If
    
    '���݃`�F�b�N
    If (IsDate(strFmtData)) = True Then
        '�ҏW����̓f�[�^��Ԃ�
        GS_Com_TxtCvtYmd = strFmtData
    Else
        GS_Com_TxtCvtYmd = "Not Date"
    End If
   
End Function

Public Function GF_Com_CheckOpen(strInDate As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �c�Ɠ�����
' �@�\   : �c�Ɠ��A��c�Ɠ��𔻕ʂ���
' ����   : strDate As String     ���͓��t
' �߂�l : Boolean   �c�Ɠ�:TRUE�A��c�Ɠ��FFALSE
' ���l   :
'------------------------------------------------------------------------------
    Dim strDate     As String    ' �R���g���[���e�L�X�g���e
    Dim strFmtData  As String   ' �ҏW����̓f�[�^
    Dim strDataFlag As String

    GF_Com_CheckOpen = False
    
    If IsDate(strInDate) = True Then
        strFmtData = Format(strInDate, "YYYY/MM/DD")
    Else
        ' �X���b�V���̃J�b�g
        strDate = GF_Com_CutNumber(strInDate)
        If Len(strDate) <> 8 Then Exit Function
        
        strFmtData = Format(strDate, "####/##/##")
    End If
    
    '���݃`�F�b�N
    If (IsDate(strFmtData)) = True Then
        Select Case Weekday(strFmtData)
            Case vbMonday, vbTuesday, vbWednesday, vbThursday, vbFriday
                GF_Com_CheckOpen = GF_Com_CheckHoriday(strFmtData)
        End Select
    End If
    
End Function

Public Function GF_Com_CheckHoriday(strInDate As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �c�Ɠ�����
' �@�\   : �c�Ɠ��A��c�Ɠ�������Ͻ������Ĕ��ʂ���
' ����   : strDate As String     ���͓��t
' �߂�l : Boolean   �c�Ɠ�:TRUE�A��c�Ɠ��FFALSE
' ���l   :
'------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTableName As String
    Dim oraDyna    As OraDynaset
    
    GF_Com_CheckHoriday = True
    
    '�j�Փ��t�@�C�����o
    strSQL = "SELECT * FROM TGCLMR"
    strSQL = strSQL & " WHERE YMD = '" & strInDate & "'"
    strSQL = strSQL & "   AND HRDFLG = '1'"
    Set oraDyna = gOraDataBase.dbcreatedynaset(strSQL, ORADYN_NOCACHE)
    
    '���R�[�h�����݂����ꍇ���s
    If oraDyna.RecordCount > 0 Then
        GF_Com_CheckHoriday = False
    End If
    
End Function

Public Function GF_LenString(strChar As String) As Integer
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �������̒������޲Đ��ŕԂ�
' �@�\   :
' ����   : strChar As String    ������
' �߂�l : Integer              ��������޲Đ�
' ���l   :
'------------------------------------------------------------------------------
    GF_LenString = LenB(strConv(strChar, vbFromUnicode))
End Function

Public Function GF_Left(strChar As String, intCount As Integer) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ������̍�����w���޲Đ����擾����
' �@�\   :
' ����   : strChar As String    ������
'          intCount As Integer  �޲Đ�
' �߂�l : String      �w���޲Đ����̕�����
' ���l   :
'------------------------------------------------------------------------------
    GF_Left = strConv(LeftB(strConv(strChar, vbFromUnicode), intCount), vbUnicode)
End Function

Public Function GF_Mid(strChar As String, intStrat As Integer, intCount As Integer) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ������̎w��ʒu����w���޲Đ����擾����
' �@�\   :
' ����   : strChar As String    ������
'          intStrat As Integer  �J�n�ʒu(�޲Đ�)
'          intCount As Integer  �޲Đ�
' �߂�l : String      �w���޲Đ����̕�����
' ���l   :
'------------------------------------------------------------------------------
    GF_Mid = strConv(MidB(strConv(strChar, vbFromUnicode), intStrat, intCount), vbUnicode)
End Function

Public Function GF_Right(strChar As String, intCount As Integer) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ������̉E����w���޲Đ����擾����
' �@�\   :
' ����   : strChar As String    ������
'          intCount As Integer  �޲Đ�
' �߂�l : String      �w���޲Đ����̕�����
' ���l   :
'------------------------------------------------------------------------------
    GF_Right = strConv(RightB(strConv(strChar, vbFromUnicode), intCount), vbUnicode)
End Function

Public Function GF_CheckContract(bolMoveKbn As Boolean, strYear As String, strMonth As String, strContractKbn As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �����Z�o
' �@�\   :
' ����   : bolMoveKbn As Boolean   True�F�O�����AFalse�F�ゾ��
'          strYear  �@As String    �N
'          strMonth �@As String    ��
'          strContractKbn As String�@1�`28�F���t�A38�F����-2�A39�F����-1�A40�F����
' �߂�l : String          �␳�������t(�װ����"")
' ���l   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strDate     As String
    Dim intStep     As Integer
    Dim intRet      As Integer
    
    GF_CheckContract = ""
    
    '�O�����A�ゾ����
    If bolMoveKbn = False Then
        intStep = 1
    Else
        intStep = -1
    End If
    
    '���t�擾
    Select Case strContractKbn
    Case "38"
        '���� - 2
        strDate = Format(DateAdd("D", -3, DateAdd("M", 1, strYear & "/" & strMonth & "/01")), "YYYY/MM/DD")
    Case "39"
        '���� - 1
        strDate = Format(DateAdd("D", -2, DateAdd("M", 1, strYear & "/" & strMonth & "/01")), "YYYY/MM/DD")
    Case "40"
        '����
        strDate = Format(DateAdd("D", -1, DateAdd("M", 1, strYear & "/" & strMonth & "/01")), "YYYY/MM/DD")
    Case Else
        strDate = strYear & "/" & strMonth & "/" & strContractKbn
        
        '���t����
        If IsDate(strDate) = False Then
            intRet = GF_MsgBoxDB("", "WTG022", "OK", "E")
            Exit Function
        End If
        
        strDate = Format(strDate, "YYYY/MM/DD")
    End Select
    
    '�y���A�j������
    Do
        '�c�Ɠ��ɂȂ����甲����
        If GF_Com_CheckOpen(strDate) = True Then Exit Do
        
        '1�����炷
        strDate = Format(DateAdd("D", intStep, strDate), "YYYY/MM/DD")
    Loop
    
    GF_CheckContract = strDate
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_CheckContract")
    
End Function

Public Function GF_ChangeQuateSing(strString As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :  �V���O���N�H�[�e�[�V�����ϊ�
' �@�\      :�@�V���O���N�H�[�e�[�V�������I���N���f�[�^�x�[�X�ɓo�^�ł���`���ɕϊ�����
' ����     �F strString   As String (I)    ���肷�镶����
' �߂�l    :  �ϊ���̕�����
' ���l      :
'------------------------------------------------------------------------------
    GF_ChangeQuateSing = Replace(strString, "'", "''", 1, , vbBinaryCompare)
End Function

'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :  �A���p�T���h(&)�����x���ɕ\���ł���`���ɂ���
' �@�\      :
' ����     �F strChar   As String (I)    ������
' �߂�l    :  ������
' ���l      :
'------------------------------------------------------------------------------
Public Function GF_ReplaceAmper(strChar As String) As String
    GF_ReplaceAmper = Replace(strChar, "&", "&&", 1, , vbBinaryCompare)
End Function

'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :  ���x���ɕ\�����Ă���A���p�T���h(&)��߂����`���ɂ���
' �@�\      :
' ����     �F strChar   As String (I)    ������
' �߂�l    :  ������
' ���l      :
'------------------------------------------------------------------------------
Public Function GF_UndoAmper(strChar As String) As String
    GF_UndoAmper = Replace(strChar, "&&", "&", 1, , vbBinaryCompare)
End Function


Public Function GF_FileNameRestrinction(strName As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   ̧�ٖ���������
' �@�\      :   �o�^�\��̧�ٖ�����������
' ����      :   strName AS String   ̧�ٖ�
' �߂�l    : �@True:���� / False:̧�ٖ������װ
' ���l      :
'------------------------------------------------------------------------------
    On Error Resume Next
    
    Dim i         As Integer
    Dim intAsc    As Integer
    
    GF_FileNameRestrinction = False
    
    For i = 1 To Len(strName)
        intAsc = Asc(Mid(strName, i, 1))
        
        '�����A�p��(�啶��,������)�Aʲ�݁A���ް�ް
        If (47 < intAsc And intAsc < 58) _
          Or (64 < intAsc And intAsc < 91) _
          Or (96 < intAsc And intAsc < 123) _
          Or (intAsc = 45) _
          Or (intAsc = 95) Then
            
        Else
            Exit Function
        End If
    Next i
    
    GF_FileNameRestrinction = True
End Function

Public Function GF_CheckNumeric(ByVal strNum As String, _
                                Optional ByVal blnIntegerFlg As Boolean = True _
                                ) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���l�`�F�b�N
' �@�\   :
' ����   : strNum As String             '�`�F�b�N�Ώە�����
'          blnIntegerFlg As Boolean     '�����`�F�b�N�t���O
'                                         True:�����`�F�b�N���� ,False:�Ȃ�
' �߂�l : True:���l / False:���l�ȊO
' ���l   :
'------------------------------------------------------------------------------
    GF_CheckNumeric = False
    
    If blnIntegerFlg = True Then
        '�����`�F�b�N�ŏ����_������Ƃ��̓G���[
        If InStr(1, strNum, ".", vbTextCompare) > 0 Then Exit Function
    End If
    
    If IsNumeric(strNum) = True Then
        GF_CheckNumeric = True
    End If
End Function

''--------------------------------------------------------------------------------
'' @(f)
'' �@�\�T�v : ���l�`�F�b�N(IsNumeric[3E4/3E+4/3E-4/(10)/\1,000/10.5/&12/12/12+/0001/&HFF/�P�Q�R(�S�p)]����ł��Ȃ�)
''
'' ����     : ByVal strNumber As String         �`�F�b�N�f�[�^
''          : Optional blnFlg As Boolean        TRUE:�}�C�i�X��, FALSE:�}�C�i�X�s��
''
'' �߂�l   : TRUE�F���� FALSE�F�ُ�  Boolean
''--------------------------------------------------------------------------------
Public Function GF_CheckNumber2(ByVal strNumber As String, Optional blnFlg As Boolean = True) As Boolean
    Dim intLen As Integer

    If Left(strNumber, 1) = "-" Then
        If blnFlg Then      ' �}�C�i�X��
            strNumber = Mid(strNumber, 2)
        Else                ' �}�C�i�X�s��
            Exit Function
        End If
    End If

    intLen = Len(strNumber)

    ' ������ɑS�p���܂܂�Ă�����֐��𔲂��� or �����񂪋�Ȃ甲����
    If LenB(strConv(strNumber, vbFromUnicode)) <> intLen Or intLen = 0 Then Exit Function

    ' �����񂪂��ׂĐ����ō\������Ă邩?
    If strNumber Like String$(intLen, "#") Then GF_CheckNumber2 = True

End Function

Public Function GF_SearchCount(ByVal strChar As String, ByVal strSearch As String) As Integer
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :  �Ώۂ̕����̌������擾����
' �@�\      :
' ����     �F strChar   As String (I)    ������
'            strSearch As String(I)     �J�E���g���镶����
' �߂�l    :  ������
' ���l      :
'------------------------------------------------------------------------------
    Dim i           As Integer
    Dim intPoint    As Integer
    Dim intCount    As Integer
    
    intCount = 0
    
    For i = 1 To Len(strChar)
        intPoint = InStr(i, strChar, strSearch, vbTextCompare)
        If intPoint > 0 Then
            intCount = intCount + 1
            i = intPoint
        Else
            Exit For
        End If
    Next i
    
    GF_SearchCount = intCount
    
End Function

Public Function GF_CharPermitChek(intPatan As Integer, ByVal strChar As String) As Integer
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ��������`�F�b�N
' �@�\   : �����ꂽ�����ȊO�̂��̂��܂܂��ꍇ�̓G���[�Ƃ���
' ����   : intPatan As Integer     ''���������
'                1  - ����  Code Non Check  "0,1,2,�`9"
'                2  - �����{��ص�� Code Non Check   "0,1,2,�`9,.,"
'                3  - �����{��ص�ށ{ϲŽ Code Non Check   "0,1,2,�`9,.,-"
'                4  - �����{�p�� Code Non Check   "0,1,2,�`9,A�`Z"
'                5  - �����{ϲŽ Code Non Check   "0,1,2,�`9,-"
'                6  - �����{�p�� Code Non Check   "0,1,2,�`9,A�`Z,a�`z,"
'                7  - '!' �` '}' �܂�OK    (���ނ��� 33 �` 125�܂�)
'                8  - �p�����{��׽�{ϲŽ�{"*" Code Non Check   "0,1,2,�`9,A�`Z,a�`z,+,-,*"
'                9  - �p�啶�� Code Non Check   "A�`Z"
'                10 - �����{ʲ�݁{�ׯ�� Code Non Check   "0,1,2,�`9,-,/"
'                11 - �����{ʲ�݁{��� Code Non Check   "0,1,2,�`9,-,(,)"
'                12 - �����{�p���{ʲ�� Code Non Check   "0,1,2,�`9,A�`Z","-"
'                13 - �����{�p���{���ݸ Code Non Check   "0,1,2,�`9,A�`Z"," "
'                14 - �����{���ݸ Code Non Check   "0,1,2,�`9," "
'                15 - ASCII����(32�`126)�̕���
'                16 - ASCII����(32�`126)+�g��ASCII����(160�`223)�̕���
'          strChar As Integer       ''������
' �߂�l : Integer          ''0����L���ȕ�����   0�ȊO��������ȕ����񂪍ŏ��Ɍ��������ʒu
' ���l   :
'------------------------------------------------------------------------------
    Dim i       As Integer
    Dim intAsc  As Integer
    Dim intLen  As Integer
    
    GF_CharPermitChek = 0
    
    intLen = Len(strChar)
    
    If intLen = 0 Then
        Exit Function
    End If
    
    '��������� ����
    Select Case intPatan
        Case 1          ''1-9
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) Then
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
            
        Case 2          '' 0-9 or .(46)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If (intAsc < 48) Or (intAsc > 57) Then
                    If intAsc <> 46 Then
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
            
        Case 3          '' 0-9 or -(45) or .(46)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If (intAsc < 48) Or (intAsc > 57) Then
                    If intAsc <> 45 And intAsc <> 46 Then
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
            
        Case 4          '' 0-9 or A-Z
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   ((intAsc < 65) Or (intAsc > 90)) Then
                   GF_CharPermitChek = i
                   Exit Function
                End If
            Next i
            
        Case 5          '' 0-9 or -(45)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And (intAsc <> 45) Then
                       GF_CharPermitChek = i
                       Exit Function
                End If
            Next i
            
        Case 6          '' 0-9 or A-Z or a-z
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   ((intAsc < 65) Or (intAsc > 90)) And _
                   ((intAsc < 97) Or (intAsc > 122)) Then
                   GF_CharPermitChek = i
                   Exit Function
                End If
            Next i
            
        Case 7          ''  "!" �` "}"
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If intAsc < 33 Or intAsc > 125 Then
                   GF_CharPermitChek = i
                   Exit Function
                End If
            Next i
            
        Case 8          '' 0-9 or A-Z or a-z or +(43) or -(45) or *(42)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                        ((intAsc < 65) Or (intAsc > 90)) And _
                        ((intAsc < 97) Or (intAsc > 122)) Then
                    If intAsc = 43 Or intAsc = 45 Or intAsc = 42 Then
                    
                    Else
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
                        
        Case 9          ''A-Z
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 65) Or (intAsc > 90)) Then
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
            
        Case 10          '' 0-9 or -(45) or /(47)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If (intAsc < 48) Or (intAsc > 57) Then
                    If intAsc <> 45 And intAsc <> 47 Then
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
            
        Case 11          '' 0-9 or -(45) or ((40) or )(41)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If (intAsc < 48) Or (intAsc > 57) Then
                    If intAsc <> 45 And intAsc <> 40 And intAsc <> 41 Then
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
            
        Case 12          '' 0-9 or A-Z or "-"
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))

'2008/02/21 Updated by N.Kigaku Start-------------------------
'ʲ�݂�ASCII���ޏC��
'                If ((intAsc < 48) Or (intAsc > 57)) And _
'                   ((intAsc < 65) Or (intAsc > 90)) And _
'                   (intAsc <> 32) Then
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   ((intAsc < 65) Or (intAsc > 90)) And _
                   (intAsc <> 45) Then
'2008/02/21 Update End ---------------------------------------
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
            
        Case 13          ''0-9 or A-Z or " "
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   ((intAsc < 65) Or (intAsc > 90)) And _
                   (intAsc <> 32) Then
                    
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
        
        Case 14          ''0-9 or " "
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   (intAsc <> 32) Then
                    
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i

'2011/05/20 Added by N.Kigaku Start --------------------------------
        Case 15         ''ASCII����(32�`126)�̕���
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))

                If (intAsc < 32) Or (intAsc > 126) Then
                    
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i

        Case 16         ''ASCII����(32�`126)+�g��ASCII����(160�`223)�̕���
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 32) Or (intAsc > 126)) And _
                   ((intAsc < 160) Or (intAsc > 223)) Then
                    
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
'2011/05/20 End ----------------------------------------------------

        Case Else
    
    End Select
    
End Function

''--------------------------------------------------------------------------------
'' @(f)
'' �@�\�T�v : ���t�^���`�F�b�N
''
'' ����     : ByRef strDt As String             �`�F�b�N�f�[�^
''          : ByRef blnEnd As Boolean           True = '99999999'������, False = '99999999'�͕s���Ƃ���
''
'' �߂�l   : TRUE�F���� FALSE�F�ُ�  Boolean
''--------------------------------------------------------------------------------
Public Function GF_DateConv(strDt As String, Optional blnEnd As Boolean = True) As Boolean
    Dim strConv As String

    GF_DateConv = False

    '' 8���ł͂Ȃ� or ���l�ł͂Ȃ�
    If Len(strDt) <> 8 Or IsNumeric(strDt) = False Then Exit Function

    '' �S�p�������܂܂��ꍇ�̓G���[
    If GF_LenString(strDt) <> Len(strDt) Then Exit Function

    ''YYYY/MM/DD�`���ɂ���
    strConv = Mid(strDt, 1, 4) & "/" & Mid(strDt, 5, 2) & "/" & Mid(strDt, 7, 2)
   
    '' 99999999�͎g�p��?
    If blnEnd = True Then
        If IsDate(strConv) = True Or strDt = "99999999" Then
            GF_DateConv = True
        End If
    ElseIf IsDate(strConv) = True Then
        GF_DateConv = True
    End If

End Function

Public Function GF_THJCMBXMR_CHK(ByVal strCMBNAME As String, ByVal strCDVAL As String) As Boolean
'--------------------------------------------------------------------------------
' @(f)
' �@�\���@�@: ���ʕ\��ؽ�ð��ٌ���
' �@�\�@�@�@: ���ʕ\��ؽ�ð��ق��������A�Y���ް������邩��������
' �����@�@�@: strCMBNAME        ''ؽĕ\�����e
' �@�@�@�@�@: strCDVAL          ''ؽĕ����l
' �߂�l�@�@: True�F�ް��L��    False�F�ް�����
' �@�\�����@:
'--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strSQL   As String
    Dim Dynaset   As OraDynaset
        
    GF_THJCMBXMR_CHK = False
        
    strSQL = ""
    strSQL = strSQL & "SELECT ROWID"
    strSQL = strSQL & "  FROM THJCMBXMR"
    strSQL = strSQL & " WHERE CDVAL ='" & strCDVAL & "'"
    strSQL = strSQL & "   AND CMBNAME ='" & strCMBNAME & "'"
    
    Set Dynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    '���ʕ\��ð����ް��L������
    If (Dynaset.EOF = True) Then
    '�ް�����
        GF_THJCMBXMR_CHK = False
    Else
    '�ް��L��
        GF_THJCMBXMR_CHK = True
    End If
    
    Set Dynaset = Nothing
    
    Exit Function
        
ErrHandler:
    Call GS_ErrorHandler("GF_THJCMBXMR_CHK")

End Function

''--------------------------------------------------------------------------------
'' @(f)
'' �@�\�T�v : �����_�����`�F�b�N
''
'' ����     : ByRef strDt As String             �`�F�b�N�f�[�^
''          : ByRef intInt As Integer           ����������
''          : ByRef intDec As Integer           ����������
''          : Optional blnFlg As Boolean        TRUE:�}�C�i�X��, FALSE:�}�C�i�X�s��
''
'' �߂�l   : 0:����, -1:�����Ⴄ, -2:���l�ł͂Ȃ�
''--------------------------------------------------------------------------------
Public Function GF_ChkDeci(ByVal strDt As String, intInt As Integer, intDec As Integer, Optional blnFlg As Boolean = True) As Integer
    On Error GoTo ErrHandler
    
    Dim intLen As Integer, intCnt As Integer
    Dim strDec() As String

    GF_ChkDeci = 0

    '' �}�C�i�X?
    If Left(strDt, 1) = "-" Then
        If blnFlg Then  '' �}�C�i�X����?
            strDt = Mid(strDt, 2)   '' �}�C�i�X������
        Else
            GF_ChkDeci = -2
            Exit Function
        End If
    End If

    strDec = Split(strDt, ".")
    If UBound(strDec) > 0 Then  '' �����l?
        '' �����`�F�b�N
        If Len(strDec(0)) > intInt Then GF_ChkDeci = -1
        If Len(strDec(1)) > intDec Then GF_ChkDeci = -1

        '' ���l�^�`�F�b�N
        If GF_CheckNumber2(strDec(0) & strDec(1)) = False Then GF_ChkDeci = -2
    Else                        '' �����l
        '' �����`�F�b�N
        If Len(strDt) > intInt Then GF_ChkDeci = -1

        '' ���l�^�`�F�b�N
        If GF_CheckNumber2(strDec(0)) = False Then GF_ChkDeci = -2
    End If

    Exit Function
        
ErrHandler:
    Call GS_ErrorHandler("GF_ChkDeci")
End Function

Public Function GF_OptFormatChk(strOpt As String, strSize As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:OPT̫�ϯ�����
' �@�\�@�@�@:OPT̫�ϯ�����
' �����@�@�@:strOpt :OPT����
' �@�@�@�@�@:strSize:���ދ敪(0:���ޖ�,1:���ލ�)
' �@�\�����@:OPT̫�ϯ�����
'------------------------------------------------------------------------------
    ''�ϐ���`
    Dim lngOptCount     As Long ''���ޖ�OPT����
    Dim lngOptSizeCount As Long ''���ޗLOPT����
    Dim intIdx          As Integer ''����
    Dim strChar         As String ''OPT1����

    ''������Ԑݒ�
    GF_OptFormatChk = False
    
    ''�ϐ�������
    lngOptCount = mlngOptLength
    lngOptSizeCount = mlngSizeOptLength
    intIdx = 1
    
    ''OPT��������
    If GF_LenString(Trim(strOpt)) = lngOptCount And strSize = "0" Then
    ElseIf GF_LenString(Trim(strOpt)) = lngOptSizeCount And strSize = "1" Then
    Else
        Exit Function
    End If
    
    ''̫�ϯ�����
    Do Until intIdx > lngOptCount
        strChar = ""
        If intIdx = 2 Or intIdx = 3 Then
        ''2����,3����(���l)
            strChar = GF_Mid(Trim(strOpt), intIdx, 1)
            If GF_CheckNumeric(strChar) = False Then
                Exit Function
            End If
        Else
        ''1����,4����(�p��)
            strChar = GF_Mid(Trim(strOpt), intIdx, 1)
            If GF_Com_CheckString(10, strChar) = False Then
                Exit Function
            End If
        End If
        intIdx = intIdx + 1
    Loop

    ''������ԍĐݒ�
    GF_OptFormatChk = True

End Function

Public Function GF_CheckLinefeed(ByVal strChar As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:���s���ޓ�������
' �@�\�@�@�@:
' �����@�@�@:strChar As String      �������镶����
' �@�\�����@:
' �߂�l�@�@:False:���s���� / True:���s�Ȃ�
'------------------------------------------------------------------------------
    If 0 < InStr(1, strChar, vbCrLf, vbBinaryCompare) Then
        GF_CheckLinefeed = False
    Else
        GF_CheckLinefeed = True
    End If
End Function

'2006/12/11 Added by N.Kigaku
''--------------------------------------------------------------------------------
'' @(f)
'' �@�\�T�v : ���p�p�����L���`�F�b�N
''
'' ����     : ByVal strDt As String             �`�F�b�N�f�[�^
''
'' �߂�l   : TRUE�F���� FALSE�F�ُ�  Boolean
''--------------------------------------------------------------------------------
Public Function GF_CheckEngNumMark(ByVal strDt As String, Optional blnCRLF_Flag As Boolean = False) As Boolean
    Dim intCnt As Integer
    Dim lngChk_Asc As Long

    '' ���p�p�����L��
    For intCnt = 1 To Len(strDt)
        lngChk_Asc = Asc(Mid(strDt, intCnt, 1))
        
        If (blnCRLF_Flag = True) And (lngChk_Asc = 10 Or lngChk_Asc = 13) Then
            '���s���ނ͑ΏۊO

        ElseIf Asc(" ") > lngChk_Asc Or Asc("~") < lngChk_Asc Then
            GF_CheckEngNumMark = False
            Exit Function

        End If
    Next

    GF_CheckEngNumMark = True

End Function

'2008/05/27 Added by N.Kigaku
Public Function GF_ChangeQuateDouble(strString As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :  �_�u���N�H�[�e�[�V�����ϊ�
' �@�\      :�@�_�u���N�H�[�e�[�V�������o�͂ł���`���ɕϊ�����
' ����     �F strString   As String (I)    ���肷�镶����
' �߂�l    :  �ϊ���̕�����
' ���l      :�@CSV̫�ϯďo�͎��̕ϊ��Ȃ�
'------------------------------------------------------------------------------
    GF_ChangeQuateDouble = Replace(strString, """", """""", 1, , vbBinaryCompare)
End Function


'Public Function GF_Com_KeypressCif(KeyAscii As Integer, crtControl As Control, strTxtData As String, crtOutControl As Control) As String
''------------------------------------------------------------------------------
'' @(f)
'' �@�\�� : �����ԍ��̃`�F�b�N�֐�
'' �@�\   : �Y������� strTxtData �ɐݒ�l��ۑ�����B�܂��A�Y�������ڋq���̂� crtOutControl �ɕ\������
'' ����   : KeyAscii As Integer      13(���ݺ��ށj ���w�肷��Ǝ��̃t�B�[���h�Ƀt�H�[�J�X���ړ�����
''                                   0 ���w�肷��ƃt�H�[�J�X�̈ړ��Ȃ��Ń`�F�b�N���s����
''                                   13(���ݺ��ށj�y�� 0 �̏ꍇ�A�G���[�������͓��͈�Ƀt�H�[�J�X�����킹��
''          crtControl As Control    �����ԍ����̓R���g���[��
''          strTxtData As String     ���o�͒l
''          crtOutControl As Control ���̕\���p�R���g���[��
'' �߂�l : String   �G���[���b�Z�[�W
'' ���l   :
''------------------------------------------------------------------------------
'    Dim strMsg      As String
'    Dim strCifName  As String
'    Const intMaxlen As Integer = 5
'
'    strMsg = ""
'    If (KeyAscii = 13) Or (KeyAscii = 0) Then           '���ݺ��ނ��m�F�p���ނȂ�����
'        If crtControl.Text <> "" Then
'            crtControl.Text = Format(crtControl.Text, "00000")
'            strCifName = GF_Com_Cifget(crtControl)
'
'            If Len(strCifName) = 0 Then                   '���̂��擾�ł�����
'                strMsg = GF_GetMsg("WTH004")
'                Call GS_Com_TxtGotFocus(crtControl)
'            Else
'                crtOutControl.Caption = strCifName
'                strTxtData = crtControl.Text
'                If (KeyAscii = 13) Then            '���ݺ��ނȂ�����
'                    If (Len(crtControl.Text) - Len(crtControl.SelText)) < intMaxlen Then
'                        Call GF_Com_Keypress(1, KeyAscii)
'                    Else
'                        KeyAscii = 0
'                    End If
'                End If
'            End If
'        Else
'            If crtControl.Enabled = True Then
'                Call GS_Com_TxtGotFocus(crtControl)
'            Else
''                MsgBox "�����ɖ��͂���܂��񂪒����˗������ĉ�����" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "���W���[���� : GF_Com_KeypressCif" & Chr(13) & Chr(10) & "�v���O�������� : " & G_APL_Job1 & G_APL_Job2 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "OK�{�^�����N���b�N�������𑱍s���ĉ�����"
'            End If
'            strMsg = GF_GetMsg("ITH005") & GF_GetMsg("ITH001")
'        End If
'    Else
'        If (Len(crtControl.Text) - Len(crtControl.SelText)) < intMaxlen Then
'            Call GF_Com_Keypress(1, KeyAscii)
'        Else
'            KeyAscii = 0
'        End If
'    End If
'
'    GF_Com_KeypressCif = strMsg
'
'End Function
'
'Public Function GF_Com_Cifget(vntControl As Control) As String
''------------------------------------------------------------------------------
'' @(f)
'' �@�\�� : �����ԍ��ɊY�����鍀�ڒl�����߂�
'' �@�\   :
'' ����   : vntControl As Control   �����ԍ��R���g���[��
'' �߂�l : String   ����於��
'' ���l   :
''------------------------------------------------------------------------------
'    Dim strChar As String    '�����o�����L�[
'
'    '�������R���g���[���̏ꍇ
'    If IsObject(vntControl) Then
'        Select Case UCase(Left(vntControl.Name, 3))
'            Case "TXT"
'                strChar = vntControl.Text
'            Case "LBL"
'                strChar = vntControl.Caption
'            Case "CBO", "LST"
'                strChar = vntControl.List(vntControl.ListIndex)
'            Case Else
'                strChar = vntControl
'        End Select
'    Else
'        strChar = vntControl
'    End If
'
'    GF_Com_Cifget = GF_Com_CifName(strChar)
'
'End Function
'
' 2017/01/11 �� M.Tanaka K545 CS�v���Z�X���P  ADD
Public Function GF_CheckStartToEnd(dtStart As Variant, dtEnd As Variant, intCheck As Integer, intKbn As Integer) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :  ���t�̊��ԃ`�F�b�N�֐�
' �@�\      :�@���t�̊��Ԃ��`�F�b�N����
' ����      :  dtStart      As Variant      �J�n��
'              dtEnd        As Variant      �I����
'              intCheck     As integer      �`�F�b�N����
'              strKbn       As String       �`�F�b�N�敪 (1�F�N�A2�F���A3�F��)
' �߂�l    :  True �F ��v / False : �s��v
' ���l      :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim dtStart_After As Date '�`�F�b�N���ԉ��Z��J�n��
    
    GF_CheckStartToEnd = False
    
    If IsNull(dtStart) = True Or IsDate(dtStart) = False Then
        'Null��������Date�^�ł͂Ȃ��ꍇfalse�ŕԂ�
        Exit Function
    End If
    
    If IsNull(dtEnd) = True Or IsDate(dtEnd) = False Then
        'Null��������Date�^�ł͂Ȃ��ꍇfalse�ŕԂ�
        Exit Function
    End If
    
    Select Case intKbn
        Case CSTE_Year
            dtStart_After = DateTime.DateAdd("yyyy", intCheck, CDate(dtStart))
        Case CSTE_Month
            dtStart_After = DateTime.DateAdd("m", intCheck, CDate(dtStart))
        Case CSTE_Date
            dtStart_After = DateTime.DateAdd("d", intCheck, CDate(dtStart))
        Case Else
            GoTo ErrHandler
    End Select
    
    If dtStart_After = CDate(dtEnd) Then
        GF_CheckStartToEnd = True
    End If
    
    Exit Function
ErrHandler:
    Call GS_ErrorHandler("GF_CheckStartToEnd")
    Err.Raise Number:=vbObjectError, Description:="GF_CheckStartToEnd�ŃG���[���������܂����B"
End Function
' 2017/01/11 �� M.Tanaka K545 CS�v���Z�X���P  ADD
