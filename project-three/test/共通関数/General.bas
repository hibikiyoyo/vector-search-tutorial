Attribute VB_Name = "basGeneral"
' @(h) General.bas  ver1.00 ( 2000/09/28 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   �v���W�F�N�g��  : TLF��ۼު��
'   ���W���[����    : basGeneral
'   �t�@�C����      : General.bas
'   Version        : 1.00
'   �@�\����       �F ��ʋ��ʊ֐�
'   �쐬��         �F T.Fukutani
'   �쐬��         �F 2000/09/28
'   �C������       �F 2001/03/19 T.Fukutani ����޲�۸ޕ\���֐��ǉ�
'   �@�@�@�@       �F 2001/04/24 N.Kigaku FormLoad�����֐��ǉ�
'   �@�@�@�@       �F 2001/04/27 N.Kigaku �d�l�ݒ�NO�̔ԏ���,�{�@�`�s�s�\�����擾�֐��ǉ�
'   �@�@�@�@       �F 2001/11/30 N.Kigaku GF_GetShiyoKbn�C��
'   �@�@�@�@       �F 2001/12/19 N.Kigaku GF_GetNextMitsumoriNo�ǉ�
'   �@�@�@�@       �F 2001/12/20 Takashi.Kato GF_FileOpenDialog�Ɉ���strFileTitle�ǉ�
'   �@�@�@�@       �F 2001/12/21 N.Kigaku GF_GetShiyoKbn_CIF�ǉ�
'   �@�@�@�@       �F 2002/01/23 N.Kigaku GS_Com_NextCntl,GF_FormInit�C��
'   �@�@�@�@       �F 2002/02/27 N.Kigaku GF_ShowHelp�ǉ�,GF_GetNextMitsumoriNo��GF_NumberingShiyoNo�C��
'   �@�@�@�@       �F 2002/03/12 N.Kigaku GF_GetNextMitsumoriNo��GF_NumberingShiyoNo�C��
'   �@�@�@�@       �F 2002/03/30 T.Nono GF_FormInit��GS_Com_NextCntl�C��
'   �@�@�@�@       �F 2005/06/17 N.KIGAKU GF_FileCopy�ǉ�
'                  �F 2006/12/05 N.Kigaku �׸�8.1.7 Nocache�Ή� �������AReadOnly����Nocache�ɕύX
'                  �F 2018/05/15 T.Nakayama K545 CS�v���Z�X���P
'                  : 2021/04/26 R.Kozasa �R�����g�R�s�[�@�\�ǉ�
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ���錾
'------------------------------------------------------------------------------
Option Explicit

Declare Function IsDBCSLeadByte Lib "kernel32" (ByVal bTestChar As Byte) As Long

'------------------------------------------------------------------------------
' �p�u���b�N�萔�錾
'------------------------------------------------------------------------------
Public Const gINPUTCOLOR    As Long = &HC0FFC0   ''���͉\�ق̐F
Public Const gNOTINPUTCOLOR As Long = &H80000005 ''���͕s�\�ق̐F
Public Const gLINECOLOR     As Long = &HFF0000   ''��؂���̐F
Public Const gEXLSMAXROW    As Long = 60000      ''Excel���o����
Public Const gTOTALCOLOR    As Long = &HC0FFFF   ''���v�s�̐F

'------------------------------------------------------------------------------
' �p�u���b�N�ϐ��錾
'------------------------------------------------------------------------------
Public gstrShiyoNo      As String   '�d�l�ݒ�NO
Public gstrHonkiAttKbn  As String   '�{�@ATT�敪
Public gstrRenban       As String   '�A��

'------------------------------------------------------------------------------
' ���W���[���ϐ��錾
'------------------------------------------------------------------------------
Private mfrmFromName    As Form
'2021/04/26�� R.Kozasa �R�����g�R�s�[�@�\�ǉ�
'Private mcrtCntl(200)   As Control
Private mcrtCntl(210)   As Control
'2021/04/26�� R.Kozasa �R�����g�R�s�[�@�\�ǉ�
Private mstrFormName    As String


Public Sub GS_CenteringForm(frmMe As Form, Optional intOption As Integer = 0)
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:�@̫�Ѿ���ݸ�
' �@�\�@�@�@:�@̫�т���ʂ̒��S�Ɉړ�����
' �����@�@�@:�@frmMe As Form   '÷��BOX
'                               0:��ʒ���
'                               1:��i����
' ���l�@�@�@:
'------------------------------------------------------------------------------
    
    Select Case intOption
    Case 0
        '̫�т𒆉��Ɉړ�
        frmMe.Move (Screen.Width - frmMe.Width) / 2, (Screen.Height - frmMe.Height) / 2
    Case 1
        '̫�т���i�����Ɉړ�
        frmMe.Move (Screen.Width - frmMe.Width) / 2, 0
    End Select
End Sub


Public Function GF_FileOpenDialog(objObject As Object, strFilter As String, intFilterIndex As Integer, _
                                    strDir As String, strOpenFile As String, _
                                    Optional strFileTitle As String) As Integer
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:�@̧�ٖ��w�����޲�۸ޕ\��
' �@�\�@�@�@:�@̧�ٖ����w�肵�ĊJ��
' �����@�@�@:�@[i]objObject      As CommonDialog  '����޲�۸޺��۰�
' �@�@�@�@�@:�@[i]strFilter      As String        '̧��̨���������
' �@�@�@�@�@:�@[i]intFilterIndex As Integer       '̨������ޯ��
' �@�@�@�@�@:�@[i]strDir         As String        '�����\���߽
' �@�@�@�@�@:�@[o]strOpenFile    As String        '�ꏊ(�߽+̧�ٖ�)
' �@�@�@�@�@:�@[o]strFileTitle   As String        '̧�ٖ�
' �߂�l�@�@:�@[vbCancel] = ��ݾ����݉�����
' ���l�@�@�@:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    With objObject
        .DialogTitle = "�t�@�C�����J��"
        .Filter = strFilter
        .FilterIndex = intFilterIndex
        .InitDir = strDir
        .CancelError = True
        .Flags = &H80000 Or &H1000 Or &H4  '&H4 = cdlOFNHideReadOnly, &H80000 = cdlOFNExplorer, &H1000 = cdlOFNFileMustExist
        .ShowOpen

        strOpenFile = .FileName
        strFileTitle = .FileTitle
    End With

    GF_FileOpenDialog = vbOK

    Exit Function

ErrHandler:
    GF_FileOpenDialog = vbCancel

End Function

Public Function GF_FileSaveDialog(oObject As Object, sDir As String, sFile As String, sSaveFile As String) As Integer
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:�@̧�ٖ��w�����޲�۸ޕ\��
' �@�\�@�@�@:�@̧�ٖ����w�肵�ĕۑ�
' �����@�@�@:�@[i]oObject As CommonDialog  '����޲�۸޺��۰�
' �@�@�@�@�@:�@[i]sDir As String           '�����\���߽
' �@�@�@�@�@:�@[i]sFile As String          '�����\��̧�ٖ�
' �@�@�@�@�@:�@[o]sSaveFile As String      '�ۑ��ꏊ(�߽+̧�ٖ�)
' �߂�l�@�@:�@[vbCancel] = ��ݾ����݉�����
' ���l�@�@�@:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    With oObject

        oObject.DialogTitle = "�o�͐�w��"
        oObject.Filter = "CSV̧��(.CSV)|*.CSV"
        oObject.InitDir = sDir
        oObject.FileName = sFile
        oObject.CancelError = True
        oObject.Flags = &H2 Or &H4  '&H4 = cdlOFNHideReadOnly, &H2 = cdlOFNOverwritePrompt
        oObject.ShowSave

    End With

    sSaveFile = oObject.FileName
    GF_FileSaveDialog = vbOK

    Exit Function

ErrHandler:

    GF_FileSaveDialog = vbCancel

End Function

Public Sub GF_FormInit(frmForm As Form)
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �R���g���[�����̎擾
' �@�\   : �w��t�H�[���̃R���g���[�������擾���܂�
'          �������̉�ʂ����݂���ꍇ�͉�ʂ��؂�ւ��x�� GF_FormInit ���Ă� Call ���ĉ�����
' ����   : frmForm As Form  �t�H�[���R���g���[��
' ���l   : GF_FormInit�𕡐��� Call ���Ă��������܂���B�A���A�R���g���[���z����g���Ă���ꍇ�͕s�B
'------------------------------------------------------------------------------
    Dim intloop        As Integer
    Dim intM_Idx       As Integer
    Dim strControlName As String
    Dim intCount       As Integer
    
    If mfrmFromName Is Nothing = False Then
        If mstrFormName = frmForm.Name Then
            Exit Sub
        End If
    End If
    
    Set mfrmFromName = frmForm
    mstrFormName = mfrmFromName.Name
 
    intCount = 0
    For intM_Idx = 0 To (mfrmFromName.Count - 1)        '�z��ɃR���g���[����TabIndex���ɐݒ肵�܂�
        For intloop = 0 To (mfrmFromName.Count - 1)
            strControlName = mfrmFromName.Controls(intloop).Name
            Select Case LCase(Left(strControlName, 3))
'>2004/03/30 Upd Nono
''2002/01/23 Update N.Kigaku
'            'Case "txt", "cbo", "lst", "cmd"
'            Case "txt", "cbo", "lst", "cmd", "chk", "opt"
            Case "txt", "cbo", "lst", "cmd", "chk", "opt", "cmb"
'<2004/03/30 Upd Nono
                If mfrmFromName.Controls(intloop).TabIndex = intM_Idx Then
                    Set mcrtCntl(intCount) = mfrmFromName.Controls(intloop)      'TabIndex���ɓ����R���g���[���z���ݒ肵�܂�
                    intCount = intCount + 1
                    Exit For
                End If
            End Select
        Next intloop
    Next intM_Idx
    Set mcrtCntl(intCount) = Nothing        '�����R���g���[���z��̍ŏI����ݒ肵�܂�
    
End Sub

Public Sub GS_Com_NextCntl(crtControl As Control)
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �t�H�[�J�X�ړ�
' �@�\   : GF_FormInit�֐��Ŏw�肵���t�H�[���ɂāA�w�肷��R���g���[��(crtControl)�̎�(TabIndex��)�̃R���g���[����focus�����킹�܂�
' ����   : crtControl As Control  ���ɂȂ�R���g���[��
' ���l   :
'------------------------------------------------------------------------------
    Dim strControlName As String
    Dim intMK_Idx      As Integer
    Dim intCount       As Integer
    Dim intLoopExit  As Integer
    
    If mfrmFromName Is Nothing = True Then
        SendKeys "{TAB}"       'TAB �- SEND(Next Field Cursol)
        Exit Sub
    Else
        If mstrFormName <> Screen.ActiveForm.Name Then
            SendKeys "{TAB}"       'TAB �- SEND(Next Field Cursol)
            Exit Sub
        End If
    End If
            
    intMK_Idx = 0
    Do While mcrtCntl(intMK_Idx) Is Nothing = False
        If mcrtCntl(intMK_Idx).Name = crtControl.Name Then
            intCount = intMK_Idx
'2002/01/23 Delete N.Kigaku
'���۰ٔz����������邽�ߺ���
'            Exit Do
        End If
        intMK_Idx = intMK_Idx + 1
    Loop
        
    If mcrtCntl(intCount + 1) Is Nothing = True Then
        intCount = 0
    Else
        intCount = intCount + 1
    End If
    intLoopExit = 0
    Do While mcrtCntl(intCount) Is Nothing = False
        strControlName = mcrtCntl(intCount).Name
        Select Case LCase(Left(strControlName, 3))
'>2004/03/30 Upd Nono
''2002/01/23 Update N.Kigaku
'        'Case "txt", "cbo", "lst", "cmd"
'        Case "txt", "cbo", "lst", "cmd", "chk", "opt"
        Case "txt", "cbo", "lst", "cmd", "chk", "opt", "cmb"
'<2004/03/30 Upd Nono
            If mcrtCntl(intCount).Visible = True And mcrtCntl(intCount).Enabled = True And mcrtCntl(intCount).TabStop = True Then
                mcrtCntl(intCount).SetFocus
                Exit Sub
            Else
                intCount = intCount + 1
            End If
        Case Else
            intCount = intCount + 1
        End Select
        If mcrtCntl(intCount) Is Nothing = True Then       '�Ō�̃R���g���[���������������x�ŏ��̃R���g���[�����烋�[�v
            If intLoopExit <> 1 Then                     '�i�v���[�v�}�~�Ή�
                intCount = 0
            End If
            intLoopExit = 1
        End If
    Loop
    
End Sub

Public Function GF_MonthCount(strMonth As String) As Integer
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   3���܂ł̌����Z�o
' �@�\      :   ���N�x��3���܂ł̌��������߂�
' ����      :   strMonth As String  ��(YYYY/MM/DD or YYYY/MM)
' �߂�l    :   Integer   ����
' ���l      :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strYear  As String
    Dim strMarch As String
    
    GF_MonthCount = 0
    
    '�N�擾
    strYear = Left(strMonth, 4)
    
    '1�����t��
    strMonth = Left(strMonth, 7) & "/01"
    
    If CInt(Mid(strMonth, 6, 2)) > 3 Then
        strMarch = CStr(CInt(strYear) + 1) & "/03/01"
    Else
        strMarch = strYear & "/03/01"
    End If
    
    GF_MonthCount = DateDiff("M", strMonth, strMarch)
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_MonthCount")
    
End Function

Public Function GF_Year(Optional strYear As String = "NoDate", Optional strMonth As String = "NoDate") As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �C�ӔN���̔N�x�Z�o(�����ȗ����ͼ��ѓ��t�̔N�x��Ԃ�)
' �@�\      :
' ����      :   strYear As String      '�N(�ȗ���)
'               strMonth As String     '��(�ȗ���)
' �߂�l    :   String     �N�x
' ���l      :   2000/12/11
'------------------------------------------------------------------------------
    Dim intYear   As Integer
    Dim intMonth  As Integer
    Dim strDate   As String
    
    '�����ȗ����A�^�p�N���̔N�x�Z�o
    If strYear = "NoDate" Or strMonth = "NoDate" Then
        strDate = Screen.ActiveForm.lblNowDate
        strYear = Left(strDate, 4)
        strMonth = Mid(strDate, 6, 2)
    End If
    
    '�N�x�Z�o����
    intYear = CInt(strYear)
    intMonth = CInt(strMonth)
    If intMonth >= 1 And intMonth <= 3 Then
        intYear = intYear - 1
    End If
    GF_Year = Format(intYear, "0000")
    
End Function

Public Function GF_CutUp(strNum As String, intPoint As Integer) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:�@�����_�؏グ�֐�
' �@�\�@�@�@:�@�����_�ȉ��w�茅��؏グ��
' �����@�@�@:�@strNum   �؏グ�Ώےl
' �@�@�@�@�@:�@intPoint �؏グ���ʒu(�����_�ȉ��؏グ�̏ꍇ = 0)
' �߂�l�@�@:�@�؏グ��l(�؏グ���ʒu��2���w�肵���ꍇ��0.621��0.63)
' ���l�@�@�@:�@[�L���l] �؏グ���ʒu+���x�͈͓̔��̏����_�l
' �@�@�@�@�@:�@[���E�l] ���������������v29���𒴂���Ƶ��ް�۰���܂�
' �@�@�@�@�@:�@���x���グ�邽�ߕϐ��͕�����^���g�p���A���������`����10�i�^(DECIMAL)
'------------------------------------------------------------------------------
    Dim strTemp    As String          '��Ɨ̈�
    Const intSeido As Integer = 5     '���x(�؏グ���ȍ~�̗L���͈�)

    On Error GoTo ErrHandler
    
    GF_CutUp = strNum
    
    ''���l����
    If IsNumeric(strNum) = False Then Exit Function

    strTemp = strNum
    strTemp = Abs(CDec(strTemp)) + ((1 - (1 / (10 ^ (intSeido)))) / (10 ^ intPoint))
    strTemp = Fix(CDec(strTemp) * (10 ^ intPoint)) / (10 ^ intPoint)
    If CDec(strNum) < 0 Then
        strTemp = CDec(strTemp) - (CDec(strTemp) * 2)
    End If

    GF_CutUp = strTemp

    Exit Function

ErrHandler:
    '�װ�����
    Call GS_ErrorHandler("GF_CutUp")

End Function

Public Function GF_WCardChenge(strString As String, intLength As Integer) As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   ���C���h�J�[�h�̍쐬
' �@�\      :
' ����      :   strString As String      '������
'               intLength As Integer     '������
' �߂�l    :   String     �ϊ���̕�����
' ���l      :   �*� ��_� �ɕω�����A������������Ȃ����̂͑���Ȃ����������u_�v��ǉ�
'------------------------------------------------------------------------------
    Dim sWork As String
    
    GF_WCardChenge = ""
    
    sWork = strString & Space(intLength)
    sWork = Replace(sWork, "*", "_")
    sWork = Left(Replace(sWork, " ", "_"), intLength)
    
    GF_WCardChenge = sWork
    
End Function

Public Function GF_Com_BunRuiName(strBunruiFlg As String, strBunrui1 As String, _
                           Optional strBunrui2 As String, Optional strBunrui3 As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ���ޖ��̎擾�֐�
' �@�\�@ : ���ރR�[�h�ɊY�����镪�ޖ��̂��擾����B
' �����@ : strBunruiFlg  As String     ���ދ敪(1�F���ނP, 2:���ނQ, 3:���ނR)
' �@�@�@   strBunrui1    As String     ���޺���1
' �@�@�@   strBunrui2    AS String     ���޺���2
' �@�@�@   strBunrui3    AS String     ���޺���3
' �߂�l : String   ���ޖ���
' ���l   :
'------------------------------------------------------------------------------
'   �ϐ���`
    Dim strSQL        As String
    Dim oDynaset      As OraDynaset
    On Error GoTo ErrHandler
'   �ر
    GF_Com_BunRuiName = ""
'   �����Ώ�DB�ݒ�
    strSQL = ""
    Select Case strBunruiFlg
        Case 1
            If Trim(strBunrui1) = "" Then
                Exit Function
            End If
            strSQL = strSQL & " SELECT BUNRUINAME1 BUNRUINAME"
            strSQL = strSQL & " FROM   THJBUNRUI1  "
            strSQL = strSQL & " WHERE  BUNRUI1     = '" & Trim(strBunrui1) & "'"
        Case 2
            If Trim(strBunrui1) = "" Or Trim(strBunrui2) = "" Then
                Exit Function
            End If
            strSQL = strSQL & " SELECT BUNRUINAME2 BUNRUINAME"
            strSQL = strSQL & " FROM   THJBUNRUI2  "
            strSQL = strSQL & " WHERE  BUNRUI1     = '" & Trim(strBunrui1) & "'"
            strSQL = strSQL & " AND    BUNRUI2     = '" & Trim(strBunrui2) & "'"
        Case 3
            If Trim(strBunrui1) = "" Or Trim(strBunrui2) = "" Or Trim(strBunrui3) = "" Then
                Exit Function
            End If
            strSQL = strSQL & " SELECT BUNRUINAME3 BUNRUINAME"
            strSQL = strSQL & " FROM   THJBUNRUI3  "
            strSQL = strSQL & " WHERE  BUNRUI1     = '" & Trim(strBunrui1) & "'"
            strSQL = strSQL & " AND    BUNRUI2     = '" & Trim(strBunrui2) & "'"
            strSQL = strSQL & " AND    BUNRUI3     = '" & Trim(strBunrui3) & "'"
    End Select
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
'   ں��ނ��Ȃ��ꍇ�װ
    If oDynaset.EOF = True Then
        Exit Function
    End If
'   ���ޖ��̾��
    GF_Com_BunRuiName = GF_VarToStr(oDynaset![BUNRUINAME])
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_Com_BunName", strSQL)
    GF_Com_BunRuiName = " "
    
End Function

Public Function GF_LoadFormProcess(frm As Form, objMe As Object)
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:FormLoad����
' �@�\�@�@�@:�t�H�[�����[�h(��ʑJ��)����
' �����@�@�@:frm    As Form     �t�H�[���I�u�W�F�N�g
'�@�@�@�@�@ :objMe  As Object�@�@Me�I�u�W�F�N�g
' �@�\�����@:�t�H�[�����[�h(��ʑJ��)����
'------------------------------------------------------------------------------
    Load frm
    '�t�H�[�����[�h��������
    If frm.LoadFlag = False Then
        Unload frm
        Screen.MousePointer = vbDefault
        objMe.Enabled = True
        Exit Function
    Else
        Screen.MousePointer = vbDefault
        frm.Show vbModal
        objMe.Enabled = True
        Exit Function
    End If
End Function

Public Function GF_ExChangeQuateSingToDbl(strString As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:�ݸ�ٸ��ð��݂�2�d��
' �@�\�@�@�@:
' �����@�@�@:strString As String (in)    ������
' �߂�l�@�@:�ݸ�ٸ��ð��݂�2�d��������̕�����
' �@�\�����@:
'------------------------------------------------------------------------------
    Dim strTmp()    As String
    Dim strRet      As String
    Dim intCnt      As Integer
    Dim intMax      As Integer
    
    strTmp = Split(strString, "'")

    strRet = ""
    
    intMax = UBound(strTmp)

    If (intMax > 0) Then
        For intCnt = 0 To intMax
            strRet = strRet + strTmp(intCnt) + "''"
        Next intCnt
    Else
        strRet = strString
    End If
    
    GF_ExChangeQuateSingToDbl = strRet
End Function

Public Function GF_NumberingShiyoNo(ByRef strShiyoNo As String, Optional ByVal intShiyuKbn As Integer) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:�d�l�ݒ�NO�̔ԏ���
' �@�\�@�@�@:�d�l�ݒ�NO�̔ԏ���
' �����@�@�@:strShiyoNo As String (out)    �d�l�ݒ�NO
'�@�@�@�@�@�@intShiyuKbn As Integer (in)   �s�A�敪�@1:�����A2:�C�O
' �@�\�����@:�d�l�ݒ�NO�̍̔Ԃ��s�� (���ӁF�g�����U�N�V�����͌Ăь��ōs��)
'------------------------------------------------------------------------------
    Dim strSQL        As String
    Dim oDynaset      As OraDynaset
    Dim oraDSQLStmt   As OraSqlStmt
    Dim strSynoDome   As String
    Dim strSynoFore   As String
    Dim strNendo      As String
    Dim strWareki     As String
    Dim strSysSeireki As String
    Dim strSeireki    As String
    Dim strField      As String
    Dim strMsgTitle   As String
    
    On Error GoTo ErrHandler
    
    GF_NumberingShiyoNo = False
    
    strShiyoNo = ""
    strSynoDome = "0"
    strSynoFore = "0"
    strNendo = "0"
    strWareki = "0"
    strSysSeireki = "0"
    strSeireki = "0"
    
    strMsgTitle = "�d�l�ݒ�NO�̔ԏ���"
    
    ''���@ �V�X�e�����t�̎擾��
    strSQL = ""
    strSQL = strSQL & "SELECT TRUNC(SYSDATE) SYSTEMDATE"
    strSQL = strSQL & "  FROM DUAL"
    
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oDynaset.EOF Then
        '�V�X�e�����t�̎擾�Ɏ��s���܂����B
        Call GF_MsgBoxDB(strMsgTitle, "WTG027", "OK", "E")
        Exit Function
    Else
        strWareki = Format(CStr(oDynaset![SYSTEMDATE]), "e")
        strSysSeireki = Format(CStr(oDynaset![SYSTEMDATE]), "yyyy")
    End If
    
    ''���A �d�l�ݒ�NO�ƔN�x�̎擾��
    strSQL = ""
    strSQL = strSQL & "SELECT NVL(SYNODOME,0)  SYNODOME"
'    strSQL = strSQL & "      ,NVL(SYNOFORE,0)  SYNOFORE"
    strSQL = strSQL & "      ,NVL(NENDO,'0')   NENDO"
    strSQL = strSQL & "  FROM THJSIYOSETNO"
    
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oDynaset.EOF Then
        ''�Y���ް����������͒ǉ�����
        strSQL = ""
        strSQL = strSQL & "INSERT INTO THJSIYOSETNO"
'        strSQL = strSQL & " (SYNODOME,SYNOFORE,NENDO,SEIREKI)"
'        strSQL = strSQL & " (SYNODOME,MTNO,NENDO,SEIREKI)"
        strSQL = strSQL & " (SYNODOME,NENDO)"
        strSQL = strSQL & "VALUES"
'        If intShiyuKbn = 2 Then
'            strSQL = strSQL & " (0,1,'" & strWareki & "','" & strSysSeireki & "')"   '�C�O
'        Else
            strSQL = strSQL & " (1,'" & strWareki & "')"   '����
'        End If
        Set oraDSQLStmt = gOraDataBase.CreateSql(strSQL, ORADYN_NO_AUTOBIND)
        If oraDSQLStmt.RecordCount = 0 Then
            '�X�V���s
            Call GF_MsgBoxDB(strMsgTitle, "WTG028", "OK", "E")
            Exit Function
        End If
        strShiyoNo = Format(strWareki, "00") & "0001"
        GF_NumberingShiyoNo = True
        Exit Function
        
    Else
        strSynoDome = oDynaset![SYNODOME]
'        strSynoFore = oDynaset![SYNOFORE]
        strNendo = oDynaset![NENDO]
        
        ''���B �̔ԏ�����
        '�̔ԃe�[�u���̘a��E������V�X�e���̘a��傫�����܂��͎d�l�ݒ�m�n���ő�l�ɒB��������
        '�d�l�ݒ�NO���N���A���ăV�X�e���̘a���o�^����
        If (CInt(strNendo) < CInt(strWareki)) Or (CInt(strSynoDome) >= 9999) Then
            strSQL = ""
            strSQL = strSQL & "UPDATE THJSIYOSETNO SET "
    '        If intShiyuKbn = 2 Then
    '            strSQL = strSQL & "  SYNODOME=1"
    '            strSQL = strSQL & " ,SYNOFORE=0"
    '        Else
                strSQL = strSQL & "  SYNODOME=1"
'                strSQL = strSQL & " ,SYNOFORE=0"
    '        End If
            If strNendo = "0" Then
                strSQL = strSQL & " ,NENDO='" & strNendo & "'"
            Else
                strSQL = strSQL & " ,NENDO=NENDO+1"
            End If
'            strSQL = strSQL & " ,NENDO='" & strWareki & "'"
'            strSQL = strSQL & " ,SEIREKI='" & strSysSeireki & "'"
        Else
        '    If intShiyuKbn = 1 Then
        '        '����
                strField = "SYNODOME=SYNODOME+1"
        '    Else
        '        '�C�O
        '        strField = "SYNOFORE=SYNOFORE+1"
        '    End If
            strSQL = ""
            strSQL = strSQL & "UPDATE THJSIYOSETNO SET " & strField
            strSQL = strSQL & " WHERE NENDO='" & strNendo & "'"
        End If
        
        Set oraDSQLStmt = gOraDataBase.CreateSql(strSQL, ORADYN_NO_AUTOBIND)
        If oraDSQLStmt.RecordCount = 0 Then
            '�X�V���s
            Call GF_MsgBoxDB(strMsgTitle, "WTG028", "OK", "E")
            Exit Function
        End If
        
        ''���C �̔Ԍ�̎d�l�ݒ�No���擾����
        strSQL = ""
    '    If intShiyuKbn = 1 Then
            strSQL = strSQL & "SELECT NVL(SYNODOME,'1') SYNODOME"
    '    Else
    '        strSQL = strSQL & "SELECT NVL(SYNOFORE,'1') SYNODOME"
    '    End If
        strSQL = strSQL & "      ,NVL(NENDO,'1') NENDO"
        strSQL = strSQL & "  FROM THJSIYOSETNO"
        Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
        If oDynaset.EOF Then
            '�X�V���s
            Call GF_MsgBoxDB(strMsgTitle, "WTG028", "OK", "E")
            Exit Function
        Else
            strShiyoNo = oDynaset![SYNODOME]
            strNendo = oDynaset![NENDO]
        End If
        strShiyoNo = Format(strNendo, "00") & Format(strShiyoNo, "0000")
        
    End If
       
    GF_NumberingShiyoNo = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_NumberingShiyoNo", strSQL)
End Function

Public Function GF_GetHonkiAttHyojiNo(ByVal strSYNO As String, ByRef strHyoji As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:�{�@�`�s�s�\�����擾
' �@�\�@�@�@:�{�@�`�s�s�̕\�������擾����
' �����@�@�@:strSYNO  As String  (in)    �d�l�ݒ�NO
'           strHyoji As String (out)    �\����
' �@�\�����@:�{�@�`�s�s�̕\�������擾����
'------------------------------------------------------------------------------
    Dim strSQL        As String
    Dim oDynaset      As OraDynaset
    
    On Error GoTo ErrHandler
    
    GF_GetHonkiAttHyojiNo = False
    
    strHyoji = ""
    
    strSQL = ""
    strSQL = strSQL & "SELECT HYOZI FROM THJMR "
    strSQL = strSQL & "WHERE SYNO = '" & strSYNO & "'"
    
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oDynaset.EOF Then
        Exit Function
    Else
'        strHyoji = CStr(oDynaset![HYOZI])
        strHyoji = IIf(IsNull(oDynaset![HYOZI]) = True, "", oDynaset![HYOZI])
    End If
    
    GF_GetHonkiAttHyojiNo = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_GetHonkiAttHyojiNo", strSQL)
End Function

Public Function GF_GetNextMitsumoriNo(ByRef strMITSUMORINO As String, ByVal strHanbaitenNo As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:����NO�擾����
' �@�\�@�@�@:����NO�擾����
' �����@�@�@:strMitsumoriNo As String (out)  ����No
'�@�@�@�@�@ :strHanbaitenNo As String (in)   �̔��X����(5��) + �c�Ə�����(2��)
' �@�\�����@:�̔��X������A���Ɏg�p���錩��NO��Ԃ�
'------------------------------------------------------------------------------
    Dim strSQL        As String
    Dim oDynaset      As OraDynaset
    Dim oraDSQLStmt   As OraSqlStmt
    Dim strYear       As String
    Dim strWkMitsumori  As String
    Dim strWareki     As String
    Dim strSysSeireki As String
    Dim strSeireki    As String
    Dim strNumber     As String
    Dim lngNumber     As Long
    Dim strMsgTitle   As String
    
    On Error GoTo ErrHandler
    
    GF_GetNextMitsumoriNo = False
    
    strMsgTitle = "����NO�擾����"
    
    ''���@ �V�X�e�����t�̎擾��
    strSQL = ""
    strSQL = strSQL & "SELECT TRUNC(SYSDATE) SYSTEMDATE"
    strSQL = strSQL & "  FROM DUAL"
    
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oDynaset.EOF Then
        '�V�X�e�����t�̎擾�Ɏ��s���܂����B
        Call GF_MsgBoxDB(strMsgTitle, "WTG027", "OK", "E")
        Exit Function
    Else
        strWareki = Format(CStr(oDynaset![SYSTEMDATE]), "e")
        strSysSeireki = Format(CStr(oDynaset![SYSTEMDATE]), "yyyy")
    End If
    
    ''���A ����NO�ƔN�x�̎擾��
    strSQL = ""
    strSQL = strSQL & "SELECT NVL(MTNO,0)      MTNO"
    strSQL = strSQL & "      ,NVL(SEIREKI,'0') SEIREKI"
    strSQL = strSQL & "  FROM THJSIYOSETNO"
    
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oDynaset.EOF Then
        ''�Y���ް����������͒ǉ�����
        strSQL = ""
        strSQL = strSQL & "INSERT INTO THJSIYOSETNO"
        strSQL = strSQL & " (MTNO,SEIREKI)"
        strSQL = strSQL & "VALUES"
        strSQL = strSQL & " (1,'" & strSysSeireki & "')"
        Set oraDSQLStmt = gOraDataBase.CreateSql(strSQL, ORADYN_NO_AUTOBIND)
        If oraDSQLStmt.RecordCount = 0 Then
            '�X�V���s
            Call GF_MsgBoxDB(strMsgTitle, "WTG028", "OK", "E")
            Exit Function
        End If
        strMITSUMORINO = strHanbaitenNo & strSysSeireki & "00001"
        GF_GetNextMitsumoriNo = True
        Exit Function
        
    Else
        strSeireki = oDynaset![SEIREKI]
        strWkMitsumori = oDynaset![MTNO]
        
        ''���B �̔ԏ�����
        
        '�̔ԃe�[�u���̐�����V�X�e���̐���傫�����܂��͌���No���ő�l�ɒB��������
        '����No���N���A���ăV�X�e���̐����o�^����
        If (CInt(strSeireki) < CInt(strSysSeireki) Or (CLng(strWkMitsumori) >= 99999)) Then
            strSQL = ""
            strSQL = strSQL & "UPDATE THJSIYOSETNO SET "
            strSQL = strSQL & "   MTNO=1"
            If strSeireki = "0" Then
                strSQL = strSQL & "  ,SEIREKI='" & strSysSeireki & "'"
            Else
                strSQL = strSQL & "  ,SEIREKI=SEIREKI+1"
            End If
'            strSQL = strSQL & "  ,SEIREKI='" & strSysSeireki & "'"
        Else
            strSQL = ""
            strSQL = strSQL & "UPDATE THJSIYOSETNO SET MTNO=MTNO+1"
        End If
        Set oraDSQLStmt = gOraDataBase.CreateSql(strSQL, ORADYN_NO_AUTOBIND)
        If oraDSQLStmt.RecordCount = 0 Then
            '�X�V���s
            Call GF_MsgBoxDB(strMsgTitle, "WTG028", "OK", "E")
            Exit Function
        End If
        
        ''���C �̔Ԍ�̌���No���擾����
        strSQL = ""
        strSQL = strSQL & "SELECT NVL(MTNO,'1') MTNO"
        strSQL = strSQL & "      ,NVL(SEIREKI,'1') SEIREKI"
        strSQL = strSQL & "  FROM THJSIYOSETNO"
        Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
        If oDynaset.EOF Then
            '�X�V���s
            Call GF_MsgBoxDB(strMsgTitle, "WTG028", "OK", "E")
            Exit Function
        Else
            strMITSUMORINO = oDynaset![MTNO]
            strSeireki = oDynaset![SEIREKI]
        End If
        strMITSUMORINO = strHanbaitenNo & strSeireki & Format(strMITSUMORINO, "00000")
        
    End If
       
    GF_GetNextMitsumoriNo = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_GetNextMitsumoriNo", strSQL)
End Function

Public Function GF_GetShiyoKbn(ByVal strShiyoNo As String, ByVal strMsgTitle As String, _
                               ByRef strAndShiyuKbn As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:  �s�A�敪�擾
' �@�\�@�@�@:  �d�l�ݒ�No����s�A�敪���擾����
' �����@�@�@:  strShiyoNo      As String  �d�l�ݒ�No
'             strMsgTitle     As String  �G���[���b�Z�[�W�^�C�g��
'             strAndShiyuKbn  As String  �s�A�敪
' �@�\�����@:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strSQL     As String      ''SQL��
    Dim oraDyna    As OraDynaset  ''�޲ž��
    Dim intRet     As Integer
    
    GF_GetShiyoKbn = False
    
    strAndShiyuKbn = ""
    
    ''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT CIF.SHIYUKBN FROM THJMR MR,THJCIF CIF"
    strSQL = strSQL & "    WHERE MR.CIFNO = CIF.CIFNO"
    strSQL = strSQL & "      AND MR.EIGYONO = CIF.EIGYONO"
'2001/11/30 Added By Kigaku
    strSQL = strSQL & "      AND MR.SYNO = '" & strShiyoNo & "'"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = False Then
        strAndShiyuKbn = GF_VarToStr(oraDyna![SHIYUKBN])
        GF_GetShiyoKbn = True
        Exit Function
    End If
    
    '�޲ž�Ẳ��
    Set oraDyna = Nothing
    
'2001/12/27 Delete �C�O�͎g�p���Ȃ����߈ꎞ�I�ɍ폜
'    ''SQL��
'    strSQL = ""
'    strSQL = strSQL & "SELECT CIF.SHIYUKBN FROM THJYSMR MR,THJCIF CIF"
'    strSQL = strSQL & "    WHERE MR.CIFNO = CIF.CIFNO"
'    strSQL = strSQL & "      AND MR.EIGYONO = CIF.EIGYONO"
''2001/11/30 Added By Kigaku
'    strSQL = strSQL & "      AND MR.SYNO = '" & strShiyoNo & "'"
'
'    '�޲ž�Ă̐���
'    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
'
'    ''�ް���������
'    If oraDyna.EOF = False Then
'        strAndShiyuKbn = GF_VarToStr(oraDyna![SHIYUKBN])
'        GF_GetShiyoKbn = True
'        Exit Function
'    End If
'
'    '�޲ž�Ẳ��
'    Set oraDyna = Nothing
    
    If strAndShiyuKbn = "" Then
        intRet = GF_MsgBoxDB(strMsgTitle, "WTG001", "OK", "E")
        Exit Function
    End If

    GF_GetShiyoKbn = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_GetShiyoKbn", strSQL)

End Function


Public Function GF_GetShiyoKbn_CIF(ByVal strCifNO As String, ByVal strEigyoNo As String _
                                , ByRef strShiyuKbn As String, Optional ByVal strMsgTitle As String = "�s�A�敪�擾") As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:   �s�A�敪�擾
' �@�\�@�@�@:   �̔��X���ނƉc�Ə����ނ���s�A�敪���擾����
' �����@�@�@:  strCifNo As String   �̔��X����
'             strEigyoNo As String �c�Ə�����
'             strShiyuKbn  As String  �s�A�敪   '�s�A�敪 '1'��߰�:�����A'2'�C�O
'             strMsgTitle As String   ү��������
' �@�\�����@:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strSQL     As String      ''SQL��
    Dim oraDyna    As OraDynaset  ''�޲ž��
    Dim intRet     As Integer
    Dim strMsg     As String
    
    GF_GetShiyoKbn_CIF = False
    
    strShiyuKbn = ""
    
    ''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT SHIYUKBN FROM THJCIF"
    strSQL = strSQL & " WHERE CIFNO = '" & strCifNO & "'"
    strSQL = strSQL & "   AND EIGYONO = '" & strEigyoNo & "'"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = False Then
        strShiyuKbn = GF_VarToStr(oraDyna![SHIYUKBN])
        GF_GetShiyoKbn_CIF = True
        Exit Function
    Else
'        intRet = GF_MsgBoxDB(strMsgTitle, "WTH134", "OK", "E")
        strMsg = GF_GetMsg("WTH134")
        strMsg = strMsg & vbCr & "�̔��X���ށF" & strCifNO
        strMsg = strMsg & vbCr & "�c�Ə����ށF" & strEigyoNo
        intRet = GF_MsgBox(strMsgTitle, strMsg, "OK", "E")
        Exit Function
    End If
    
    '�޲ž�Ẳ��
    Set oraDyna = Nothing
    
    GF_GetShiyoKbn_CIF = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_GetShiyoKbn_CIF", strSQL)

End Function

Public Function GF_CutCharLength(ByRef strMoji As String, ByVal intCutLngth As Integer) As String
''------------------------------------------------------------------------------
'' @(f)
''
'' �@�\��    :   ��������w��o�C�g���Ő؂���
'' �@�\      :
'' ����      :   strMoji As String      (in/out) '������
''               intCutLngth As Integer (in)     '������
'' �߂�l    :   String     �؂�o������̕�����
'' ���l      :   strMoji �ɂ͐؂�o���������ȍ~�̕����񂪖߂�
''------------------------------------------------------------------------------
    Dim strDummy As String
    
    strDummy = LeftB(StrConv(strMoji, vbFromUnicode), intCutLngth + 1)
    If LenB(strDummy) > intCutLngth Then
        strDummy = StrConv(strDummy, vbUnicode)
        strMoji = Mid(strMoji, Len(strDummy))
        GF_CutCharLength = Left(strDummy, Len(strDummy) - 1)
    Else
        strMoji = Mid(strMoji, intCutLngth + 1)
        GF_CutCharLength = StrConv(strDummy, vbUnicode)
    End If
    
End Function


Public Function GF_ShowHelp(ByRef strLinkID As String) As Boolean
''------------------------------------------------------------------------------
'' @(f)
''
'' �@�\��    :�@�@�w���v��ʕ\��
'' �@�\      :
'' ����      :   strLinkID As String      (in/out) '��ʃ����NID
'' �߂�l    :
'' ���l      :
''------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim lngRet    As Long
    Dim blnErrFlg As Boolean
    
    GF_ShowHelp = False
    
    blnErrFlg = False
    
    On Error Resume Next
    
    '�w���v�N���v���O�����̗L���`�F�b�N
    If Dir(App.Path & "\help.exe") = "" Then
        blnErrFlg = True
    End If
    If Err.Number <> 0 Then
        blnErrFlg = True
    End If
    Err.Clear
    If blnErrFlg = True Then
        lngRet = GF_MsgBoxDB("�w���v", "WTG040", "OK", "E")
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    '�փ��v�̋N��
    lngRet = Shell(App.Path & "\help.exe " & strLinkID, vbNormalFocus)
    DoEvents
    
    GF_ShowHelp = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_ShowHelp")
End Function

Public Function GF_FileCopy(ByVal strSource As String _
                              , ByVal strDestination As String) As Boolean
''--------------------------------------------------------------------------------
'' @(f)
''
' �@�\���@�@:�@̧�ٺ�߰
' �@�\�@�@�@:�@̧�ق��߰����
' �����@�@�@:�@strSource        As String     ''��߰��̧�ٖ�
' �@�@�@�@�@ �@strDestination   As String     ''��߰��̧�ٖ�
'
'' �߂�l   : TRUE�F���� FALSE:�G���[ Boolean
''--------------------------------------------------------------------------------
On Error GoTo ErrHandler
    
    GF_FileCopy = False
    
    Call FileCopy(strSource, strDestination)
    
    GF_FileCopy = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_FileCopy", "��߰��:" & strSource & " , ��߰��:" & strDestination)
End Function

' 2018/05/15 �� T.Nakayama K545 CS�v���Z�X���P
Public Function GF_Chk_AutoModelWork(ByVal strAutoTypeFlg As String _
                                    , ByVal strTSDRNo As String _
                                    , ByVal strAcceptNo As String _
                                    , ByVal strDeliveryNo As String _
                                    , ByVal strHonkiAttKubun As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'�@�@�\���@: �����K�p�͈�ܰ���������
'�@�@�\�@�@:
'�@�����@�@: strAutoTypeFlg     As String      (in)  ��������׸�
'�@    �@�@: strTsdrNo          As String      (in)  �d�l�ݒ�NO
'�@    �@�@: strAcceptNo        As String      (in)  ��NO
'�@    �@�@: strDeliveryNo      As String      (in)  �����[���V�X�e��NO
'�@    �@�@: strHonkiAttKubun   As String      (in)  �{�@ATT�敪
'�@�߂�l�@: True = �L�� / False = ����
'�@���l�@�@:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL          As String
    Dim oraDyna         As OraDynaset
    '�Ώۃe�[�u��
    Dim strTaisyoTable  As String

    GF_Chk_AutoModelWork = False
    
    '�Ώۃe�[�u���̐ݒ�
    Select Case strAutoTypeFlg
    '�d�l�ݒ�̏ꍇ
    Case 1
        Select Case strHonkiAttKubun
        'ATT�̏ꍇ
        Case 1
            '�Ώۃe�[�u��:�����K�p�͈�ܰ�(�d�l�ݒ�)(����)(ATT)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_THJ_ATT"
        '�{�@�̏ꍇ
        Case 2
            '�Ώۃe�[�u��:�����K�p�͈�ܰ�(�d�l�ݒ�)(����)(�{�@)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_THJ_H"
        End Select
    '�����[���̏ꍇ
    Case 2
        Select Case strHonkiAttKubun
        'ATT�̏ꍇ
        Case 1
            '�Ώۃe�[�u��:�����K�p�͈�ܰ�(�����[��)(����)(ATT)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_INQ_ATT"
        '�{�@�̏ꍇ
        Case 2
            '�Ώۃe�[�u��:�����K�p�͈�ܰ�(�����[��)(����)(�{�@)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_INQ_H"
        End Select
    '�d�������̏ꍇ
    Case 3
        Select Case strHonkiAttKubun
        'ATT�̏ꍇ
        Case 1
            '�Ώۃe�[�u��:�����K�p�͈�ܰ�(����)(ATT)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_ATT"
        '�{�@�̏ꍇ
        Case 2
            '�Ώۃe�[�u��:�����K�p�͈�ܰ�(����)(�{�@)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_H"
        End Select
    End Select

    '��NO
    strAcceptNo = IIf(strAcceptNo <> "", strAcceptNo, "            ")
    '�����[���V�X�e��NO
    strDeliveryNo = IIf(strDeliveryNo <> "", strDeliveryNo, "          ")

    strSQL = ""
    strSQL = strSQL & " SELECT VCTSDR_NO"
    strSQL = strSQL & strTaisyoTable
    strSQL = strSQL & "  WHERE VCTSDR_NO = '" & Trim(strTSDRNo) & "'"
    strSQL = strSQL & "    AND CACCEPTNO = '" & strAcceptNo & "'"
    strSQL = strSQL & "    AND NDELIVERY_NO = '" & strDeliveryNo & "'"

    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oraDyna.EOF = False Then
        GF_Chk_AutoModelWork = True
    End If
    Set oraDyna = Nothing

    Exit Function
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_Chk_AutoModelWork", strSQL)
End Function
' 2018/05/15 �� T.Nakayama K545 CS�v���Z�X���P
