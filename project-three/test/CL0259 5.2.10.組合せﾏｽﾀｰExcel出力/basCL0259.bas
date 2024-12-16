Attribute VB_Name = "basCL0259"
' @(h) basCL0259.bas  ver1.0.0.1 ( 2004/10/20 J.Hamaji )
'------------------------------------------------------------------------------
' @(s)
'   �v���W�F�N�g��  : TLF��ۼު��
'   ���W���[����    : basCL0259
'   �t�@�C����      : basCL0259.bas
'   Version         : 1.0.0.1
'   �@�\����       �F �g����Ͻ��Excel�o��
'   �쐬��         �F J.Hamaji
'   �쐬��         �F 2004/10/20
'   �C������       �F 2004/10/28 THS T.Y (�ײ��Ăɒ���Excel�o�͂���)
'   �@�@�@�@       �F 2005/07/07 THS J.Yamaoka (�㏑��ү���ޕt������Ver�ɕύX)
'   �@�@�@�@       �F 2006/04/13 THS Sugawara Ver1.0.0.1  6�����ȏ�̃f�[�^�Ŏ��̃V�[�g�ɏo�͂���Ȃ��s����C��
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ���錾
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ���W���[���ϐ��錾
'------------------------------------------------------------------------------

Public iCount As Integer
Public strMsg() As String

Public gstrServerPath As String     'EXCEL�o�͐�̃p�X(�T�[�o)
Public gstrClientPath As String     'EXCEL�o�͐�̃p�X(�N���C�A���g)
Public gstrFileName   As String     'EXCEL�t�@�C����

'------------------------------------------------------------------------------
' ���W���[���萔�錾
'------------------------------------------------------------------------------
Public Const cProtectColor As Long = &H8000000F
Public Const cNoProtectColor As Long = &H80000005

Public Const gaOK           As Integer = 0
Public Const gaNG           As Integer = -1
Public Const gaNothing      As Integer = 1
Public Const strEndDAte     As String = "99999999"
Public Const strProAlarm    As String = "0"         '���ǃA���[���t���O
Public Const strKumiKbn     As String = "TGLOBAL_SET_CS"

Private Const mstrPGMID As String = "CL0259"        '�v���O����ID

Private Sub Main()
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:Main����
' �@�\�@�@�@:�T�u�V�X�e����������
' �����@�@�@:
' �@�\�����@:�T�u�V�X�e����ʕ\����������
'------------------------------------------------------------------------------
    On Error GoTo Err_Main
    
    Dim bolRes As Boolean
    
    '�}�E�X�|�C���^�ݒ�(�����v)
    Screen.MousePointer = vbHourglass
    
    '�����N���̗}�~
    If App.PrevInstance = True Then
        MsgBox "���łɋN������Ă��܂��B", vbExclamation
        Exit Sub
    End If
    
    'Main����������
    If LF_Main_Initialize = False Then
        '�}�E�X�|�C���^�ݒ�
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    '��������
    If GF_Initialize() = False Then
        '�}�E�X�|�C���^�ݒ�
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    '�w�g����Ͻ��Excel�o�́x��ʕ\��
    frmCL0259.Show
    
    '�t�H�[�����[�h��������
    If (frmCL0259.LoadFlag = False) Then
        '�t�H�[���A�����[�h
        Unload frmCL0259
        'DB�ؒf
        If Not (gOraParam Is Nothing) Or _
           Not (gOraDataBase Is Nothing) Or _
           Not (gOraSession Is Nothing) Then
    
            Call GS_DBClose
    
        End If
    Else
        frmCL0259.cmbExportCs.SetFocus
    End If
    
    'ini�t�@�C���̊��ݒ�����擾
    Call LF_GetiniInf
    
    '�}�E�X�|�C���^�ݒ�
    Screen.MousePointer = vbDefault

Exit_Main:
    Exit Sub

Err_Main:
    'Main���̎��s���G���[����
    Call GS_ErrorHandler("Main")
    Resume Exit_Main
End Sub

Private Function LF_Main_Initialize() As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : Main�������̏�����
' �@�\   :
' ����   :
' �߂�l :
' ���l   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    '�װү���ޕ\���׸ސݒ�
    basMsgFunc.DispErrMsgFlg = False
    
    '��۸��і��ݒ�
    basMsgFunc.PGMCD = mstrPGMID
    
    '۸�̧�ٖ��ݒ�
    basMsgFunc.LogFile = GF_ReadINI("DIR", "ERR_LOG")
    
    If Len(Trim(basMsgFunc.LogFile)) = 0 Then Exit Function
    
    basMsgFunc.LogFile = basMsgFunc.LogFile & mstrPGMID & "_" & Format(Date, "YYYYMMDD") & ".LOG"

    LF_Main_Initialize = True
    
    Exit Function
    
ErrHandler:
    LF_Main_Initialize = False
    
End Function

Private Function LF_GetiniInf() As Boolean
'--------------------------------------------------------------------------------
' @(f)
' �@�\��    : ini�t�@�C���̊��ݒ�����擾(�T�[�o/�N���C�A���g)
' �@�\      :
' ����      :
' �߂�l    : TRUE�F���� FALSE:�G���[ Boolean
' �@�\�����@:
'--------------------------------------------------------------------------------
On Error GoTo ErrHandler
        
    LF_GetiniInf = False
         
    ''DEL 2004/10/28 THS T.Y (�ײ��Ăɒ���Excel�o�͂���) START>>>>>
'''''    'ini�t�@�C����EXCEL�o�͐�t�H���_�擾(�T�[�o)
'''''    gstrServerPath = GF_ReadINI("MASTER", "MST_EXPORT_DIR")
'''''
'''''    'EXCEL�o�͐�t�H���_�擾�̔���(�T�[�o)
'''''    If gstrServerPath = "" Then
'''''        '�o�͐�t�H���_���擾���s
'''''        '���O�o��
'''''        Call GF_GetMsg_Addition("WTK398", , False, True)
'''''        'MSG�\��
'''''        Call GF_MsgBoxDB(frmCL0259.Caption, "WTK398", "OK", "C")
'''''        Exit Function
'''''    '�p�X���̍Ō����"\"�����Ă��邩
'''''    ElseIf Right(gstrServerPath, 1) <> "\" Then
'''''        gstrServerPath = gstrServerPath & "\"
'''''    End If
    ''<<<<<END
    
    'ini�t�@�C����EXCEL�o�͐�t�H���_�擾(�N���C�A���g)
    gstrClientPath = GF_ReadINI("MASTER", "MST_I/O_DIR")
    
    'EXCEL�o�͐�t�H���_�擾�̔���(�N���C�A���g)
    If gstrClientPath = "" Then
        '�o�͐�t�H���_���擾���s
        '���O�o��
        Call GF_GetMsg_Addition("WTK399", , False, True)
        'MSG�\��
        Call GF_MsgBoxDB(frmCL0259.Caption, "WTK399", "OK", "C")
        Exit Function
    '�p�X���̍Ō����"\"�����Ă��邩
    ElseIf Right(gstrClientPath, 1) <> "\" Then
        gstrClientPath = gstrClientPath & "\"
    End If
    
    'ini�t�@�C����EXCEL�t�@�C�����̎擾
    gstrFileName = GF_ReadINI(mstrPGMID, "OUTPUT_EXCEL_FILE")
    
    'EXCEL�t�@�C�����擾�̔���
    If gstrFileName = "" Then
        'EXCEL�t�@�C�����擾���s
        '���O�o��
        Call GF_GetMsg_Addition("WTK400", , False, True)
        'MSG�\��
        Call GF_MsgBoxDB(frmCL0259.Caption, "WTK400", "OK", "C")
        Exit Function
    End If
    
    LF_GetiniInf = True
    
    Exit Function

ErrHandler:

    Call GS_ErrorHandler("LF_GetiniInf")

End Function

