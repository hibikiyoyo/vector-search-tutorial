Attribute VB_Name = "basMainFunc"
' @(h) MainFunc.bas  ver1.00 ( 2000/08/30 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   �v���W�F�N�g��  : TLF��ۼު��
'   ���W���[����    : basMainFunc
'   �t�@�C����      : MainFunc.bas
'   Version         : 1.00
'   �@�\����       �F EXE�N�����̏��������Ɋւ��鋤�ʊ֐�
'   �쐬��         �F T.Fukutani
'   �쐬��         �F 2000/08/30
'   �C������       �F 2000/12/22 T.Fukutani հ�ޔF�ؕ��@�ύX etc.
'                  �F 2003/04/07 N.Kigaku   GF_CreateArgument�ALF_GetCommndLine �����C��
'                  �F 2005/07/27 N.Kigaku LF_GetCommndLine �����ƭ��N���Ή�
'                  �F 2007/02/26 N.Kigaku GF_GetConnectUserName�֐��ǉ�
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ���錾
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' �p�u���b�N�萔�錾
'------------------------------------------------------------------------------
Public Const gINIFILE = "Order.ini"          ''ini̧�ق̎w��

'------------------------------------------------------------------------------
' ���W���[���萔�錾
'------------------------------------------------------------------------------
Private Const TLF_BYCOMMAND = &H0&
Private Const TLF_BYPOSITION = &H400&
Private Const SC_CLOSE = &HF060

'Private Const gstrDBUser = "LFUSR"           ''DB�ڑ�հ��
'Private Const mstrDBPwd = "LFUSR"            ''DB�ڑ��߽ܰ��

'------------------------------------------------------------------------------
' �p�u���b�N�ϐ��錾
'------------------------------------------------------------------------------
Public gstrUserID       As String           ''հ��ID
Public gstrArgument     As String           ''EXE���Ƃ̈���

'Public gstrDBUser      As String           ''DB�ڑ�հ��
'Public gstrDBPwd       As String           ''DB�ڑ��߽ܰ��
'Public gstrDBInstance  As String           ''DB��

'------------------------------------------------------------------------------
' ���W���[���ϐ��錾
'------------------------------------------------------------------------------
'Private gclsUserInfo    As New clsUserInfo
Private mstrUserPWD     As String           ''հ���߽ܰ��
Private mstrDBUser      As String           ''DB�ڑ�հ��
Private mstrDBPwd       As String           ''DB�ڑ��߽ܰ��
Private mstrDBInstance  As String           ''DB��
Private mstrTVL_Log_Dir As String           ''TIVOLI����۸ޏo�͐�
Private mblnTVL_Log_Flg As Boolean          ''TIVOLI����۸ޏo���׸� [False:�o�͂��Ȃ��ATrue:�o�͂��� ]

'------------------------------------------------------------------------------
' �O���v���V�[�W���̃v���g�^�C�v�錾
'------------------------------------------------------------------------------
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long ''ini̧�ق̓Ǐo��
Declare Function GetSystemMenu Lib "USER32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "USER32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

' ���O�I�����[�U�[�����擾����API
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Property Get Password() As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:�߽ܰ�ގ擾�p�����è
' �@�\�@�@�@:
' �����@�@�@:�Ȃ�
' �߂�l�@�@:�߽ܰ��
' �@�\�����@:Login��ʂɂē��͂��ꂽ�߽ܰ�ނ�Ԃ�
'------------------------------------------------------------------------------
    Password = mstrDBPwd
End Property

Public Property Get TVL_LOG_FLG() As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:TIVOLI۸ޏo���׸ގ擾�p�����è
' �@�\�@�@�@:
' �����@�@�@:�Ȃ�
' �߂�l�@�@:True / False
' �@�\�����@:
'------------------------------------------------------------------------------
    TVL_LOG_FLG = mblnTVL_Log_Flg
End Property

Public Property Get TVL_LOG_DIR() As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:TIVOLI۸ޏo�͐�擾�p�����è
' �@�\�@�@�@:
' �����@�@�@:�Ȃ�
' �߂�l�@�@:TIVOLI۸ޏo�͐�
' �@�\�����@:
'------------------------------------------------------------------------------
    TVL_LOG_DIR = mstrTVL_Log_Dir
End Property


Public Function GF_Initialize(Optional bolAuthenticationFlg As Boolean = True, _
                              Optional nCountFlg As Integer = 0, _
                              Optional blnShowLoginFlg As Boolean = True, _
                              Optional blnWOraSessionConnectFlag As Boolean = True) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : ۸޲ݏ���
' �@�\   : DB�֐ڑ��Aհ�ޔF�؂��s��
' ����   : bolAuthenticationFlg As Boolean ''հ�ޔF���׸�(TRUE:�s���AFALSE:�s��Ȃ�)
'          nCountFlg As Integer            ''�ċA�̉�(�ʏ�͏ȗ�)
'          blnShowLoginFlg As Boolean      ''۸޲݉�ʕ\���׸�(TRUE:�\���AFALSE:��\��)
'          blnWOraSessionConnectFlag As Boolean    ''۸ޏo�͗pOracleDB�ڑ��׸�(TRUE:�ڑ��AFALSE:��ڑ�)
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strUser_Kbn    As String        ''հ�ދ敪
    Dim strSoshiki_Kbn As String        ''�g�D�敪
    Dim nCount         As Integer       ''�w�蕶���̈ʒu
    Dim strErrType     As String        ''�װ����
    Dim intRet         As Integer       ''�Ԃ�l
    Dim strNetUser     As String
    Dim strNetpass     As String
'    Dim strPass        As String
    Dim strPath        As String
    Dim strMsg         As String
'    Static stnCount    As Integer       ''����
    
    GF_Initialize = False
    
    '��d�N������
    If App.PrevInstance = True Then Exit Function

'    ''�ċA�񐔁{�P
'    stnCount = stnCount + 1
            
    ''INI̧�ّ�������
    If Dir(App.Path & "\" & gINIFILE) = "" Then
'        stnCount = 3
        strErrType = "INI"
        Err.Raise Number:=vbObjectError, Description:=gINIFILE & " �����݂��܂���B"
    End If
    
    ''�����ײ݈�������
    If (nCountFlg > 0 Or LF_GetCommndLine = False) And (blnShowLoginFlg = True) Then
    
        ''հ�ޖ��̎擾
        If LF_GetConnectUserName(gstrUserID) = False Then Exit Function
    
        '''۸޲݉�ʌďo��
        '��ݾَ��A�I��
        If Len(Trim(gstrUserID)) = 0 Then
            If LF_LogIn = False Then Exit Function
        End If
    End If
    
    'հ��ID�ݒ�
    basMsgFunc.UserID = gstrUserID
    '�[��CD�ݒ�
    basMsgFunc.TerminalCD = Environ("COMPUTERNAME")
    

'    ''INI̧�ٓǍ���
'
'    ''ȯ�ܰ��ڑ������p�߽�擾
'    strPath = GF_ReadINI("SERVER", "CHECK_PATH")
'    If strPath <> "" Then
'        ''ȯ�ܰ��ڑ������p�߽����
'        If GF_NetWorkShareCheck(strPath) < 0 Then
'            ''ȯ�ܰ��ڑ�հ�ގ擾
'            strNetUser = GF_ReadINI("SERVER", "USER")
'            ''ȯ�ܰ��ڑ��߽ܰ�ގ擾
'            strNetpass = GF_ReadINI("SERVER", "PASS")
'
'            'ȯ�ܰ��ڑ�
'            If GF_NetConnect(strNetUser, strNetpass, strPath) = False Then
'                ''�m�Fү���ޕ\��
'                strMsg = "���̂܂܏����𑱍s����ƈ���������g�p�ł��Ȃ��\��������܂��B" & vbCr & "���s���܂����H"
'                intRet = GF_MsgBox("NetWork", strMsg, "OC", "Q")
'                If intRet = 0 Or intRet = vbCancel Then
'                    '''��ݾى�����
'                    Exit Function
'                End If
'            End If
'        End If
'    End If
    
    '''DB���擾
    mstrDBInstance = GF_ReadINI("ORACLE", "DSN")
    If mstrDBInstance = "" Then
'        stnCount = 3
        strErrType = "INI"
        Err.Raise Number:=vbObjectError, Description:="DB���� " & gINIFILE & " �ɐ������ݒ肳��Ă��܂���B"
    End If

    '''DB�ڑ�հ�ޖ��擾
    mstrDBUser = GF_ReadINI("ORACLE", "USERNAME")
    If mstrDBUser = "" Then
'        stnCount = 3
        strErrType = "INI"
        Err.Raise Number:=vbObjectError, Description:="DB�ڑ����[�U���� " & gINIFILE & " �ɐ������ݒ肳��Ă��܂���B"
    End If
    
    '''DB�ڑ��߽ܰ�ގ擾
    mstrDBPwd = GF_ReadINI("ORACLE", "PASSWORD")
    If mstrDBPwd = "" Then
'        stnCount = 3
        strErrType = "INI"
        Err.Raise Number:=vbObjectError, Description:="DB�ڑ����[�U���� " & gINIFILE & " �ɐ������ݒ肳��Ă��܂���B"
    End If
    
    '�O���[�o�����[�U���ϐ��Ɋi�[����
'    gstrUserID = mstrDBUser
    
    ''DB�ڑ��߽ܰ�ގ擾
'    If bolAuthenticationFlg = True Then
'        strPass = mstrDBPwd
'    Else
'        strPass = mstrUserPWD
'    End If
    
    'TIVOLI����۸ޏo�͑Ώ���۸��є���
    If LF_Read_Tivoli_Log_PGM = False Then Exit Function
    
    ''DB�ڑ�(DB�ڑ�հ�ށADB�ڑ��߽ܰ�ނŐڑ�)
    If GF_DBOpen(mstrDBInstance, mstrDBUser, mstrDBPwd, blnWOraSessionConnectFlag) = False Then Exit Function
    
    If bolAuthenticationFlg = True Then
'        ''հ�ޔF��
'        If LF_UserAuthentication = False Then Exit Function
    End If
    
    GF_Initialize = True
    
    Exit Function
    
ErrHandler:
    Dim strLocation    As String    ''�װ�����ꏊ
    Dim lngErrNum      As Long      ''�װ���ް
    Dim strErrMsg      As String    ''�װү����
    
    strLocation = "GF_Initialize"
    strErrMsg = Err.Description
    
    ''�װү���ޕ\��
    If basMsgFunc.DispErrMsgFlg = True Then
        intRet = GF_MsgBox("Login", strErrMsg, "OK", "E")
    End If
    ''�װ۸ޏo��
    intRet = GF_LogOut(strErrType, strLocation, "", Replace(strErrMsg, vbCr, ""))
    
'    ''3��ԈႦ����I��
'    If stnCount <= 2 Then
'        '''۸޲ݏ���(�ċA)
'        If GF_Initialize(stnCount) = False Then Exit Function
'        '''۸޲ݏ����A����I��
'        GF_Initialize = True
'    End If
    
End Function

Public Function LF_Read_Tivoli_Log_PGM() As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : TIVOLI����۸ޏo�͑Ώ���۸��є���
' �@�\   :
' ����   :
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    Dim strPGM      As String
    Dim strErrMsg   As String
    Dim i           As Integer
    
    On Error GoTo ErrHandler
    
    LF_Read_Tivoli_Log_PGM = False
    
    mblnTVL_Log_Flg = False
    
    'TIVOLI�T�[�o���O�f�B���N�g�����擾
    mstrTVL_Log_Dir = GF_ReadINI("TIVOLI_LOG", "TVL_ERR_LOG")
    
    'TIVOLI�T�[�o���O�o�͑Ώۃv���O�������擾
    For i = 1 To 99
        strPGM = GF_ReadINI("TIVOLI_LOG", "TVL_LOG_EXE_" & Format(i, "00"))
        If Len(Trim(strPGM)) = 0 Then Exit For
        
        If StrComp(App.EXEName, strPGM, vbTextCompare) = 0 Then
            mblnTVL_Log_Flg = True
            Exit For
        End If
    Next i
    
    '�o�͐�`�F�b�N
    If (mblnTVL_Log_Flg = True) And (Len(Trim(mstrTVL_Log_Dir)) = 0) Then
        Err.Raise Number:=vbObjectError, Description:="�o�͐悪�w�肳��Ă��܂���B"
    End If
    
    LF_Read_Tivoli_Log_PGM = True
    Exit Function
    
ErrHandler:
    strErrMsg = Err.Description
    ''�װү���ޕ\��
    If basMsgFunc.DispErrMsgFlg = True Then
        Call GF_MsgBox("Login", "TIVOLI����۸ޏo�͑Ώ���۸��т̎擾�Ɏ��s���܂����B" & vbCrLf & strErrMsg, "OK", "E")
    End If
    ''�װ۸ޏo��
    Call GF_LogOut("VB", "LF_GetConnectUserName", "", "TIVOLI����۸ޏo�͑Ώ���۸��т̎擾�Ɏ��s���܂����B  " & strErrMsg)
End Function

Private Function LF_GetConnectUserName(ByRef strUserName As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : հ�ޖ��擾
' �@�\   :
' ����   : strUserName  As String (out)      հ�ޖ�
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    Dim strDummy As String
    Dim strErrMsg As String
    
On Error GoTo ErrHandler

    LF_GetConnectUserName = False
    
    strDummy = String(256, Chr(0))
    ''հ�ޖ��̎擾
    Call GetUserName(strDummy, 256)
    strDummy = Mid(strDummy, 1, InStr(1, strDummy, Chr(0), vbTextCompare) - 1)
    strUserName = Trim(strDummy)
        
    LF_GetConnectUserName = True
        
    Exit Function
ErrHandler:
    strErrMsg = Err.Description
    ''�װү���ޕ\��
    If basMsgFunc.DispErrMsgFlg = True Then
        Call GF_MsgBox("Login", "�ڑ����[�U�̎擾�Ɏ��s���܂����B" & vbCrLf & strErrMsg, "OK", "E")
    End If
    ''�װ۸ޏo��
    Call GF_LogOut("VB", "LF_GetConnectUserName", "", "�ڑ����[�U�̎擾�Ɏ��s���܂����B  " & strErrMsg)
End Function

'2007/02/26 Added by N.Kigaku
Public Function GF_GetConnectUserName(ByRef strUserName As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : հ�ޖ��擾(��۰��ٔ�)
' �@�\   :
' ����   : strUserName  As String (out)      հ�ޖ�
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    GF_GetConnectUserName = LF_GetConnectUserName(strUserName)
End Function

Private Function LF_GetCommndLine() As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �����ײ݈����擾
' �@�\   :
' ����   :
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strCmdLine As String
    Dim intPosition As String
    
    LF_GetCommndLine = False
    
    strCmdLine = ""
    strCmdLine = Trim(Command())
    
    '�R�}���h���C���������Ȃ��� EXIT
    If Len(strCmdLine) = 0 Then Exit Function
    
'    If InStr(1, strCmdLine, "/", vbTextCompare) <> 0 Then

    intPosition = InStr(1, strCmdLine, " ", vbTextCompare)
    If intPosition <> 0 Then
    
        ''հ��ID�擾
        gstrUserID = Trim(Left(strCmdLine, InStr(1, strCmdLine, " ", vbTextCompare) - 1))
'2005/07/27 Added by N.Kigaku
''�����ƭ��N���Ή��@հ�ޖ���"/"������ꍇ�͍폜����
        If InStr(1, gstrUserID, "/", vbTextCompare) <> 0 Then
        
            gstrUserID = Trim(Left(strCmdLine, InStr(1, gstrUserID, "/", vbTextCompare) - 1))
        End If
        
        ''հ���߽ܰ��
        mstrUserPWD = ""

        ''EXE���Ƃ̈����擾
        gstrArgument = Trim(Mid(strCmdLine, InStr(1, strCmdLine, " ", vbTextCompare) + 1))
        
        ''հ��ID�擾
'        gstrUserID = Trim(Left(strCmdLine, InStr(1, strCmdLine, "/", vbTextCompare) - 1))
        ''DB�ڑ��߽ܰ�ގ擾
    '    mstrDBPwd = Trim(Mid(strCmdLine, InStr(1, strCmdLine, "/", vbTextCompare) + 1))
        ''հ���߽ܰ�ގ擾
'        If InStr(1, strCmdLine, " ", vbTextCompare) = 0 Then
'            ''EXE���Ƃ̈������Ȃ��ꍇ
'            mstrUserPWD = Trim(Mid(strCmdLine, InStr(1, strCmdLine, "/", vbTextCompare) + 1))
'        Else
'            '''EXE���Ƃ̈���������ꍇ
'            mstrUserPWD = Trim(Mid(strCmdLine, InStr(1, strCmdLine, "/", vbTextCompare) + 1, InStr(1, strCmdLine, " ", vbTextCompare) - 1 - InStr(1, strCmdLine, "/", vbTextCompare)))
            ''EXE���Ƃ̈����擾
'            gstrArgument = Trim(Mid(strCmdLine, InStr(1, strCmdLine, " ", vbTextCompare) + 1))
'        End If
    Else
        ''հ��ID�擾
        gstrUserID = Trim(strCmdLine)
        ''�߽ܰ��
        mstrUserPWD = ""
    End If
    
    LF_GetCommndLine = True
    
    Exit Function
      
ErrHandler:
    Dim strLocation    As String    ''�װ�����ꏊ
    Dim lngErrNum      As Long      ''�װ���ް
    Dim strErrMsg      As String    ''�װү����
    Dim strErrType     As String    ''�װ����
    Dim intRet         As Integer   ''�Ԃ�l
    
    strLocation = "LF_GetCommndLine"
    strErrMsg = "�R�}���h���C���������Ⴂ�܂��B"
    strErrType = "INI"
    
    ''�װү���ޕ\��
    If basMsgFunc.DispErrMsgFlg = True Then
        intRet = GF_MsgBox("Login", strErrMsg, "OK", "E")
    End If
    ''�װ۸ޏo��
    intRet = GF_LogOut(strErrType, strLocation, "", strErrMsg)
    
End Function

Private Function LF_LogIn(Optional intLoginKBN As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : Login
' �@�\   : Login��ʌďo��
' ����   : intLoginKBN As Integer     ''۸޲݋敪 0:�]�ƈ��ԍ�, 1:�߽ܰ��
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    Dim bolCancel As Boolean
    
    LF_LogIn = False
    
    Screen.MousePointer = vbDefault
    
    DoEvents
    
    ''۸޲݉�ʌďo��
    Screen.MousePointer = vbDefault
    frmLogin.LoginKBN = intLoginKBN
    frmLogin.Show vbModal
    Screen.MousePointer = vbHourglass
    
    '�߽ܰ�ގ擾
'    mstrDBPwd = frmLogin.Password
    'mstrUserPWD = frmLogin.Password
    '��ݾ��׸ގ擾
    bolCancel = frmLogin.CancelFlg
    
    '۸޲݉�ʱ�۰��
    Unload frmLogin
    
    DoEvents
    
    'Login��ʂŷ�ݾق��ꂽ���H
    If bolCancel = True Then Exit Function
                     
    LF_LogIn = True
    
End Function

Private Function LF_UserAuthentication() As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : հ�ޔF��
' �@�\   : ۸޲݉�ʂɂē��͂��ꂽհ�ނ̔F�؂��s��
' ����   :
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strErrType   As String
    Dim strMsgID   As String
    Dim strSQL       As String         ''SQL��
    Dim oraDyna      As OraDynaset     ''�޲ž��
    Dim bolRetFlg    As Boolean        ''�׸�
    Dim intRet       As Integer
    Dim strPassWord  As String
    
    ''հ�ޔF���׸ޏ�����
    LF_UserAuthentication = False
    
    'SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM THJUSRMR"
    strSQL = strSQL & " WHERE SYAINCD = '" & UCase(Trim(gstrUserID)) & "'"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = True Then
        ''''�Y���ް��Ȃ�
        strErrType = "Login"
        '�Ј�Ͻ��ɓo�^����Ă��܂���B
        strMsgID = "WTG008"
        intRet = GF_MsgBoxDB(strErrType, strMsgID, "OK", "E")
        intRet = GF_LogOutDB(strErrType, "mMLR_UserAuthentication", strMsgID)
        Exit Function
    End If
    
'    strPassWord = GF_VarToStr(oraDyna![Password])
'    If UCase(mstrUserPWD) <> strPassWord Then
'        '''�߽ܰ�ފԈႢ
'        strErrType = "Login"
'        '�߽ܰ�ނ��Ⴂ�܂��B
'        strMsgID = "WTG002"
'        intRet = GF_MsgBoxDB(strErrType, strMsgID, "OK", "E")
'        intRet = GF_LogOutDB(strErrType, "mMLR_UserAuthentication", strMsgID)
'        Exit Function
'    End If
    
    '�޲ž�ĊJ��
    Set oraDyna = Nothing
    
    LF_UserAuthentication = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("LF_UserAuthentication", strSQL)

End Function

Public Function GF_CreateArgument() As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : հ�ޏ������쐬
' �@�\   : EXE�ɓn��հ�ޏ��������쐬����
' ����   :
' �߂�l : String     հ�ޏ�����(հ��ID/հ���߽ܰ�� EXE���Ƃ̈���)
' ���l   :
'------------------------------------------------------------------------------
'    GF_CreateArgument = gstrUserID & "/" & mstrUserPWD
    GF_CreateArgument = gstrUserID
End Function

Public Sub GS_DelControlBox(frmForm As Form)
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : �������ݎg�p�s��
' �@�\   : �������݂��g�p�s�ɂ���
' ����   : frmForm As Form      ''̫�ѵ�޼ު��
' ���l   :
'------------------------------------------------------------------------------
    Dim lngSysMenu As Long
    Dim intRet     As Integer
    
    lngSysMenu = GetSystemMenu(frmForm.hwnd, 0)
    
    intRet = DeleteMenu(lngSysMenu, 5, TLF_BYPOSITION)
    intRet = DeleteMenu(lngSysMenu, SC_CLOSE, TLF_BYCOMMAND)

End Sub

Public Function GF_ReadINI(strSection As String, strKey As String) As String
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : INI̧�ق�ǂݏo��
' �@�\   : INI̧�ق��w��̾���݁A�w���Key�̒l���擾����
' ����   : strSection As String   ''�����
'       �F strKey     As String   ''Key
' �߂�l : �w�肵��Key�̒l
' ���l   :
'------------------------------------------------------------------------------

    Dim lngRet  As Long            ''GetPrivateProfileString�̖߂�l�@0�F�װ
    Dim strBuff As String * 256

    GF_ReadINI = ""
    lngRet = GetPrivateProfileString(strSection, strKey, "", _
                                        strBuff, 255, App.Path & "\" & gINIFILE)

    ''�����񐔂�"0"�̎��A�G���[
    If lngRet <> 0 Then
        GF_ReadINI = strConv(MidB(strConv(strBuff, vbFromUnicode), 1, lngRet), vbUnicode)
    End If

End Function
