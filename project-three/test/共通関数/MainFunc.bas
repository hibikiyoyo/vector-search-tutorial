Attribute VB_Name = "basMainFunc"
' @(h) MainFunc.bas  ver1.00 ( 2000/08/30 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   プロジェクト名  : TLFﾌﾟﾛｼﾞｪｸﾄ
'   モジュール名    : basMainFunc
'   ファイル名      : MainFunc.bas
'   Version         : 1.00
'   機能説明       ： EXE起動時の初期処理に関する共通関数
'   作成者         ： T.Fukutani
'   作成日         ： 2000/08/30
'   修正履歴       ： 2000/12/22 T.Fukutani ﾕｰｻﾞ認証方法変更 etc.
'                  ： 2003/04/07 N.Kigaku   GF_CreateArgument、LF_GetCommndLine 引数修正
'                  ： 2005/07/27 N.Kigaku LF_GetCommndLine 引合ﾒﾆｭｰ起動対応
'                  ： 2007/02/26 N.Kigaku GF_GetConnectUserName関数追加
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 環境宣言
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' パブリック定数宣言
'------------------------------------------------------------------------------
Public Const gINIFILE = "Order.ini"          ''iniﾌｧｲﾙの指定

'------------------------------------------------------------------------------
' モジュール定数宣言
'------------------------------------------------------------------------------
Private Const TLF_BYCOMMAND = &H0&
Private Const TLF_BYPOSITION = &H400&
Private Const SC_CLOSE = &HF060

'Private Const gstrDBUser = "LFUSR"           ''DB接続ﾕｰｻﾞ
'Private Const mstrDBPwd = "LFUSR"            ''DB接続ﾊﾟｽﾜｰﾄﾞ

'------------------------------------------------------------------------------
' パブリック変数宣言
'------------------------------------------------------------------------------
Public gstrUserID       As String           ''ﾕｰｻﾞID
Public gstrArgument     As String           ''EXEごとの引数

'Public gstrDBUser      As String           ''DB接続ﾕｰｻﾞ
'Public gstrDBPwd       As String           ''DB接続ﾊﾟｽﾜｰﾄﾞ
'Public gstrDBInstance  As String           ''DB名

'------------------------------------------------------------------------------
' モジュール変数宣言
'------------------------------------------------------------------------------
'Private gclsUserInfo    As New clsUserInfo
Private mstrUserPWD     As String           ''ﾕｰｻﾞﾊﾟｽﾜｰﾄﾞ
Private mstrDBUser      As String           ''DB接続ﾕｰｻﾞ
Private mstrDBPwd       As String           ''DB接続ﾊﾟｽﾜｰﾄﾞ
Private mstrDBInstance  As String           ''DB名
Private mstrTVL_Log_Dir As String           ''TIVOLIｻｰﾊﾞﾛｸﾞ出力先
Private mblnTVL_Log_Flg As Boolean          ''TIVOLIｻｰﾊﾞﾛｸﾞ出力ﾌﾗｸﾞ [False:出力しない、True:出力する ]

'------------------------------------------------------------------------------
' 外部プロシージャのプロトタイプ宣言
'------------------------------------------------------------------------------
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long ''iniﾌｧｲﾙの読出し
Declare Function GetSystemMenu Lib "USER32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "USER32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

' ログオンユーザー名を取得するAPI
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Property Get Password() As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:ﾊﾟｽﾜｰﾄﾞ取得用ﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　:なし
' 戻り値　　:ﾊﾟｽﾜｰﾄﾞ
' 機能説明　:Login画面にて入力されたﾊﾟｽﾜｰﾄﾞを返す
'------------------------------------------------------------------------------
    Password = mstrDBPwd
End Property

Public Property Get TVL_LOG_FLG() As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:TIVOLIﾛｸﾞ出力ﾌﾗｸﾞ取得用ﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　:なし
' 戻り値　　:True / False
' 機能説明　:
'------------------------------------------------------------------------------
    TVL_LOG_FLG = mblnTVL_Log_Flg
End Property

Public Property Get TVL_LOG_DIR() As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:TIVOLIﾛｸﾞ出力先取得用ﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　:なし
' 戻り値　　:TIVOLIﾛｸﾞ出力先
' 機能説明　:
'------------------------------------------------------------------------------
    TVL_LOG_DIR = mstrTVL_Log_Dir
End Property


Public Function GF_Initialize(Optional bolAuthenticationFlg As Boolean = True, _
                              Optional nCountFlg As Integer = 0, _
                              Optional blnShowLoginFlg As Boolean = True, _
                              Optional blnWOraSessionConnectFlag As Boolean = True) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : ﾛｸﾞｲﾝ処理
' 機能   : DBへ接続、ﾕｰｻﾞ認証を行う
' 引数   : bolAuthenticationFlg As Boolean ''ﾕｰｻﾞ認証ﾌﾗｸﾞ(TRUE:行う、FALSE:行わない)
'          nCountFlg As Integer            ''再帰の回数(通常は省略)
'          blnShowLoginFlg As Boolean      ''ﾛｸﾞｲﾝ画面表示ﾌﾗｸﾞ(TRUE:表示、FALSE:非表示)
'          blnWOraSessionConnectFlag As Boolean    ''ﾛｸﾞ出力用OracleDB接続ﾌﾗｸﾞ(TRUE:接続、FALSE:非接続)
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strUser_Kbn    As String        ''ﾕｰｻﾞ区分
    Dim strSoshiki_Kbn As String        ''組織区分
    Dim nCount         As Integer       ''指定文字の位置
    Dim strErrType     As String        ''ｴﾗｰﾀｲﾌﾟ
    Dim intRet         As Integer       ''返り値
    Dim strNetUser     As String
    Dim strNetpass     As String
'    Dim strPass        As String
    Dim strPath        As String
    Dim strMsg         As String
'    Static stnCount    As Integer       ''ｶｳﾝﾀ
    
    GF_Initialize = False
    
    '二重起動ﾁｪｯｸ
    If App.PrevInstance = True Then Exit Function

'    ''再帰回数＋１
'    stnCount = stnCount + 1
            
    ''INIﾌｧｲﾙ存在ﾁｪｯｸ
    If Dir(App.Path & "\" & gINIFILE) = "" Then
'        stnCount = 3
        strErrType = "INI"
        Err.Raise Number:=vbObjectError, Description:=gINIFILE & " が存在しません。"
    End If
    
    ''ｺﾏﾝﾄﾞﾗｲﾝ引数ﾁｪｯｸ
    If (nCountFlg > 0 Or LF_GetCommndLine = False) And (blnShowLoginFlg = True) Then
    
        ''ﾕｰｻﾞ名の取得
        If LF_GetConnectUserName(gstrUserID) = False Then Exit Function
    
        '''ﾛｸﾞｲﾝ画面呼出し
        'ｷｬﾝｾﾙ時、終了
        If Len(Trim(gstrUserID)) = 0 Then
            If LF_LogIn = False Then Exit Function
        End If
    End If
    
    'ﾕｰｻﾞID設定
    basMsgFunc.UserID = gstrUserID
    '端末CD設定
    basMsgFunc.TerminalCD = Environ("COMPUTERNAME")
    

'    ''INIﾌｧｲﾙ読込み
'
'    ''ﾈｯﾄﾜｰｸ接続ﾁｪｯｸ用ﾊﾟｽ取得
'    strPath = GF_ReadINI("SERVER", "CHECK_PATH")
'    If strPath <> "" Then
'        ''ﾈｯﾄﾜｰｸ接続ﾁｪｯｸ用ﾊﾟｽﾁｪｯｸ
'        If GF_NetWorkShareCheck(strPath) < 0 Then
'            ''ﾈｯﾄﾜｰｸ接続ﾕｰｻﾞ取得
'            strNetUser = GF_ReadINI("SERVER", "USER")
'            ''ﾈｯﾄﾜｰｸ接続ﾊﾟｽﾜｰﾄﾞ取得
'            strNetpass = GF_ReadINI("SERVER", "PASS")
'
'            'ﾈｯﾄﾜｰｸ接続
'            If GF_NetConnect(strNetUser, strNetpass, strPath) = False Then
'                ''確認ﾒｯｾｰｼﾞ表示
'                strMsg = "このまま処理を続行すると印刷処理が使用できない可能性があります。" & vbCr & "続行しますか？"
'                intRet = GF_MsgBox("NetWork", strMsg, "OC", "Q")
'                If intRet = 0 Or intRet = vbCancel Then
'                    '''ｷｬﾝｾﾙ押下時
'                    Exit Function
'                End If
'            End If
'        End If
'    End If
    
    '''DB名取得
    mstrDBInstance = GF_ReadINI("ORACLE", "DSN")
    If mstrDBInstance = "" Then
'        stnCount = 3
        strErrType = "INI"
        Err.Raise Number:=vbObjectError, Description:="DB名が " & gINIFILE & " に正しく設定されていません。"
    End If

    '''DB接続ﾕｰｻﾞ名取得
    mstrDBUser = GF_ReadINI("ORACLE", "USERNAME")
    If mstrDBUser = "" Then
'        stnCount = 3
        strErrType = "INI"
        Err.Raise Number:=vbObjectError, Description:="DB接続ユーザ名が " & gINIFILE & " に正しく設定されていません。"
    End If
    
    '''DB接続ﾊﾟｽﾜｰﾄﾞ取得
    mstrDBPwd = GF_ReadINI("ORACLE", "PASSWORD")
    If mstrDBPwd = "" Then
'        stnCount = 3
        strErrType = "INI"
        Err.Raise Number:=vbObjectError, Description:="DB接続ユーザ名が " & gINIFILE & " に正しく設定されていません。"
    End If
    
    'グローバルユーザ名変数に格納する
'    gstrUserID = mstrDBUser
    
    ''DB接続ﾊﾟｽﾜｰﾄﾞ取得
'    If bolAuthenticationFlg = True Then
'        strPass = mstrDBPwd
'    Else
'        strPass = mstrUserPWD
'    End If
    
    'TIVOLIｻｰﾊﾞﾛｸﾞ出力対象ﾌﾟﾛｸﾞﾗﾑ判定
    If LF_Read_Tivoli_Log_PGM = False Then Exit Function
    
    ''DB接続(DB接続ﾕｰｻﾞ、DB接続ﾊﾟｽﾜｰﾄﾞで接続)
    If GF_DBOpen(mstrDBInstance, mstrDBUser, mstrDBPwd, blnWOraSessionConnectFlag) = False Then Exit Function
    
    If bolAuthenticationFlg = True Then
'        ''ﾕｰｻﾞ認証
'        If LF_UserAuthentication = False Then Exit Function
    End If
    
    GF_Initialize = True
    
    Exit Function
    
ErrHandler:
    Dim strLocation    As String    ''ｴﾗｰ発生場所
    Dim lngErrNum      As Long      ''ｴﾗｰﾅﾝﾊﾞｰ
    Dim strErrMsg      As String    ''ｴﾗｰﾒｯｾｰｼﾞ
    
    strLocation = "GF_Initialize"
    strErrMsg = Err.Description
    
    ''ｴﾗｰﾒｯｾｰｼﾞ表示
    If basMsgFunc.DispErrMsgFlg = True Then
        intRet = GF_MsgBox("Login", strErrMsg, "OK", "E")
    End If
    ''ｴﾗｰﾛｸﾞ出力
    intRet = GF_LogOut(strErrType, strLocation, "", Replace(strErrMsg, vbCr, ""))
    
'    ''3回間違えたら終了
'    If stnCount <= 2 Then
'        '''ﾛｸﾞｲﾝ処理(再帰)
'        If GF_Initialize(stnCount) = False Then Exit Function
'        '''ﾛｸﾞｲﾝ処理、正常終了
'        GF_Initialize = True
'    End If
    
End Function

Public Function LF_Read_Tivoli_Log_PGM() As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : TIVOLIｻｰﾊﾞﾛｸﾞ出力対象ﾌﾟﾛｸﾞﾗﾑ判定
' 機能   :
' 引数   :
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    Dim strPGM      As String
    Dim strErrMsg   As String
    Dim i           As Integer
    
    On Error GoTo ErrHandler
    
    LF_Read_Tivoli_Log_PGM = False
    
    mblnTVL_Log_Flg = False
    
    'TIVOLIサーバログディレクトリを取得
    mstrTVL_Log_Dir = GF_ReadINI("TIVOLI_LOG", "TVL_ERR_LOG")
    
    'TIVOLIサーバログ出力対象プログラムを取得
    For i = 1 To 99
        strPGM = GF_ReadINI("TIVOLI_LOG", "TVL_LOG_EXE_" & Format(i, "00"))
        If Len(Trim(strPGM)) = 0 Then Exit For
        
        If StrComp(App.EXEName, strPGM, vbTextCompare) = 0 Then
            mblnTVL_Log_Flg = True
            Exit For
        End If
    Next i
    
    '出力先チェック
    If (mblnTVL_Log_Flg = True) And (Len(Trim(mstrTVL_Log_Dir)) = 0) Then
        Err.Raise Number:=vbObjectError, Description:="出力先が指定されていません。"
    End If
    
    LF_Read_Tivoli_Log_PGM = True
    Exit Function
    
ErrHandler:
    strErrMsg = Err.Description
    ''ｴﾗｰﾒｯｾｰｼﾞ表示
    If basMsgFunc.DispErrMsgFlg = True Then
        Call GF_MsgBox("Login", "TIVOLIｻｰﾊﾞﾛｸﾞ出力対象ﾌﾟﾛｸﾞﾗﾑの取得に失敗しました。" & vbCrLf & strErrMsg, "OK", "E")
    End If
    ''ｴﾗｰﾛｸﾞ出力
    Call GF_LogOut("VB", "LF_GetConnectUserName", "", "TIVOLIｻｰﾊﾞﾛｸﾞ出力対象ﾌﾟﾛｸﾞﾗﾑの取得に失敗しました。  " & strErrMsg)
End Function

Private Function LF_GetConnectUserName(ByRef strUserName As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : ﾕｰｻﾞ名取得
' 機能   :
' 引数   : strUserName  As String (out)      ﾕｰｻﾞ名
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    Dim strDummy As String
    Dim strErrMsg As String
    
On Error GoTo ErrHandler

    LF_GetConnectUserName = False
    
    strDummy = String(256, Chr(0))
    ''ﾕｰｻﾞ名の取得
    Call GetUserName(strDummy, 256)
    strDummy = Mid(strDummy, 1, InStr(1, strDummy, Chr(0), vbTextCompare) - 1)
    strUserName = Trim(strDummy)
        
    LF_GetConnectUserName = True
        
    Exit Function
ErrHandler:
    strErrMsg = Err.Description
    ''ｴﾗｰﾒｯｾｰｼﾞ表示
    If basMsgFunc.DispErrMsgFlg = True Then
        Call GF_MsgBox("Login", "接続ユーザの取得に失敗しました。" & vbCrLf & strErrMsg, "OK", "E")
    End If
    ''ｴﾗｰﾛｸﾞ出力
    Call GF_LogOut("VB", "LF_GetConnectUserName", "", "接続ユーザの取得に失敗しました。  " & strErrMsg)
End Function

'2007/02/26 Added by N.Kigaku
Public Function GF_GetConnectUserName(ByRef strUserName As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : ﾕｰｻﾞ名取得(ｸﾞﾛｰﾊﾞﾙ版)
' 機能   :
' 引数   : strUserName  As String (out)      ﾕｰｻﾞ名
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    GF_GetConnectUserName = LF_GetConnectUserName(strUserName)
End Function

Private Function LF_GetCommndLine() As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : ｺﾏﾝﾄﾞﾗｲﾝ引数取得
' 機能   :
' 引数   :
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strCmdLine As String
    Dim intPosition As String
    
    LF_GetCommndLine = False
    
    strCmdLine = ""
    strCmdLine = Trim(Command())
    
    'コマンドラインが何もない時 EXIT
    If Len(strCmdLine) = 0 Then Exit Function
    
'    If InStr(1, strCmdLine, "/", vbTextCompare) <> 0 Then

    intPosition = InStr(1, strCmdLine, " ", vbTextCompare)
    If intPosition <> 0 Then
    
        ''ﾕｰｻﾞID取得
        gstrUserID = Trim(Left(strCmdLine, InStr(1, strCmdLine, " ", vbTextCompare) - 1))
'2005/07/27 Added by N.Kigaku
''引合ﾒﾆｭｰ起動対応　ﾕｰｻﾞ名に"/"がある場合は削除する
        If InStr(1, gstrUserID, "/", vbTextCompare) <> 0 Then
        
            gstrUserID = Trim(Left(strCmdLine, InStr(1, gstrUserID, "/", vbTextCompare) - 1))
        End If
        
        ''ﾕｰｻﾞﾊﾟｽﾜｰﾄﾞ
        mstrUserPWD = ""

        ''EXEごとの引数取得
        gstrArgument = Trim(Mid(strCmdLine, InStr(1, strCmdLine, " ", vbTextCompare) + 1))
        
        ''ﾕｰｻﾞID取得
'        gstrUserID = Trim(Left(strCmdLine, InStr(1, strCmdLine, "/", vbTextCompare) - 1))
        ''DB接続ﾊﾟｽﾜｰﾄﾞ取得
    '    mstrDBPwd = Trim(Mid(strCmdLine, InStr(1, strCmdLine, "/", vbTextCompare) + 1))
        ''ﾕｰｻﾞﾊﾟｽﾜｰﾄﾞ取得
'        If InStr(1, strCmdLine, " ", vbTextCompare) = 0 Then
'            ''EXEごとの引数がない場合
'            mstrUserPWD = Trim(Mid(strCmdLine, InStr(1, strCmdLine, "/", vbTextCompare) + 1))
'        Else
'            '''EXEごとの引数がある場合
'            mstrUserPWD = Trim(Mid(strCmdLine, InStr(1, strCmdLine, "/", vbTextCompare) + 1, InStr(1, strCmdLine, " ", vbTextCompare) - 1 - InStr(1, strCmdLine, "/", vbTextCompare)))
            ''EXEごとの引数取得
'            gstrArgument = Trim(Mid(strCmdLine, InStr(1, strCmdLine, " ", vbTextCompare) + 1))
'        End If
    Else
        ''ﾕｰｻﾞID取得
        gstrUserID = Trim(strCmdLine)
        ''ﾊﾟｽﾜｰﾄﾞ
        mstrUserPWD = ""
    End If
    
    LF_GetCommndLine = True
    
    Exit Function
      
ErrHandler:
    Dim strLocation    As String    ''ｴﾗｰ発生場所
    Dim lngErrNum      As Long      ''ｴﾗｰﾅﾝﾊﾞｰ
    Dim strErrMsg      As String    ''ｴﾗｰﾒｯｾｰｼﾞ
    Dim strErrType     As String    ''ｴﾗｰﾀｲﾌﾟ
    Dim intRet         As Integer   ''返り値
    
    strLocation = "LF_GetCommndLine"
    strErrMsg = "コマンドライン引数が違います。"
    strErrType = "INI"
    
    ''ｴﾗｰﾒｯｾｰｼﾞ表示
    If basMsgFunc.DispErrMsgFlg = True Then
        intRet = GF_MsgBox("Login", strErrMsg, "OK", "E")
    End If
    ''ｴﾗｰﾛｸﾞ出力
    intRet = GF_LogOut(strErrType, strLocation, "", strErrMsg)
    
End Function

Private Function LF_LogIn(Optional intLoginKBN As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : Login
' 機能   : Login画面呼出し
' 引数   : intLoginKBN As Integer     ''ﾛｸﾞｲﾝ区分 0:従業員番号, 1:ﾊﾟｽﾜｰﾄﾞ
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    Dim bolCancel As Boolean
    
    LF_LogIn = False
    
    Screen.MousePointer = vbDefault
    
    DoEvents
    
    ''ﾛｸﾞｲﾝ画面呼出し
    Screen.MousePointer = vbDefault
    frmLogin.LoginKBN = intLoginKBN
    frmLogin.Show vbModal
    Screen.MousePointer = vbHourglass
    
    'ﾊﾟｽﾜｰﾄﾞ取得
'    mstrDBPwd = frmLogin.Password
    'mstrUserPWD = frmLogin.Password
    'ｷｬﾝｾﾙﾌﾗｸﾞ取得
    bolCancel = frmLogin.CancelFlg
    
    'ﾛｸﾞｲﾝ画面ｱﾝﾛｰﾄﾞ
    Unload frmLogin
    
    DoEvents
    
    'Login画面でｷｬﾝｾﾙされたか？
    If bolCancel = True Then Exit Function
                     
    LF_LogIn = True
    
End Function

Private Function LF_UserAuthentication() As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : ﾕｰｻﾞ認証
' 機能   : ﾛｸﾞｲﾝ画面にて入力されたﾕｰｻﾞの認証を行う
' 引数   :
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strErrType   As String
    Dim strMsgID   As String
    Dim strSQL       As String         ''SQL文
    Dim oraDyna      As OraDynaset     ''ﾀﾞｲﾅｾｯﾄ
    Dim bolRetFlg    As Boolean        ''ﾌﾗｸﾞ
    Dim intRet       As Integer
    Dim strPassWord  As String
    
    ''ﾕｰｻﾞ認証ﾌﾗｸﾞ初期化
    LF_UserAuthentication = False
    
    'SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM THJUSRMR"
    strSQL = strSQL & " WHERE SYAINCD = '" & UCase(Trim(gstrUserID)) & "'"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        ''''該当ﾃﾞｰﾀなし
        strErrType = "Login"
        '社員ﾏｽﾀに登録されていません。
        strMsgID = "WTG008"
        intRet = GF_MsgBoxDB(strErrType, strMsgID, "OK", "E")
        intRet = GF_LogOutDB(strErrType, "mMLR_UserAuthentication", strMsgID)
        Exit Function
    End If
    
'    strPassWord = GF_VarToStr(oraDyna![Password])
'    If UCase(mstrUserPWD) <> strPassWord Then
'        '''ﾊﾟｽﾜｰﾄﾞ間違い
'        strErrType = "Login"
'        'ﾊﾟｽﾜｰﾄﾞが違います。
'        strMsgID = "WTG002"
'        intRet = GF_MsgBoxDB(strErrType, strMsgID, "OK", "E")
'        intRet = GF_LogOutDB(strErrType, "mMLR_UserAuthentication", strMsgID)
'        Exit Function
'    End If
    
    'ﾀﾞｲﾅｾｯﾄ開放
    Set oraDyna = Nothing
    
    LF_UserAuthentication = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("LF_UserAuthentication", strSQL)

End Function

Public Function GF_CreateArgument() As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : ﾕｰｻﾞ情報引数作成
' 機能   : EXEに渡すﾕｰｻﾞ情報引数を作成する
' 引数   :
' 戻り値 : String     ﾕｰｻﾞ情報引数(ﾕｰｻﾞID/ﾕｰｻﾞﾊﾟｽﾜｰﾄﾞ EXEごとの引数)
' 備考   :
'------------------------------------------------------------------------------
'    GF_CreateArgument = gstrUserID & "/" & mstrUserPWD
    GF_CreateArgument = gstrUserID
End Function

Public Sub GS_DelControlBox(frmForm As Form)
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 閉じるﾎﾞﾀﾝ使用不可
' 機能   : 閉じるﾎﾞﾀﾝを使用不可にする
' 引数   : frmForm As Form      ''ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ
' 備考   :
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
' 機能名 : INIﾌｧｲﾙを読み出す
' 機能   : INIﾌｧｲﾙより指定のｾｸｼｮﾝ、指定のKeyの値を取得する
' 引数   : strSection As String   ''ｾｸｼｮﾝ
'       ： strKey     As String   ''Key
' 戻り値 : 指定したKeyの値
' 備考   :
'------------------------------------------------------------------------------

    Dim lngRet  As Long            ''GetPrivateProfileStringの戻り値　0：ｴﾗｰ
    Dim strBuff As String * 256

    GF_ReadINI = ""
    lngRet = GetPrivateProfileString(strSection, strKey, "", _
                                        strBuff, 255, App.Path & "\" & gINIFILE)

    ''文字列数が"0"の時、エラー
    If lngRet <> 0 Then
        GF_ReadINI = strConv(MidB(strConv(strBuff, vbFromUnicode), 1, lngRet), vbUnicode)
    End If

End Function
