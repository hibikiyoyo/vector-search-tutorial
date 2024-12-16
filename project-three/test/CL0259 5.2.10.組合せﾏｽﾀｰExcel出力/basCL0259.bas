Attribute VB_Name = "basCL0259"
' @(h) basCL0259.bas  ver1.0.0.1 ( 2004/10/20 J.Hamaji )
'------------------------------------------------------------------------------
' @(s)
'   プロジェクト名  : TLFﾌﾟﾛｼﾞｪｸﾄ
'   モジュール名    : basCL0259
'   ファイル名      : basCL0259.bas
'   Version         : 1.0.0.1
'   機能説明       ： 組合せﾏｽﾀｰExcel出力
'   作成者         ： J.Hamaji
'   作成日         ： 2004/10/20
'   修正履歴       ： 2004/10/28 THS T.Y (ｸﾗｲｱﾝﾄに直接Excel出力する)
'   　　　　       ： 2005/07/07 THS J.Yamaoka (上書きﾒｯｾｰｼﾞ付加情報ｾｯﾄVerに変更)
'   　　　　       ： 2006/04/13 THS Sugawara Ver1.0.0.1  6万件以上のデータで次のシートに出力されない不具合を修正
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 環境宣言
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' モジュール変数宣言
'------------------------------------------------------------------------------

Public iCount As Integer
Public strMsg() As String

Public gstrServerPath As String     'EXCEL出力先のパス(サーバ)
Public gstrClientPath As String     'EXCEL出力先のパス(クライアント)
Public gstrFileName   As String     'EXCELファイル名

'------------------------------------------------------------------------------
' モジュール定数宣言
'------------------------------------------------------------------------------
Public Const cProtectColor As Long = &H8000000F
Public Const cNoProtectColor As Long = &H80000005

Public Const gaOK           As Integer = 0
Public Const gaNG           As Integer = -1
Public Const gaNothing      As Integer = 1
Public Const strEndDAte     As String = "99999999"
Public Const strProAlarm    As String = "0"         '生管アラームフラグ
Public Const strKumiKbn     As String = "TGLOBAL_SET_CS"

Private Const mstrPGMID As String = "CL0259"        'プログラムID

Private Sub Main()
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:Main処理
' 機能　　　:サブシステム初期処理
' 引数　　　:
' 機能説明　:サブシステム画面表示初期処理
'------------------------------------------------------------------------------
    On Error GoTo Err_Main
    
    Dim bolRes As Boolean
    
    'マウスポインタ設定(砂時計)
    Screen.MousePointer = vbHourglass
    
    '複数起動の抑止
    If App.PrevInstance = True Then
        MsgBox "すでに起動されています。", vbExclamation
        Exit Sub
    End If
    
    'Main処理初期化
    If LF_Main_Initialize = False Then
        'マウスポインタ設定
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    '初期処理
    If GF_Initialize() = False Then
        'マウスポインタ設定
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    '『組合せﾏｽﾀｰExcel出力』画面表示
    frmCL0259.Show
    
    'フォームロード完了判定
    If (frmCL0259.LoadFlag = False) Then
        'フォームアンロード
        Unload frmCL0259
        'DB切断
        If Not (gOraParam Is Nothing) Or _
           Not (gOraDataBase Is Nothing) Or _
           Not (gOraSession Is Nothing) Then
    
            Call GS_DBClose
    
        End If
    Else
        frmCL0259.cmbExportCs.SetFocus
    End If
    
    'iniファイルの環境設定情報を取得
    Call LF_GetiniInf
    
    'マウスポインタ設定
    Screen.MousePointer = vbDefault

Exit_Main:
    Exit Sub

Err_Main:
    'Main内の実行時エラー処理
    Call GS_ErrorHandler("Main")
    Resume Exit_Main
End Sub

Private Function LF_Main_Initialize() As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : Main処理時の初期化
' 機能   :
' 引数   :
' 戻り値 :
' 備考   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    'ｴﾗｰﾒｯｾｰｼﾞ表示ﾌﾗｸﾞ設定
    basMsgFunc.DispErrMsgFlg = False
    
    'ﾌﾟﾛｸﾞﾗﾑ名設定
    basMsgFunc.PGMCD = mstrPGMID
    
    'ﾛｸﾞﾌｧｲﾙ名設定
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
' 機能名    : iniファイルの環境設定情報を取得(サーバ/クライアント)
' 機能      :
' 引数      :
' 戻り値    : TRUE：正常 FALSE:エラー Boolean
' 機能説明　:
'--------------------------------------------------------------------------------
On Error GoTo ErrHandler
        
    LF_GetiniInf = False
         
    ''DEL 2004/10/28 THS T.Y (ｸﾗｲｱﾝﾄに直接Excel出力する) START>>>>>
'''''    'iniファイルのEXCEL出力先フォルダ取得(サーバ)
'''''    gstrServerPath = GF_ReadINI("MASTER", "MST_EXPORT_DIR")
'''''
'''''    'EXCEL出力先フォルダ取得の判定(サーバ)
'''''    If gstrServerPath = "" Then
'''''        '出力先フォルダ情報取得失敗
'''''        'ログ出力
'''''        Call GF_GetMsg_Addition("WTK398", , False, True)
'''''        'MSG表示
'''''        Call GF_MsgBoxDB(frmCL0259.Caption, "WTK398", "OK", "C")
'''''        Exit Function
'''''    'パス名の最後尾に"\"がついているか
'''''    ElseIf Right(gstrServerPath, 1) <> "\" Then
'''''        gstrServerPath = gstrServerPath & "\"
'''''    End If
    ''<<<<<END
    
    'iniファイルのEXCEL出力先フォルダ取得(クライアント)
    gstrClientPath = GF_ReadINI("MASTER", "MST_I/O_DIR")
    
    'EXCEL出力先フォルダ取得の判定(クライアント)
    If gstrClientPath = "" Then
        '出力先フォルダ情報取得失敗
        'ログ出力
        Call GF_GetMsg_Addition("WTK399", , False, True)
        'MSG表示
        Call GF_MsgBoxDB(frmCL0259.Caption, "WTK399", "OK", "C")
        Exit Function
    'パス名の最後尾に"\"がついているか
    ElseIf Right(gstrClientPath, 1) <> "\" Then
        gstrClientPath = gstrClientPath & "\"
    End If
    
    'iniファイルのEXCELファイル名の取得
    gstrFileName = GF_ReadINI(mstrPGMID, "OUTPUT_EXCEL_FILE")
    
    'EXCELファイル名取得の判定
    If gstrFileName = "" Then
        'EXCELファイル名取得失敗
        'ログ出力
        Call GF_GetMsg_Addition("WTK400", , False, True)
        'MSG表示
        Call GF_MsgBoxDB(frmCL0259.Caption, "WTK400", "OK", "C")
        Exit Function
    End If
    
    LF_GetiniInf = True
    
    Exit Function

ErrHandler:

    Call GS_ErrorHandler("LF_GetiniInf")

End Function

