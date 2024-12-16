Attribute VB_Name = "basMsgFunc"
' @(h) MsgFunc.bas  ver1.00 ( 2003/02/05 N.Kigaku )
'------------------------------------------------------------------------------
' @(s)
'   プロジェクト名 : TLFﾌﾟﾛｼﾞｪｸﾄ
'   モジュール名　 : basMsgFunc
'   ファイル名　　 : MsgFunc.bas
'   バージョン　　 : 1.00
'   機能説明　　　 : ﾒｯｾｰｼﾞ処理関連
'   作成者　　　　 : N.Kigaku
'   作成日　　　　 : 2003/02/05
'   修正履歴　　　 : 2004/03/22 N.Kigaku GF_ExeLogOut 追加
'                    2004/08/24 N.Kigaku GF_GetMsg_Additionにﾒｯｾｰｼﾞﾎﾞｯｸｽ表示で
'                                        ｱｲｺﾝを変更できるように修正。
'                    2004/09/10 N.Kigaku GF_GetMsg_MasterMente 追加
'                    2005/07/05 N.Kigaku GF_GetMsg_Addition, GF_GetMsg_MasterMenteにOn Error文追記
'                    2006/12/05 N.Kigaku ｵﾗｸﾙ8.1.7 Nocache対応 検索時、ReadOnlyからNocacheに変更
'                    2007/07/30 N.Kigaku GF_WriteLogDataのﾛｸﾞ出力のﾕｰｻﾞIDを7桁で切るように修正
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 環境宣言
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' モジュール定数宣言
'------------------------------------------------------------------------------
Private Const mstrCommonErrMsgCD As String = "WTK009"     ''共通ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ

'------------------------------------------------------------------------------
' モジュール変数宣言
'------------------------------------------------------------------------------
Private mstrLogFile         As String       ''ﾊﾟｽ付きﾛｸﾞﾌｧｲﾙ名      （ﾌﾟﾛﾊﾟﾃｨ）
Private mblnMsgDispFlg      As Boolean      ''ｴﾗｰﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞ（ﾌﾟﾛﾊﾟﾃｨ）
Private mstrUserID          As String       ''ﾕｰｻﾞID（ﾌﾟﾛﾊﾟﾃｨ）
Private mstrPGMCD           As String       ''ﾌﾟﾛｸﾞﾗﾑCD（ﾌﾟﾛﾊﾟﾃｨ）
Private mstrTerminalCD      As String       ''端末CD（ﾌﾟﾛﾊﾟﾃｨ）

'ﾒｯｾｰｼﾞ格納配列 START>>>>>
Private Type TYPE_MSG
    TYPE_MSG_CD     As String 'ﾒｯｾｰｼﾞCD
    TYPE_MSG_NAIYO  As String 'ﾒｯｾｰｼﾞ内容
End Type
Private ADD_TYPE_MSG() As TYPE_MSG
'<<<<<END



Public Property Get LogFile() As String
'------------------------------------------------------------------------------
' 機能名　　: ﾛｸﾞﾌｧｲﾙ名ﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: なし
' 戻り値　　: ﾊﾟｽ付きﾛｸﾞﾌｧｲﾙ名
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨの値を戻す
'------------------------------------------------------------------------------
    LogFile = mstrLogFile
End Property

Public Property Let LogFile(ByVal strLog As String)
'------------------------------------------------------------------------------
' 機能名　　: ﾛｸﾞﾌｧｲﾙ名ﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: ByVal strLog As String    ''ﾊﾟｽ付きﾛｸﾞﾌｧｲﾙ名
' 戻り値　　: なし
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨに値を入れる
'------------------------------------------------------------------------------
    mstrLogFile = strLog
End Property

Public Property Get DispErrMsgFlg() As Boolean
'------------------------------------------------------------------------------
' 機能名　　: ｴﾗｰﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: なし
' 戻り値　　: ｴﾗｰﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞ
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨの値を戻す
'------------------------------------------------------------------------------
    DispErrMsgFlg = mblnMsgDispFlg
End Property

Public Property Let DispErrMsgFlg(ByVal blnFlg As Boolean)
'------------------------------------------------------------------------------
' 機能名　　: ｴﾗｰﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: ByVal blnFlg As Boolean    ''ｴﾗｰﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞ
' 戻り値　　: なし
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨに値を入れる
'------------------------------------------------------------------------------
    mblnMsgDispFlg = blnFlg
End Property

Public Property Get UserID() As String
'------------------------------------------------------------------------------
' 機能名　　: ﾕｰｻﾞIDﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: なし
' 戻り値　　: ﾕｰｻﾞID
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨの値を戻す
'------------------------------------------------------------------------------
    UserID = mstrUserID
End Property

Public Property Let UserID(ByVal strUserID As String)
'------------------------------------------------------------------------------
' 機能名　　: ﾕｰｻﾞIDﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: ByVal strUserID As String    ''ﾕｰｻﾞID
' 戻り値　　: なし
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨに値を入れる
'------------------------------------------------------------------------------
    mstrUserID = strUserID
End Property

Public Property Get PGMCD() As String
'------------------------------------------------------------------------------
' 機能名　　: ﾌﾟﾛｸﾞﾗﾑCDﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: なし
' 戻り値　　: ﾌﾟﾛｸﾞﾗﾑCD
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨの値を戻す
'------------------------------------------------------------------------------
    PGMCD = mstrPGMCD
End Property

Public Property Let PGMCD(ByVal strPGMCD As String)
'------------------------------------------------------------------------------
' 機能名　　: ﾌﾟﾛｸﾞﾗﾑIDﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: ByVal strPGMCD As String    ''ﾌﾟﾛｸﾞﾗﾑID
' 戻り値　　: なし
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨに値を入れる
'------------------------------------------------------------------------------
    mstrPGMCD = strPGMCD
End Property

Public Property Get TerminalCD() As String
'------------------------------------------------------------------------------
' 機能名　　: 端末CDﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: なし
' 戻り値　　: 端末CD
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨの値を戻す
'------------------------------------------------------------------------------
    TerminalCD = mstrTerminalCD
End Property

Public Property Let TerminalCD(ByVal strTerminalCd As String)
'------------------------------------------------------------------------------
' 機能名　　: 端末CDﾌﾟﾛﾊﾟﾃｨ
' 機能　　　:
' 引数　　　: ByVal strTerminalCD As String    ''端末CD
' 戻り値　　: なし
' 機能説明　: ﾌﾟﾛﾊﾟﾃｨに値を入れる
'------------------------------------------------------------------------------
    mstrTerminalCD = strTerminalCd
End Property



Public Function GF_MsgBox(sTITLE As String, sMSG As String, _
                            sBTN As String, sICON As String, _
                              Optional iDefBTN As Integer = 1) As Integer
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示
' 機能　　　:　ﾒｯｾｰｼﾞﾎﾞｯｸｽを表示する
' 引数　　　:　[I] sTITLE     As String    ''ﾀｲﾄﾙ
' 　　　　　 　[I] sMSG       As String    ''ﾒｯｾｰｼﾞ
' 　　　　　 　[I] sBTN       As String    ''ﾎﾞﾀﾝﾀｲﾌﾟ
' 　　　　　 　[I] sICON      As String    ''ｱｲｺﾝﾀｲﾌﾟ
'             iDefBTN    As Integer   ''ﾎﾞﾀﾝﾃﾞﾌｫﾙﾄﾌｫｰｶｽ位置
'             1 = 第1ﾎﾞﾀﾝ / 2 = 第2ﾎﾞﾀﾝ / 3 = 第3ﾎﾞﾀﾝ / 4 = 第4ﾎﾞﾀﾝ
' 戻り値　　:　1 = OK / 2 = CANCEL / 6 = YES / 7 = NO
' 　　　　　 　0 = ERROR
' 機能説明　:
'------------------------------------------------------------------------------
    
    Dim mlngStyle As Long
    Dim mintDefBtn As Integer
    
    On Error GoTo ErrHandler

    '[ﾎﾞﾀﾝ]
    Select Case UCase(Trim(sBTN))
        Case "OK"   '[OK]
                    mlngStyle = mlngStyle + vbOKOnly
        Case "OC"   '[OK][CANCEL]
                    mlngStyle = mlngStyle + vbOKCancel
        Case "YNC"  '[YES][NO][CANCEL]
                    mlngStyle = mlngStyle + vbYesNoCancel
        Case "YN"   '[YES][NO]
                    mlngStyle = mlngStyle + vbYesNo
    End Select

    '[ｱｲｺﾝ]
    Select Case UCase(Trim(sICON))
        Case "C"    '[警告]
                    mlngStyle = mlngStyle + vbCritical
                    'ﾋﾞｰﾌﾟ音鳴らす
                    Beep
        Case "Q"    '[問い合わせ]
                    mlngStyle = mlngStyle + vbQuestion
        Case "E"    '[注意]
                    mlngStyle = mlngStyle + vbExclamation
                    'ﾋﾞｰﾌﾟ音鳴らす
                    Beep
        Case "I"    '[情報]
                    mlngStyle = mlngStyle + vbInformation
    End Select
    
    '[ﾃﾞﾌｫﾙﾄﾎﾞﾀﾝ]
    Select Case iDefBTN
        Case 1
            mintDefBtn = vbDefaultButton1
        Case 2
            mintDefBtn = vbDefaultButton2
        Case 3
            mintDefBtn = vbDefaultButton3
        Case 4
            mintDefBtn = vbDefaultButton4
        Case Else
            mintDefBtn = vbDefaultButton1
    End Select
    
    'ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示
    GF_MsgBox = MsgBox(sMSG, mlngStyle + mintDefBtn, sTITLE)
    
    Exit Function
    
ErrHandler:
    
    GF_MsgBox = 0
    
End Function

Public Function GF_MsgBoxDB(sTITLE As String, sMSGID As String, sBTN As String, _
                              sICON As String, Optional iDefBTN As Integer = 1) As Integer
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾒｯｾｰｼﾞ表示（DB版）
' 機能　　　:　ﾒｯｾｰｼﾞﾎﾞｯｸｽ又はｽﾃｰﾀｽﾊﾞｰに指定ﾒｯｾｰｼﾞを表示する
' 引数　　　:　[I] sTITLE     As String    ''ﾀｲﾄﾙ
' 　　　　　 　[I] sMSGID     As String    ''ﾒｯｾｰｼﾞID
' 　　　　　 　[I] sBTN       As String    ''ﾎﾞﾀﾝﾀｲﾌﾟ
' 　　　　　 　[I] sICON      As String    ''ｱｲｺﾝﾀｲﾌﾟ
'             [I] iDefBTN    As Integer   ''ﾎﾞﾀﾝﾃﾞﾌｫﾙﾄﾌｫｰｶｽ位置
'             1 = 第1ﾎﾞﾀﾝ / 2 = 第2ﾎﾞﾀﾝ / 3 = 第3ﾎﾞﾀﾝ / 4 = 第4ﾎﾞﾀﾝ
' 戻り値　　:　1 = OK / 2 = CANCEL / 6 = YES / 7 = NO / 9 = ｽﾃｰﾀｽﾊﾞｰへ出力
' 　　　　　 　0 = ERROR
' 機能説明　:　ﾒｯｾｰｼﾞはﾃﾞｰﾀﾍﾞｰｽから取得
' 　　　　　 　該当するﾃﾞｰﾀがない場合は、ｴﾗｰﾒｯｾｰｼﾞを表示
'------------------------------------------------------------------------------
    
    Dim oDynaset   As OraDynaset
    Dim mlngStyle  As Long
    Dim sSQL       As String
    Dim mstrOutMsg As String
    Dim mintOutFlg As Integer
    Dim mbolRet    As Boolean
    Dim mintDefBtn As Integer
    Dim intRet     As Integer
    
    On Error GoTo ErrHandler
    
    '[ﾎﾞﾀﾝ]
    Select Case UCase(Trim(sBTN))
        Case "OK"   '[OK]
                    mlngStyle = mlngStyle + vbOKOnly
        Case "OC"   '[OK][CANCEL]
                    mlngStyle = mlngStyle + vbOKCancel
        Case "YNC"  '[YES][NO][CANCEL]
                    mlngStyle = mlngStyle + vbYesNoCancel
        Case "YN"   '[YES][NO]
                    mlngStyle = mlngStyle + vbYesNo
    End Select
    
    '[ｱｲｺﾝ]
    Select Case UCase(Trim(sICON))
        Case "C"    '[警告]
                    mlngStyle = mlngStyle + vbCritical
                    'ﾋﾞｰﾌﾟ音鳴らす
                    Beep
        Case "Q"    '[問い合わせ]
                    mlngStyle = mlngStyle + vbQuestion
        Case "E"    '[注意]
                    mlngStyle = mlngStyle + vbExclamation
                    'ﾋﾞｰﾌﾟ音鳴らす
                    Beep
        Case "I"    '[情報]
                    mlngStyle = mlngStyle + vbInformation
    End Select
    
    '[ﾃﾞﾌｫﾙﾄﾎﾞﾀﾝ]
    Select Case iDefBTN
        Case 1
            mintDefBtn = vbDefaultButton1
        Case 2
            mintDefBtn = vbDefaultButton2
        Case 3
            mintDefBtn = vbDefaultButton3
        Case 4
            mintDefBtn = vbDefaultButton4
        Case Else
            mintDefBtn = vbDefaultButton1
    End Select
    
    'SQL文生成
    sSQL = ""
    sSQL = "SELECT * FROM THJMSG WHERE MSGCD = '" & UCase(Trim(sMSGID)) & "'"
    
    'ﾀﾞｲﾅｾｯﾄ生成
    Set oDynaset = gOraDataBase.CreateDynaset(sSQL, ORADYN_NOCACHE)
    
    'ﾃﾞｰﾀが見つかった場合
    If (oDynaset.EOF = False) Then
        mstrOutMsg = GF_VarToStr(oDynaset![MSGNAIYO])
        mintOutFlg = GF_VarToNum(oDynaset![OUTFLG])
    Else
        'ﾀﾞｲﾅｾｯﾄ解放
        Set oDynaset = Nothing
        Beep
        intRet = MsgBox("ﾒｯｾｰｼﾞﾃｰﾌﾞﾙに登録されていないﾒｯｾｰｼﾞIDが指定されました。" & vbCrLf & _
                        "MSGID -> [" & UCase(Trim(sMSGID)) & "]", vbOKOnly + vbExclamation + mintDefBtn, "GF_MsgBoxDB")
        GF_MsgBoxDB = 0
        Exit Function
    End If
    
    'ﾀﾞｲﾅｾｯﾄ解放
    Set oDynaset = Nothing
    
    '出力先選択
    If (mintOutFlg = 0) Then
        'ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示
        GF_MsgBoxDB = MsgBox(GF_CnvCtrChar(mstrOutMsg), mlngStyle + mintDefBtn, sTITLE)
    Else
        'ｽﾃｰﾀｽﾊﾞｰ表示
        Screen.ActiveForm.stbStatusBar.Panels(2).Text = GF_DelCtrChar(mstrOutMsg)
        GF_MsgBoxDB = 9
    End If
    
    Exit Function
    
ErrHandler:
    
    Call GS_ErrorHandler("GF_MsgBoxDB", "")
    
    GF_MsgBoxDB = 0
    
End Function

Public Function GF_GetMsg(strMsgCD As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾒｯｾｰｼﾞ取得
' 機能　　　:　ﾒｯｾｰｼﾞをﾃﾞｰﾀﾍﾞｰｽから取得する
' 引数　　　:　[I] strMsgCD As String             ''ﾒｯｾｰｼﾞCD
' 戻り値　　:　取得したﾒｯｾｰｼﾞ
' 　　　　　 　該当ﾃﾞｰﾀがない又はｴﾗｰの時はMsgErr[ﾒｯｾｰｼﾞｺｰﾄﾞ]
' 機能説明　:
'------------------------------------------------------------------------------

    Dim oDynaset   As OraDynaset
    Dim sSQL       As String

    On Error GoTo ErrHandler

    'SQL文生成
    sSQL = ""
    sSQL = "SELECT * FROM THJMSG WHERE MSGCD = '" & UCase(Trim(strMsgCD)) & "'"

    'ﾀﾞｲﾅｾｯﾄ生成
    Set oDynaset = gOraDataBase.CreateDynaset(sSQL, ORADYN_NOCACHE)

    If oDynaset.EOF = False Then
        GF_GetMsg = GF_VarToStr(oDynaset![MSGNAIYO])
    Else
        GF_GetMsg = " MsgErr[" & strMsgCD & "] "
    End If

    Set oDynaset = Nothing

    Exit Function

ErrHandler:

    Call GS_ErrorHandler("GF_GetMsg", sSQL)

    GF_GetMsg = " MsgErr[" & strMsgCD & "] "

End Function

Public Function GF_GetMsgInfo(ByVal strMsgCD As String, _
                              ByRef strMsg As String, _
                              ByRef strMsgLevel As String, _
                              Optional blnNonErrFlg As Boolean = False) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾒｯｾｰｼﾞ情報取得
' 機能　　　:　ﾒｯｾｰｼﾞ情報をﾃﾞｰﾀﾍﾞｰｽから取得する
' 引数　　　:　[I] strMsgCD As String               ''ﾒｯｾｰｼﾞｺｰﾄﾞ
' 　　　　　　 [O] strMsg As String                 ''ﾒｯｾｰｼﾞ
' 　　　　　　 [O] strMsgLevel As String            ''ﾒｯｾｰｼﾞﾚﾍﾞﾙ
'             [I]  blnNonErrFlg As Boolean         ''ｴﾗｰ処理有無ﾌﾗｸﾞ
'                                                      False:ｴﾗｰ処理有り,True:なし
' 戻り値　　:　取得したﾒｯｾｰｼﾞ
' 　　　　　 　該当ﾃﾞｰﾀがない又はｴﾗｰの時はMsgErr[ﾒｯｾｰｼﾞｺｰﾄﾞ]
' 機能説明　:
'------------------------------------------------------------------------------

    Dim oDynaset   As OraDynaset
    Dim sSQL       As String

    On Error GoTo ErrHandler

    'SQL文生成
    sSQL = ""
    sSQL = "SELECT MSGNAIYO,MSGLEVEL FROM THJMSG WHERE MSGCD = '" & UCase(Trim(strMsgCD)) & "'"

    'ﾀﾞｲﾅｾｯﾄ生成
    Set oDynaset = gOraDataBase.CreateDynaset(sSQL, ORADYN_NOCACHE)

    If oDynaset.EOF = False Then
        strMsg = GF_VarToStr(oDynaset![MSGNAIYO])
        strMsgLevel = GF_VarToStr(oDynaset![MSGLEVEL])
    Else
        strMsg = " MsgErr[" & strMsgCD & "] "
        strMsgLevel = ""
    End If

    Set oDynaset = Nothing

    GF_GetMsgInfo = True

    Exit Function

ErrHandler:
    If blnNonErrFlg = False Then
        Call GS_ErrorHandler("GF_GetMsgInfo", sSQL)
    End If
    
    GF_GetMsgInfo = False

End Function

Public Function GF_GetMsg_Addition(ByVal strMsgCD As String, _
                          Optional ByVal vntAddMsg As Variant = "", _
                          Optional ByVal blnDispFlg As Boolean = False, _
                          Optional ByVal blnLogTblFlg As Boolean = False, _
                          Optional ByVal strInfo As String = "", _
                          Optional ByVal blnTivoliLogFlg As Boolean = True, _
                          Optional ByVal strICON As String = "E") As String
'--------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾒｯｾｰｼﾞ取得（付加文字列付き）
' 機能　　　:　ﾒｯｾｰｼﾞをﾃﾞｰﾀﾍﾞｰｽから取得する
'
' 引数　　　: [I] strMsgCD As String                ''ﾒｯｾｰｼﾞｺｰﾄﾞ
'　　　　　　 [I] vntAddMsg As Variant              ''付加文字列の配列　（添字は0から開始）
'                                                       1件の時は配列でなくてもOK
'　　　　　　 [I] ByVal blnDispFlg As Boolean       ''ｴﾗｰﾒｯｾｰｼﾞﾎﾞｯｸｽ出力ﾌﾗｸﾞ
'　　　　　　 [I] ByVal blnLogTblFlg As Boolean     ''ﾛｸﾞﾃｰﾌﾞﾙ出力ﾌﾗｸﾞ
'　　　　　　 [I] ByVal strInfo As String           ''付与情報
'　　　　　　 [I] ByVal blnTivoliLogFlg As Boolean  ''Tivoliﾛｸﾞ出力ﾌﾗｸﾞ
' 　　　　 　 [I] ByVal strICON As String           ''ｱｲｺﾝﾀｲﾌﾟ  2004/08/24 Add by N.Kigaku
'
' 戻り値　　: 取得したﾒｯｾｰｼﾞ
'　　　　　　 該当ﾃﾞｰﾀがない又はｴﾗｰの時は "MsgErr[ﾒｯｾｰｼﾞｺｰﾄﾞ]"
' 機能説明　: DBから取得したﾒｯｾｰｼﾞの%1～%nまで付加文字列で置き換える
'--------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL      As String
    Dim oDynaset    As OraDynaset
    Dim strMessage  As String
    Dim strCnvMsg   As String
    Dim strMsgLevel As String
    Dim intMsgCount As Integer
    Dim i           As Integer
    Dim strTemp     As String
    Dim intRet      As Integer
    
    
    '' 配列の数を数える
    If IsArray(vntAddMsg) = True Then
        intMsgCount = UBound(vntAddMsg) + 1
    Else
        intMsgCount = 0
    End If
    
    '' SQLを作成する
    strSQL = "SELECT MSGNAIYO,MSGLEVEL FROM THJMSG WHERE MSGCD = '" & UCase(Trim(strMsgCD)) & "'"

    '' ﾀﾞｲﾅｾｯﾄ生成
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)

    If oDynaset.EOF = False Then
        strMessage = GF_VarToStr(oDynaset![MSGNAIYO])
        strMsgLevel = GF_VarToStr(oDynaset![MSGLEVEL])
        
        '付加文字列に変更
        If intMsgCount > 0 Then
        
            '配列時
            For i = 1 To intMsgCount
                strTemp = "%" & CStr(i)
                strMessage = Replace(strMessage, strTemp, vntAddMsg(i - 1))
            Next
            
        ElseIf (Len(Trim(vntAddMsg)) > 0) Then
            '配列以外
            strTemp = "%1"
            strMessage = Replace(strMessage, strTemp, vntAddMsg)
        End If
        
    Else
        strMessage = " MsgErr[" & strMsgCD & "] "
        strMsgLevel = ""
    End If

    Set oDynaset = Nothing
    
    '' ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示
    If blnDispFlg = True Then
    
'2004/08/24 Update by N.Kigaku
''ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示でｱｲｺﾝを変更できるように修正。引数にstrICONを追加。
        strICON = UCase(strICON)
        If (strICON <> "C") And (strICON <> "Q") And (strICON <> "I") And (strICON <> "E") Then
            strICON = "E"
        End If
    
        '' "@@"を改行に置き換える
        strCnvMsg = GF_CnvCtrChar(strMessage)
        
        If Forms.Count > 0 Then
            'ﾌｫｰﾑがある時はﾌｫｰﾑのCaptionを表示する
            intRet = GF_MsgBox(Screen.ActiveForm.Caption, strCnvMsg, "OK", strICON)
        Else
            'ｱﾌﾟﾘｹｰｼｮﾝﾀｲﾄﾙを表示する
            intRet = GF_MsgBox(App.Title, strCnvMsg, "OK", strICON)
        End If
    End If

    '' ﾛｸﾞﾃｰﾌﾞﾙ出力
    If blnLogTblFlg = True Then
        intRet = GF_WriteLogData(mstrUserID, mstrPGMCD, strMsgCD, strMsgLevel, strMessage, strInfo, mstrTerminalCD)
    End If
    
    '' TIVOLI用ﾛｸﾞﾌｧｲﾙ出力
    If (blnTivoliLogFlg = True) And ((strMsgLevel = "1") Or (strMsgLevel = "2")) Then
        intRet = GF_LogOut("VB6", "GF_GetMsg_Addition", strMsgCD, GF_DelChrCode(strMessage) & IIf(Len(Trim(strInfo)) = 0, "", IIf(Len(Trim(strInfo)) > 0, vbCrLf & Space(4) & strInfo, "")), 1, strMsgLevel)
    End If
    
    GF_GetMsg_Addition = strMessage
    
    Exit Function
    
ErrHandler:
    
    Call GS_ErrorHandler("GF_GetMsg_Addition", strSQL)
    
    GF_GetMsg_Addition = " MsgErr[" & strMsgCD & "] "

End Function

Public Function GF_VarToStr(vVALUE As Variant) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　変換処理（VARIANT -> STRING）
' 機能　　　:
' 引数　　　:　[I] vVALUE As Variant    ''変換前ﾊﾞﾘｱﾝﾄﾃﾞｰﾀ
' 戻り値　　:　変換後文字列
' 機能説明　:　引数がNULLの場合は、""(長さ0の文字列)を返す
'------------------------------------------------------------------------------
    
    Dim mstrTEMP As String

    On Error GoTo ErrHandler
    
    If (IsNull(vVALUE) = True) Then
      mstrTEMP = ""
    Else
      mstrTEMP = CStr(vVALUE)
    End If
    
    GF_VarToStr = mstrTEMP

    Exit Function

ErrHandler:
    
    Call GS_ErrorHandler("GF_VarToStr", "")
    
    GF_VarToStr = ""

End Function

Public Function GF_VarToNum(vVALUE As Variant) As Double
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　変換処理（VARIANT -> DOUBLE）
' 機能　　　:
' 引数　　　:　[I] vVALUE As Variant    ''変換前ﾊﾞﾘｱﾝﾄﾃﾞｰﾀ
' 戻り値　　:　変換後数値
' 機能説明　:　引数がNULLの場合は、0を返す
'------------------------------------------------------------------------------
    
    Dim mdblTEMP As Double

    If (IsNull(vVALUE) = True Or IsNumeric(vVALUE) = False) Then
        mdblTEMP = 0#
    Else
        mdblTEMP = CDbl(vVALUE)
    End If
    
    GF_VarToNum = mdblTEMP

    Exit Function

ErrHandler:
    
    Call GS_ErrorHandler("GF_VarToNum", "")
    
    GF_VarToNum = 0#

End Function

Public Function GF_CnvCtrChar(sMSG As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　改行指示文字変換処理（@@ -> vbCrLf）
' 機能　　　:
' 引数　　　:　[I] sMSG As String       ''変換前文字列
' 戻り値　　:　成功 - 変換後文字列
' 　　　　　:　失敗 - 変換前文字列
' 機能説明　:
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    ''文字列置換処理
    GF_CnvCtrChar = Replace(sMSG, "@@", vbCrLf)
    
    Exit Function

ErrHandler:
    
    Call GS_ErrorHandler("GF_CnvCtrChar", "")
    
    GF_CnvCtrChar = sMSG

End Function

Public Function GF_DelCtrChar(sMSG As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　改行指示文字削除処理（@@ -> DELETE）
' 機能　　　:
' 引数　　　:　[I] sMSG As String       ''変換前文字列 / 変換後文字列
' 戻り値　　:　True = 成功 / False = 失敗
' 機能説明　:
'------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ''文字列置換処理
    GF_DelCtrChar = Replace(sMSG, "@@", "")
    
    Exit Function

ErrHandler:
    
    Call GS_ErrorHandler("GF_DelCtrChar", "")
    
    GF_DelCtrChar = sMSG

End Function

Public Sub GS_ErrorHandler(strLocation As String, _
                  Optional strAdditon As String = "", _
                  Optional intDefSqlCdErr As Integer = 0, _
                  Optional strMsgCD As String = mstrCommonErrMsgCD)
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ｴﾗｰﾊﾝﾄﾞﾗ
' 機能　　　:　ｴﾗｰ時の処理を行う
' 引数　　　:　[I] strLocation As String           ''ｴﾗｰ発生場所
' 　　　　　 　[I] strAdditon As String            ''付加情報
'             [I] intDefSqlCdErr As Integer       ''SQLｺｰﾄﾞﾊﾟﾗﾒｰﾀのﾃﾞﾌｫﾙﾄｴﾗｰｺｰﾄﾞ判定 ﾃﾞﾌｫﾙﾄ:0
' 　　　　　 　[I] strMsgCD As String              ''ﾒｯｾｰｼﾞｺｰﾄﾞ
' 戻り値　　:　なし
' 機能説明　:
'------------------------------------------------------------------------------
    
    Dim intRet      As Integer
    Dim lngErrNum   As Long
    Dim strErrMsg   As String
    Dim strErrType  As String
    Dim blnErrFlg   As Boolean
    Dim strMsgLevel As String
    Dim strMsg      As String

    blnErrFlg = False
    strMsg = ""
    strMsgLevel = "1"   '致命的ｴﾗｰを設定

    '=== ORACLE SCRIPT ERROR ===
    If gOraParam!sql_code <> intDefSqlCdErr Then
    
        lngErrNum = gOraParam!sql_code
        strErrMsg = gOraParam!sql_errm
        strErrType = "ORACLE"
        
        gOraParam!sql_code = 0
        gOraParam!sql_errm = ""
        
        blnErrFlg = True

    '=== VB ERROR ===
    ElseIf gOraSession.LastServerErr = 0 And gOraDataBase.LastServerErr = 0 Then
    
        If Err.Number <> 0 Then
        
            lngErrNum = Err.Number
            strErrMsg = Err.Description
            strErrType = "VB6"
        
            Err.Clear
            
            blnErrFlg = True
        End If
        
    '=== ORACLE DATABASE ERROR ===
    ElseIf gOraDataBase.LastServerErr <> 0 Then
    
        lngErrNum = gOraDataBase.LastServerErr
        strErrMsg = gOraDataBase.LastServerErrText
        strErrType = "ORACLE"
        
        gOraDataBase.LastServerErrReset
        
        blnErrFlg = True
    
    '=== ORACLE SESSION ERROR ===
    ElseIf gOraSession.LastServerErr <> 0 Then
        
        lngErrNum = gOraSession.LastServerErr
        strErrMsg = gOraSession.LastServerErrText
        strErrType = "ORACLE"

        gOraSession.LastServerErrReset
        
        blnErrFlg = True
        
    End If

    If blnErrFlg = True Then
    
'2003/08/05 修正
        'ﾒｯｾｰｼﾞ、ﾒｯｾｰｼﾞﾚﾍﾞﾙを取得
        If Len(Trim(strMsgCD)) > 0 Then
            intRet = GF_GetMsgInfo(strMsgCD, strMsg, strMsgLevel, True)
        End If
    
        'ﾛｸﾞ出力
        intRet = GF_LogOut(strErrType, strLocation, CStr(lngErrNum), GF_DelChrCode(strErrMsg) & IIf(Len(Trim(strAdditon)) = 0, "", IIf(Len(Trim(strAdditon)) > 0, vbCrLf & Space(4) & strAdditon, "")), 2, strMsgLevel)
        
        'ﾛｸﾞﾃｰﾌﾞﾙ出力
        intRet = GF_WriteLogData(mstrUserID, mstrPGMCD, strMsgCD, strMsgLevel, strMsg, strLocation & IIf(Len(Trim(strErrMsg)) > 0, " , ", "") & strErrMsg & IIf(Len(Trim(strAdditon)) > 0, " , " & strAdditon, ""), mstrTerminalCD)

        'ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞがTrueの時にﾒｯｾｰｼﾞﾎﾞｯｸｽを表示する
        If mblnMsgDispFlg = True Then
            intRet = GF_MsgBox("ERROR NO. " & lngErrNum & " - " & strLocation, strErrMsg & vbCrLf & strLocation, "OK", "E")
        End If
    End If
End Sub

Public Sub GS_ErrorClear(bVB As Boolean, _
                         Optional bParam As Boolean = False, _
                         Optional bDataBase As Boolean = False, _
                         Optional bSession As Boolean = False)
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ｴﾗｰｸﾘｱ
' 機能　　　:　ｴﾗｰをｸﾘｱする
' 引数　　　:　[I] bVB        VBｴﾗｰｴﾗｰｸﾘｱﾌﾗｸﾞ
'             [I] bParam     ﾊﾟﾗﾒｰﾀｴﾗｰｸﾘｱﾌﾗｸﾞ
'             [I] bDataBase  ﾃﾞｰﾀﾍﾞｰｽｴﾗｰｸﾘｱﾌﾗｸﾞ
'             [I] bSession   ｾｯｼｮﾝｴﾗｰｸﾘｱﾌﾗｸﾞ
'
' 戻り値　　:　なし
' 機能説明　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    If bVB = True Then
        Err.Clear
    End If
    If bParam = True Then
        gOraParam!sql_code = 0
        gOraParam!sql_errm = ""
    End If
    If bDataBase = True Then
        gOraDataBase.LastServerErrReset
    End If
    If bSession = True Then
        gOraSession.LastServerErrReset
    End If
    
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("GS_ErrorClear", "")
End Sub

Public Function GF_LogOut(strErrType As String, _
                          strLocation As String, _
                          strErrNum As String, _
                          strErrMsg As String, _
                 Optional intTivoliLogKbn As Integer = 0, _
                 Optional strErrMsgLvl As String = "") As Integer
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾛｸﾞ出力（ﾃｷｽﾄ版）
' 機能　　　:　ｴﾗｰ時の状況をﾃｷｽﾄﾌｧｲﾙへ出力する
' 引数　　　:　[I] strErrType  As String        ''ｴﾗｰﾀｲﾌﾟ(VB or ORACLE)
' 　　　　　 　[I] strLocation As String        ''ｴﾗｰ発生場所
'             [I] strErrNum As String          ''ｴﾗｰ番号 or ｴﾗｰｺｰﾄﾞ
' 　　　　　 　[I] strErrMSG As String          ''ｴﾗｰﾒｯｾｰｼﾞ
'             [I] intTivoliLogKbn As Integer   ''Tivoliﾛｸﾞ出力区分
'             [I] strErrMsgLvl As String       ''ｴﾗｰﾒｯｾｰｼﾞﾚﾍﾞﾙ
'                                                 0:通常出力, 1:Tivoliﾛｸﾞ出力, 2:両方
' 機能説明　:
' 備考　　　:  ﾛｸﾞﾌｧｲﾙ名をLogFileﾌﾟﾛﾊﾟﾃｨにｾｯﾄしておくこと
'------------------------------------------------------------------------------

    Dim intRet     As Integer
    Dim strYmd     As String
    Dim strHms     As String
    Dim strMsg     As String
    Dim FileID     As Integer

    On Error GoTo ErrHandler
    
    '出力ﾒｯｾｰｼﾞの編集
    strYmd = Format(Date, "yyyy/mm/dd")
    strHms = Format(Time, "hh:mm:ss")

    'ﾛｸﾞ出力ﾒｯｾｰｼﾞの編集
    strMsg = strYmd & " " & strHms & ", " & gstrUserID & ", " & App.EXEName & vbCr & _
                        strErrType & ", " & strLocation & ", " & "[" & strErrNum & "] " & strErrMsg
    
    If (intTivoliLogKbn = 0) Or (intTivoliLogKbn = 2) Then
    
        If Len(Trim(mstrLogFile)) = 0 Then
            Err.Raise Number:=vbObjectError, Description:="ﾌｧｲﾙ名が設定されていません。"
        End If
    
        'ﾃﾞｨﾚｸﾄﾘ存在ﾁｪｯｸ
        If GF_DirCheck(mstrLogFile) = False Then
            Err.Raise Number:=vbObjectError, Description:="出力先が存在しません。" & "[" & mstrLogFile & "]"
        End If
        
        FileID = FreeFile
        
        Open mstrLogFile For Append As #FileID
        Print #FileID, strMsg
        Close #FileID
        
    End If
    
    'TIVOLI用ﾛｸﾞ出力
    If basMainFunc.TVL_LOG_FLG = True Then
        If (intTivoliLogKbn = 1) Or (intTivoliLogKbn = 2) Then
        
            'ﾃﾞｨﾚｸﾄﾘ存在ﾁｪｯｸ
            If GF_DirCheck(basMainFunc.TVL_LOG_DIR) = False Then
                Err.Raise Number:=vbObjectError, Description:="出力先が存在しません。" & "[" & basMainFunc.TVL_LOG_DIR & "]"
            End If
        
            If (strErrMsgLvl = "1") Or (strErrMsgLvl = "2") Then
        
                FileID = FreeFile
            
                If strErrMsgLvl = "1" Then
                    'ｴﾗｰﾛｸﾞ
                    Open basMainFunc.TVL_LOG_DIR & App.EXEName & "_ERRO.ERR" For Append As #FileID
                ElseIf strErrMsgLvl = "2" Then
                    'ﾜｰﾆﾝｸﾞﾛｸﾞ
                    Open basMainFunc.TVL_LOG_DIR & App.EXEName & "_WARN.ERR" For Append As #FileID
                End If
                Print #FileID, strMsg
                Close #FileID
                
            End If
        End If
    End If
    
    GF_LogOut = True
    
    Exit Function

ErrHandler:
    'ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞがTrueの時はﾒｯｾｰｼﾞﾎﾞｯｸｽを表示
    If mblnMsgDispFlg = True Then
        intRet = GF_MsgBox("ERROR NO. " & Err.Number & " - GF_LogOut", Err.Description, "OK", "E")
    End If
    GF_LogOut = False

End Function

Public Function GF_LogOutDB(strErrType As String, strLocation As String, strMsgID As String) As Integer
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾛｸﾞ出力（ﾒｯｾｰｼﾞｺｰﾄﾞ版）
' 機能　　　:　ﾒｯｾｰｼﾞｺｰﾄﾞよりﾒｯｾｰｼﾞをDBから取得し、ﾃｷｽﾄﾌｧｲﾙへ出力する
' 引数　　　:　[I] strErrType  As String      ''ｴﾗｰﾀｲﾌﾟ(VB or ORACLE)
' 　　　　　 　[I] strLocation As String      ''ｴﾗｰ発生場所
' 　　　　　 　[I] strMsgID    As String      ''ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
' 戻り値　　:　なし
' 備考     ：
'------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Dim strMsg   As String      ''ﾒｯｾｰｼﾞ内容
    Dim intRet   As Integer
    
    ''ﾒｯｾｰｼﾞ取得
    strMsg = GF_GetMsg(strMsgID)
    If strMsg = "" Then
        'ﾒｯｾｰｼﾞ取得失敗
        strMsg = "ﾒｯｾｰｼﾞ取得不能"
    End If
        
    ''ﾛｸﾞ出力
    intRet = GF_LogOut(strErrType, strLocation, strMsgID, strMsg)
    
    Exit Function

ErrHandler:
    intRet = GF_MsgBox("ERROR NO. " & Err.Number & " - GF_LogOutDB", Err.Description, "OK", "E")
    GF_LogOutDB = False

End Function

Public Function GF_ExeLogOut(Optional strAddMsg As String = "") As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾌﾟﾛｸﾞﾗﾑ終了ﾛｸﾞ出力（ﾃｷｽﾄ版）
' 機能　　　:　ﾌﾟﾛｸﾞﾗﾑ終了時にﾛｸﾞを出力する。
' 引数　　　:  [I] strAddMsg As String   ''付加ﾒｯｾｰｼﾞ
' 機能説明　:
' 備考　　　:  ﾛｸﾞﾌｧｲﾙ名をLogFileﾌﾟﾛﾊﾟﾃｨにｾｯﾄしておくこと
'------------------------------------------------------------------------------

    Dim intRet     As Integer
    Dim strMsg     As String
    Dim FileID     As Integer

    On Error GoTo ErrHandler

    FileID = -1
        
    'TIVOLI用ﾛｸﾞ出力
    If basMainFunc.TVL_LOG_FLG = True Then
    
        'ﾛｸﾞ出力ﾒｯｾｰｼﾞの編集
        strMsg = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & ", " & gstrUserID & ", " & App.EXEName & IIf(strMsg = "", "", ", " & strAddMsg)
    
        'ﾃﾞｨﾚｸﾄﾘ存在ﾁｪｯｸ
        If GF_DirCheck(basMainFunc.TVL_LOG_DIR) = False Then
            Err.Raise Number:=vbObjectError, Description:="出力先が存在しません。" & "[" & basMainFunc.TVL_LOG_DIR & "]"
        End If
    
        FileID = FreeFile
    
        'ﾛｸﾞ出力
        Open basMainFunc.TVL_LOG_DIR & App.EXEName & "_EXE.LOG" For Append As #FileID
        Print #FileID, strMsg
        Close #FileID
                
    End If
    
    GF_ExeLogOut = True
    
    Exit Function

ErrHandler:
    'ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞがTrueの時はﾒｯｾｰｼﾞﾎﾞｯｸｽを表示
    If mblnMsgDispFlg = True Then
        intRet = GF_MsgBox("ERROR NO. " & Err.Number & " - GF_ExeLogOut", Err.Description, "OK", "E")
    End If
    
    On Error Resume Next
    If FileID > -1 Then
        Close #FileID
    End If
    
    GF_ExeLogOut = False

End Function


'Public Function GF_WriteLogData(ByVal strUserID As String, _
'                            ByVal strProgCD As String, _
'                            ByVal strMsgCD As String, _
'                            ByVal strMsgLevel As String, _
'                            ByVal strMsg As String, _
'                            ByVal strNote As String, _
'                   Optional ByVal strTerminalCd As String = "" _
'                            ) As Boolean
''--------------------------------------------------------------------------------
'' @(f)
''
'' 機能名　　: ﾛｸﾞﾃﾞｰﾀ出力
'' 機能　　　: ﾛｸﾞを国海)ﾛｸﾞﾃﾞｰﾀに書き出す
'' 引数　　　: [I] strUserID As String        ''ﾕｰｻﾞID
''　　　　　 : [I] strProgCD As String        ''ﾌﾟﾛｸﾞﾗﾑｺｰﾄﾞ
''　　　　　 : [I] strMsgCD As String         ''ﾒｯｾｰｼﾞｺｰﾄﾞ
''　　　　　 : [I] strMsgLevel As String      ''ﾒｯｾｰｼﾞﾚﾍﾞﾙ
''　　　　　 : [I] strMsg As String           ''ﾒｯｾｰｼﾞ
''　　　　　 : [I] strNote As String          ''付与情報
''　　　　　 : [I] strTerminalCD As String    ''端末ｺｰﾄﾞ　(省略時：空）
'' 戻り値　　:
'' 備考　　　:
''
''--------------------------------------------------------------------------------
'
'    Dim strSQL      As String
'    Dim intRet      As Integer
'    Dim strErrNum   As String
'    Dim strErrMsg   As String
'
'    On Error GoTo ErrHandler
'
'    '' Nullﾁｪｯｸ
'    If strUserID = "" Then strUserID = " "
'    If strProgCD = "" Then strProgCD = " "
'    If strMsgCD = "" Then strMsgCD = " "
'    If strMsgLevel = "" Then strMsgLevel = " "
'    If strMsg = "" Then strMsg = " "
'
'    '' シングルクォーテーションを置換する
'    strMsg = Replace(strMsg, "'", "''")
'    strNote = Replace(strNote, "'", "''")
'
'    '' ログテーブルに書き出す
'    strSQL = ""
'    strSQL = strSQL & "INSERT INTO T31_LOG_DATA ("
'    strSQL = strSQL & "  NSERIAL_NO,"
'    strSQL = strSQL & "  CUSER_ID,"
'    strSQL = strSQL & "  VCTERMINAL_CD,"
'    strSQL = strSQL & "  VCPGM_CD,"
'    strSQL = strSQL & "  DDATE,"
'    strSQL = strSQL & "  CMSG_CD,"
'    strSQL = strSQL & "  CMSG_LEVEL,"
'    strSQL = strSQL & "  VCMSG_CONTENTS,"
'    strSQL = strSQL & "  VCINVEST_INFO"
'    strSQL = strSQL & ") VALUES ("
'    strSQL = strSQL & " (SELECT NVL(MAX(NSERIAL_NO),0)+1 FROM T31_LOG_DATA),"
''2007/07/30 Updated by N.Kigaku Start --------------------
'    strSQL = strSQL & "  '" & Right(Trim(strUserID), 7) & "',"
''    strSQL = strSQL & "  '" & strUserID & "',"
''2007/07/30 Update End -----------------------------------
'    strSQL = strSQL & "  '" & strTerminalCd & "',"
'    strSQL = strSQL & "  '" & strProgCD & "',"
'    strSQL = strSQL & "  SYSDATE,"
'    strSQL = strSQL & "  '" & strMsgCD & "',"
'    strSQL = strSQL & "  '" & strMsgLevel & "',"
'    strSQL = strSQL & "  '" & strMsg & "',"
'    strSQL = strSQL & "  '" & strNote & "'"
'    strSQL = strSQL & ")"
'
'    '' SQLを実行する
'    Call gWOraDataBase.ExecuteSQL(strSQL)
'
'    GF_WriteLogData = True
'
'    Exit Function
'
'ErrHandler:
'    If gWOraSession.LastServerErr <> 0 Then
'
'        strErrNum = gWOraSession.LastServerErr
'        strErrMsg = gWOraSession.LastServerErrText
'
'        gWOraSession.LastServerErrReset
'
'    ElseIf gWOraDataBase.LastServerErr <> 0 Then
'
'        strErrNum = gWOraDataBase.LastServerErr
'        strErrMsg = gWOraDataBase.LastServerErrText
'
'        gWOraDataBase.LastServerErrReset
'
'    ElseIf Err.Number <> 0 Then
'
'        strErrNum = Err.Number
'        strErrMsg = Err.Description
'
'        Err.Clear
'
'    End If
'
'    'ﾛｸﾞﾌｧｲﾙ出力
'    Call GF_LogOut("VB6", "GF_WriteLogData", strErrNum, strErrMsg)
'
'    'ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞがTrueの時はﾒｯｾｰｼﾞﾎﾞｯｸｽを表示
'    If mblnMsgDispFlg = True Then
'        intRet = GF_MsgBox("ERROR NO. " & strErrNum & " - GF_WriteLogData", strErrMsg, "OK", "E")
'    End If
'    GF_WriteLogData = False
'
'End Function

Public Function GF_WriteLogData(ByVal strUserID As String, _
                                ByVal strProgCD As String, _
                                ByVal strMsgCD As String, _
                                ByVal strMsgLevel As String, _
                                ByVal strMsg As String, _
                                ByVal strNote As String, _
                       Optional ByVal strTerminalCd As String = "" _
                               ) As Boolean
'--------------------------------------------------------------------------------
' @(f)
'
' 機能名　　: ﾛｸﾞﾃﾞｰﾀ出力
' 機能　　　: ﾛｸﾞを国海)ﾛｸﾞﾃﾞｰﾀに書き出す
' 引数　　　: [I] strUserID As String        ''ﾕｰｻﾞID
'　　　　　 : [I] strProgCD As String        ''ﾌﾟﾛｸﾞﾗﾑｺｰﾄﾞ
'　　　　　 : [I] strMsgCD As String         ''ﾒｯｾｰｼﾞｺｰﾄﾞ
'　　　　　 : [I] strMsgLevel As String      ''ﾒｯｾｰｼﾞﾚﾍﾞﾙ
'　　　　　 : [I] strMsg As String           ''ﾒｯｾｰｼﾞ
'　　　　　 : [I] strNote As String          ''付与情報
'　　　　　 : [I] strTerminalCD As String    ''端末ｺｰﾄﾞ　(省略時：空）
' 戻り値　　:
' 備考　　　:
'--------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL              As String
    Dim strErrMsg           As String           'ﾁｪｯｸ結果ﾒｯｾｰｼﾞ
    Dim lngErrNum           As Long             'ｴﾗｰNo
    Dim lclsOraClass        As New clsOraClass  'ｽﾄｱﾄﾞ呼出し用
    Dim blnCreOraClass      As Boolean          'ｸﾗｽｵﾌﾞｼﾞｪｸﾄ作成済ﾌﾗｸﾞ

    GF_WriteLogData = False

    ''Nullの場合、ブランク1桁を設定
    If strUserID = "" Then strUserID = " "
    If strProgCD = "" Then strProgCD = " "
    If strMsgCD = "" Then strMsgCD = " "
    If strMsgLevel = "" Then strMsgLevel = " "
    If strMsg = "" Then strMsg = " "

    strErrMsg = ""

    'ｽﾄｱﾄﾞ用 Object 宣言
    Set lclsOraClass = New clsOraClass
    Set lclsOraClass.OraDataBase_Strcall = gOraDataBase
    blnCreOraClass = True

    'ｻｰﾊﾞｰｴﾗｰのﾘｾｯﾄ
    Call lclsOraClass.ErrReset_Strcall

    'ﾊﾞｲﾝﾄﾞ変数追加
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_USR_ID", Right(Trim(strUserID), 7)   'ﾕｰｻﾞID
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_TER_CD", strTerminalCd       '端末CD
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_PGM_CD", strProgCD           'ﾌﾟﾛｸﾞﾗﾑCD
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_MSG_CD", strMsgCD            'ﾒｯｾｰｼﾞCD
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_MSG_LV", strMsgLevel         'ﾒｯｾｰｼﾞﾚﾍﾞﾙ
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_MSG", strMsg                 'ﾒｯｾｰｼﾞ
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_INVEST_INFO", strNote        '付与情報
    lclsOraClass.Add_Binds ORAPARM_OUTPUT, ORATYPE_VARCHAR2, "p_ErrMsg", strErrMsg      'ｴﾗｰﾒｯｾｰｼﾞ

    If (lclsOraClass.ErrCode_Strcall <> 0 Or lclsOraClass.ErrText_Strcall <> "") Then
        Err.Raise Number:=lclsOraClass.ErrCode_Strcall, Description:=lclsOraClass.ErrText_Strcall
    End If

    'ｽﾄｱｰﾄﾞﾌｧﾝｸｼｮﾝへの引数ｾｯﾄ
    strSQL = ""
    strSQL = strSQL & "BEGIN "
    strSQL = strSQL & ":sql_code:=TOS_SE_LOGDATA_OUT"      'ｽﾄｱｰﾄﾞﾌｧﾝｸｼｮﾝ名
    strSQL = strSQL & " (:p_USR_ID"
    strSQL = strSQL & ", :p_TER_CD"
    strSQL = strSQL & ", :p_PGM_CD"
    strSQL = strSQL & ", :p_MSG_CD"
    strSQL = strSQL & ", :p_MSG_LV"
    strSQL = strSQL & ", :p_MSG"
    strSQL = strSQL & ", :p_INVEST_INFO"
    strSQL = strSQL & ", :p_ErrMsg"
    strSQL = strSQL & "); "
    strSQL = strSQL & "END;"

    'ｻｰﾊﾞｴﾗｰのﾘｾｯﾄ
    Call lclsOraClass.ErrReset_Strcall

    'SQLの実行
    lclsOraClass.ExecSql_Strcall strSQL

    'ﾁｪｯｸ結果ﾒｯｾｰｼﾞ取得
    strErrMsg = GF_VarToStr(gOraParam!p_ErrMsg)

    If (gOraParam!sql_code = -1) Then
        'ストアドシステムエラー
        Err.Raise Number:=vbObjectError, Description:=strErrMsg
    End If

    'ﾊﾞｲﾝﾄﾞﾊﾟﾗﾒｰﾀｰの全解放
    lclsOraClass.RemoveAll
    Set lclsOraClass = Nothing
    blnCreOraClass = False
    
    'ﾊﾟﾗﾒｰﾀの初期化
    Call GS_ErrorClear(False, True, True, False)

    GF_WriteLogData = True
    Exit Function

ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    
    lngErrNum = Err.Number
    strErrMsg = Err.Description
    
    'ﾃｷｽﾄﾛｸﾞ出力
    Call GF_LogOut("ORACLE", "GF_WriteLogData", CStr(lngErrNum), strErrMsg)
    
    'ﾒｯｾｰｼﾞﾎﾞｯｸｽ表示ﾌﾗｸﾞがTrueの時はﾒｯｾｰｼﾞﾎﾞｯｸｽを表示
    If mblnMsgDispFlg = True Then
        Call GF_MsgBox("ERROR NO. " & lngErrNum & " - GF_WriteLogData", strErrMsg, "OK", "E")
    End If
    
    'ｸﾗｽｵﾌﾞｼﾞｪｸﾄの解放
    If blnCreOraClass = True Then
        'ﾊﾞｲﾝﾄﾞﾊﾟﾗﾒｰﾀｰの全解放
        lclsOraClass.RemoveAll
        Set lclsOraClass = Nothing
    End If

End Function


'Public Sub GS_AppLog(strLogCode As String, strLogValue As String, strNote As String)
''------------------------------------------------------------------------------
'' @(f)
''
'' 機能名　　:　稼働ﾛｸﾞ出力
'' 機能　　　:　稼働状況をﾃﾞｰﾀﾍﾞｰｽへ出力する
'' 引数　　　:　sLogCode  As String      ''ﾛｸﾞｺｰﾄﾞ
'' 　　　　　 　sLogValue As String      ''ﾛｸﾞ内容
'' 　　　　　 　sNote     As String      ''備考
'' 戻り値　　:　なし
'' 機能説明　:　MLTT_028(ﾛｸﾞ情報ﾃｰﾌﾞﾙ)へ出力
''------------------------------------------------------------------------------
'
'    Dim oSTORED     As Object
'    Dim sSQL        As String
'    Dim strErrorMsg As String
'    Dim lngErrorNo  As Long
'    Dim intRet      As Integer
'
'    On Error GoTo ErrHandler
'
'    'ｽﾄｱﾄﾞ用 Object 宣言
'    Set oSTORED = New clsOraClass
'    Set oSTORED.OraDataBase_Strcall = gOraDataBase
'
'    'ｻｰﾊﾞｰｴﾗｰのﾘｾｯﾄ
'    Call oSTORED.ErrReset_Strcall
'
'    'ﾊﾞｲﾝﾄﾞ変数追加
'    oSTORED.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_SHAIN_CD", gstrUserID
'    oSTORED.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_SHOZOKU_CD", gstrGroupID
'    oSTORED.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_GYOMU_CD", gstrGyomuCode
'    oSTORED.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_SAGYO_CD", gstrSagyoCode
'    oSTORED.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_LOG_CD", strLogCode
'    oSTORED.Add_Binds ORAPARM_INPUT, ORATYPE_VARCHAR2, "p_LOGNAIYO", strLogValue
'    oSTORED.Add_Binds ORAPARM_INPUT, ORATYPE_VARCHAR2, "p_BIKO", strNote
'    oSTORED.Add_Binds ORAPARM_INPUT, ORATYPE_VARCHAR2, "p_UPD_MAN", gstrUserID
'
'    If (oSTORED.ErrCode_Strcall <> 0 Or oSTORED.ErrText_Strcall <> "") Then
'
'        'ﾊﾞｲﾝﾄﾞ変数の全解放
'        oSTORED.RemoveAll
'
'        Exit Sub
'    End If
'
'    'ｽﾄｱﾄﾞﾌﾟﾛｼｰｼﾞｬへの引数ｾｯﾄ
'    sSQL = "BEGIN "
'    sSQL = sSQL & "MIPC_001 ("          'ｽﾄｱﾄﾞﾌﾟﾛｼｰｼﾞｬ名
'    sSQL = sSQL & ":p_SHAIN_CD, "       'ﾕｰｻﾞID
'    sSQL = sSQL & ":p_SHOZOKU_CD, "     '所属ｸﾞﾙｰﾌﾟｺｰﾄﾞ
'    sSQL = sSQL & ":p_GYOMU_CD, "       '業務ｺｰﾄﾞ
'    sSQL = sSQL & ":p_SAGYO_CD, "       '作業ｺｰﾄﾞ
'    sSQL = sSQL & ":p_LOG_CD, "         'ﾛｸﾞｺｰﾄﾞ
'    sSQL = sSQL & ":p_LOGNAIYO, "       'ﾛｸﾞ内容
'    sSQL = sSQL & ":p_BIKO, "           '備考
'    sSQL = sSQL & ":p_UPD_MAN"          '更新者
'    sSQL = sSQL & "); "
'    sSQL = sSQL & "END;"
'
'    'ｻｰﾊﾞｴﾗｰのﾘｾｯﾄ
'    Call oSTORED.ErrReset_Strcall
'
'    'SQLの実行
'    oSTORED.ExecSql_Strcall sSQL
'
'    If (oSTORED.ErrCode_Strcall <> 0 Or oSTORED.ErrText_Strcall <> "") Then
'
'        'ﾊﾞｲﾝﾄﾞ変数の全解放
'        oSTORED.RemoveAll
'
'        Exit Sub
'    End If
'
'    'バインドパラメータの全解放
'    oSTORED.RemoveAll
'
'    Exit Sub
'
'ErrHandler:
'
'    Call GS_ErrorHandler("GS_AppLog", sSQL)
'
'End Sub

Public Function GF_DelChrCode(sMSG As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾗｲﾝﾌｨｰﾄﾞ文字削除(ORACLEｴﾗｰﾒｯｾｰｼﾞ用)
' 機能　　　:　ORACLEからのｴﾗｰﾒｯｾｰｼﾞの最後尾に付加されているﾗｲﾝﾌｨｰﾄﾞ文字を削除
' 引数　　　:　[I] sMSG As String       ''ﾗｲﾝﾌｨｰﾄﾞ文字削除前文字列
' 戻り値　　:　ﾗｲﾝﾌｨｰﾄﾞ文字削除後文字列
' 機能説明　:
'------------------------------------------------------------------------------
    Dim mstrTEMP As String
    Dim i As Integer

    i = InStrB(1, sMSG, Chr(10), vbTextCompare)

    If (i <> 0) Then
        mstrTEMP = MidB(sMSG, 1, i - 1)
        GF_DelChrCode = mstrTEMP
    Else
        GF_DelChrCode = sMSG
    End If

End Function

Public Function GF_DirCheck(strPath As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾃﾞｨﾚｸﾄﾘ作成
' 機能　　　:　ﾃﾞｨﾚｸﾄﾘが存在するかﾁｪｯｸし、ない場合は作成する
' 引数　　　:　[I] strPath As String       ''ﾊﾟｽ
' 戻り値　　:　True = 成功 / False = 失敗
' 備考     ：
'------------------------------------------------------------------------------
    
    Dim strBuf   As String
    Dim strBuf2  As String
    Dim intLen   As Integer
    Dim intCount As Integer
    Dim bolNetWk As Boolean
    Dim intRoot  As Integer
    
    On Error GoTo ErrDirCheck
    
    GF_DirCheck = False
    
    intLen = Len(strPath)
    bolNetWk = False
    intRoot = 0
    For intCount = 1 To intLen
        'If intCount = 1 And InStr(1, strPath, ":", vbTextCompare) = 0 Then
'        If (intCount = 1) And (InStr(1, strPath, ":", vbTextCompare) = 0) And (Left(strPath, 1) <> ".") Then
        If (intCount = 1) And (InStr(1, strPath, ":", vbBinaryCompare) = 0) And (Left(strPath, 1) <> ".") Then
            strBuf2 = Mid$(strPath, intCount, 2)
            strBuf = strBuf & strBuf2
            intCount = intCount + 2
            bolNetWk = True
        End If
        strBuf2 = Mid$(strPath, intCount, 1)
        strBuf = strBuf & strBuf2
        If strBuf2 = "\" Or strBuf2 = "/" Then
            If bolNetWk = False Or (bolNetWk = True And intRoot > 0) Then
                If Dir(strBuf, vbDirectory) = "" Then
                    MkDir strBuf
                End If
            End If
            intRoot = intRoot + 1
        End If
    Next intCount
    
    GF_DirCheck = True
    
    Exit Function
    
ErrDirCheck:
    'ｴﾗｰ処理へ
    Call GS_ErrorHandler("GF_DirCheck")

End Function

'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :  スペースのNULL変換
' 機能      :  文字列が空白のときにNULL文字列を返す
'              空白でないときはTRIMして、シングルクォーテーションで囲んだ文字列を返す
' 引数     ： [I] strChar   As String    判定する文字列
' 戻り値    :  文字列
' 備考      :
'------------------------------------------------------------------------------
Public Function GF_ChangeSpaceToNull(strChar As String) As String
    GF_ChangeSpaceToNull = IIf(Len(Trim(strChar)) = 0, "IS NULL", "='" & Trim(strChar) & "'")
End Function

'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :  スペースのNULL変換(数値用)
' 機能      :  文字列が空白のときにNULL文字列を返す
'              空白でないときはもと文字列を戻す
' 引数     ： [I] strChar   As String    判定する文字列
'          :  [I] blnEqualFlg As Boolea イコール付加フラグ
' 戻り値    :  文字列
' 備考      :
'------------------------------------------------------------------------------
Public Function GF_ChangeNumSpaceToNull(strChar As String, Optional blnEqualFlg As Boolean = False) As String
    If blnEqualFlg = False Then
        GF_ChangeNumSpaceToNull = IIf(Len(Trim(strChar)) = 0, "NULL", Trim(strChar))
    Else
        GF_ChangeNumSpaceToNull = IIf(Len(Trim(strChar)) = 0, "= NULL", "=" & Trim(strChar))
    End If
End Function

'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　: スペースのNULL変換
' 機能　　　: 文字列が空白のときにNULL文字列を返す
'            空白でないときはTRIMして、シングルクォーテーションで囲んだ文字列を返す
' 引数　　　: [I] strChar   As String    判定する文字列
' 　　　　　: [I] blnEqualFlg As Boolean イコール有り無しフラグ（省略時：有り）
' 戻り値　　:  文字列
' 備考　　　:
'------------------------------------------------------------------------------
Public Function GF_ChangeSpaceToNull2(strChar As String, Optional blnEqualFlg As Boolean = True) As String
    If blnEqualFlg = True Then
        GF_ChangeSpaceToNull2 = IIf(Len(Trim(strChar)) = 0, "=NULL", "='" & Trim(strChar) & "'")
    Else
        GF_ChangeSpaceToNull2 = IIf(Len(Trim(strChar)) = 0, "NULL", "'" & Trim(strChar) & "'")
    End If
End Function

'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:
' 機能　　　: 文字列を指定した文字で囲む
' 引数　　　: [I] strChar   As String       囲む文字列
' 　　　　　: [I] strEnclose As Boolean     囲む文字
' 　　　　　: [I] blnEncloseFlg As Boolean  囲む文字列が空の時に囲うか否か
'                                         False :囲わない
'                                         True  :囲う
' 戻り値　　: 囲んだ文字列
' 備考　　　:
'------------------------------------------------------------------------------
Public Function GF_Enclose(strChar As String, strEnclose As String, Optional blnEncloseFlg As Boolean = False) As String
    If (Len(strChar) = 0) And (blnEncloseFlg = False) Then
        GF_Enclose = strChar
    Else
        GF_Enclose = strEnclose & strChar & strEnclose
    End If
End Function


Public Function GF_GetMsg_MasterMente(ByVal strMsgCD As String, _
                                            lngMaxRow As Long, _
                             Optional ByVal vntAddMsg As Variant = "") As String
'--------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　ﾏｽﾀﾒﾝﾃﾒｯｾｰｼﾞ取得（付加文字列付き）
' 機能　　　:　ﾒｯｾｰｼﾞをﾃﾞｰﾀﾍﾞｰｽから取得する
'
' 引数　　　: [I] strMsgCD  As String               ''ﾒｯｾｰｼﾞｺｰﾄﾞ
'　　　　　　 [I] lngMaxRow As Long                 ''配列数
'　　　　　　 [I] vntAddMsg As Variant              ''付加文字列の配列　（添字は0から開始）
'                                                    1件の時は配列でなくてもOK
' 戻り値　　: 取得したﾒｯｾｰｼﾞ
'　　　　　　 該当ﾃﾞｰﾀがない又はｴﾗｰの時は "MsgErr[ﾒｯｾｰｼﾞｺｰﾄﾞ]"
' 機能説明　: DBから取得したﾒｯｾｰｼﾞの%1～%nまで付加文字列で置き換える
'--------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim lngCounter  As Long
    Dim strExistFlg As String
    Dim strTemp     As String
    Dim strMessage  As String
    Dim intMsgCount As Integer
    Dim i           As Integer
    Dim intRet      As Integer
    Dim strMsgLevel As String
    
    ''戻り値設定
    GF_GetMsg_MasterMente = ""
    
    ''変数初期化
    strExistFlg = ""
    
    ''入力ﾊﾟﾗﾒｰﾀの配列の数を数える
    If IsArray(vntAddMsg) = True Then
        intMsgCount = UBound(vntAddMsg) + 1
    Else
        intMsgCount = 0
    End If
    
    ''保持配列に該当のMSGｺｰﾄﾞが存在するかﾁｪｯｸ
    lngCounter = 1
    Do Until lngCounter > lngMaxRow
        If Trim(ADD_TYPE_MSG(lngCounter).TYPE_MSG_CD) = Trim(strMsgCD) Then
            strMessage = Trim(ADD_TYPE_MSG(lngCounter).TYPE_MSG_NAIYO)
            strExistFlg = "1"
            Exit Do
        End If
        lngCounter = lngCounter + 1
    Loop
    
    ''該当MSGが存在しない場合は共通関数より取得する
    If Trim(strExistFlg) = "" Then
        ''MSG再検索
        If GF_GetMsgInfo(strMsgCD, strMessage, strMsgLevel) = False Then
            GoTo ErrHandler
        End If
        ''配列再定義
        lngMaxRow = lngMaxRow + 1
        ReDim ADD_TYPE_MSG(lngMaxRow)
        ''配列ﾃﾞｰﾀｾｯﾄ
        ADD_TYPE_MSG(lngMaxRow).TYPE_MSG_CD = Trim(strMsgCD)
        ADD_TYPE_MSG(lngMaxRow).TYPE_MSG_NAIYO = Trim(strMessage)
    End If
    
    ''MSGｾｯﾄ
    If intMsgCount > 0 Then
        '配列時
        For i = 1 To intMsgCount
            strTemp = "%" & CStr(i)
            strMessage = Replace(strMessage, strTemp, vntAddMsg(i - 1))
        Next
        
    ElseIf (Len(Trim(vntAddMsg)) > 0) Then
        '配列以外
        strTemp = "%1"
        strMessage = Replace(strMessage, strTemp, vntAddMsg)
    End If
    
    ''戻り値再設定
    GF_GetMsg_MasterMente = strMessage
    
    Exit Function
    
ErrHandler:
    
    Call GS_ErrorHandler("GF_GetMsg_MasterMente")

End Function


