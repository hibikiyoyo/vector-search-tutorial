Attribute VB_Name = "basDBFunc"
' @(h) DBFunc.bas  ver1.00 ( 2000/08/30 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   プロジェクト名  : TLFﾌﾟﾛｼﾞｪｸﾄ
'   モジュール名    : basDBFunc
'   ファイル名      : DBFunc.bas
'   Version        : 1.00
'   機能説明       ： DB接続に関する共通関数
'   作成者         ： T.Fukutani
'   作成日         ： 2000/08/30
'   修正履歴       ： 2007/10/30 N.Kigaku DB切断時にｵﾗｸﾙﾊﾞｲﾝﾄﾞ変数を破棄するように修正。
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 環境宣言
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' パブリック変数宣言
'------------------------------------------------------------------------------
'' 共通用
Public gOraSession          As OraSession       ''セッション定義
Public gOraDataBase         As OraDatabase      ''データベース定義
Public gOraParam            As OraParameters    ''パラメータオブジェクト

'' ログ出力用
Public gWOraSession         As OraSession       ''セッション定義
Public gWOraDataBase        As OraDatabase      ''データベース定義
'Public gWOraParam           As OraParameters    ''パラメータオブジェクト


Public Function GF_DBOpen(strInstance As String, strUserID As String, strPassWord As String, _
                 Optional blnWOraSessionConnectFlag As Boolean = True) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : DB接続
' 機能   : サーバー側オラクルとのセッション確立
' 引数   : strInstance As String    ''DB名
'          strUserID   As String    ''DB接続ﾕｰｻﾞ名
'          strPassWord As String    ''DB接続ﾊﾟｽﾜｰﾄﾞ
'          blnWOraSessionConnectFlag As Boolean    ''ﾛｸﾞ出力用OracleDB接続ﾌﾗｸﾞ(TRUE:接続、FALSE:非接続)
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    GF_DBOpen = False
   
    ''Oracleデータベースの接続・オープン
    Set gOraSession = CreateObject("OracleInProcServer.XOraSession")
    
    Set gOraDataBase = gOraSession.DbOpenDatabase(strInstance, _
                                                    strUserID _
                                                    & "/" & _
                                                    strPassWord, 0&)

    Set gOraParam = gOraDataBase.Parameters
    
    ''ｴﾗｰｺｰﾄﾞ、ｴﾗｰﾒｯｾｰｼﾞ取得用ﾊﾟﾗﾒｰﾀ設定
    gOraParam.Add "sql_code", 0, ORAPARM_OUTPUT
    gOraParam!sql_code.serverType = ORATYPE_NUMBER
    gOraParam.Add "sql_errm", "", ORAPARM_OUTPUT
    gOraParam!sql_errm.serverType = ORATYPE_VARCHAR2
    
    If blnWOraSessionConnectFlag = True Then
        ''ログ出力用Oracleデータベースの接続・オープン
        Set gWOraSession = CreateObject("OracleInProcServer.XOraSession")
        
        Set gWOraDataBase = gWOraSession.DbOpenDatabase(strInstance, _
                                                       strUserID _
                                                        & "/" & _
                                                        strPassWord, 0&)
    End If
    
    GF_DBOpen = True
    
    Exit Function

'エラー発生時、オラクル側でエラー情報を返せる場合はそれを参照する
'引数が有効で無い（VB ErrorNumber=440、Set OraDatabase が実行
'出来ない）場合は抜け出る
ErrHandler:
    Dim intRet     As Integer
    Dim lngErrNum  As Long      ''ｴﾗｰﾅﾝﾊﾞｰ
    Dim strErrMsg  As String    ''ｴﾗｰﾒｯｾｰｼﾞ
    Dim strErrType As String    ''ｴﾗｰﾀｲﾌﾟ
    
    strErrType = "ORACLE"
    If Err.Number = 429 Or Err.Number = 440 Then
        lngErrNum = Err.Number
        strErrMsg = "DBへの接続に失敗しました。"
    Else
        '=== ORACLE SESSION EEROR ===
        If gOraSession.LastServerErr <> 0 Then
            
            lngErrNum = gOraSession.LastServerErr
            strErrMsg = gOraSession.LastServerErrText
        
            gOraSession.LastServerErrReset
            
        '=== ORACLE DATABASE ERROR ===
        ElseIf gOraDataBase.LastServerErr <> 0 Then
        
            lngErrNum = gOraDataBase.LastServerErr
            strErrMsg = gOraDataBase.LastServerErrText
            
            gOraDataBase.LastServerErrReset

        Else

            If blnWOraSessionConnectFlag = True Then
                '=== ORACLE DATABASE ERROR ===
                If gWOraSession.LastServerErr <> 0 Then
                
                    lngErrNum = gWOraSession.LastServerErr
                    strErrMsg = gWOraSession.LastServerErrText
                
                    gWOraSession.LastServerErrReset
                
                '=== ORACLE SESSION EEROR ===
                ElseIf gWOraDataBase.LastServerErr <> 0 Then
                
                    lngErrNum = gWOraDataBase.LastServerErr
                    strErrMsg = gWOraDataBase.LastServerErrText
                    
                    gWOraDataBase.LastServerErrReset
                End If
            End If
            
        End If
    End If

    If basMsgFunc.DispErrMsgFlg = True Then
        intRet = GF_MsgBox("ERROR NO. " & lngErrNum & " - GF_DBOpen", strErrMsg, "OK", "E")
    End If
    intRet = GF_LogOut(strErrType, "GF_DBOpen", CStr(lngErrNum), strErrMsg, 1, "1")
    
End Function

Public Sub GS_DBClose(Optional blnWOraSessionConnectFlag As Boolean = True)
'------------------------------------------------------------------------------
' @(f)
' 機能名 : DB切断
' 機能   : サーバー側オラクルとのセッション終了
' 引数   :
' 備考   :
'------------------------------------------------------------------------------

    On Error GoTo ErrHandler

'2007/10/30 Added by N.Kigaku  Start -------
    'Oracleバインドパラメータ削除
    Call GF_RemoveAllBindParameter
'2007/10/30 Add End ------------------------

    ''Ｏｒａｃｌｅデータベースのクローズ・接続解除
    Set gOraParam = Nothing
    gOraDataBase.Close
    Set gOraDataBase = Nothing
    Set gOraSession = Nothing

    If blnWOraSessionConnectFlag = True Then
        ''ログ出力用Ｏｒａｃｌｅデータベースのクローズ・接続解除
        gWOraDataBase.Close
        Set gWOraDataBase = Nothing
        Set gWOraSession = Nothing
    End If

    Exit Sub

ErrHandler:
    Call GS_ErrorHandler("GS_DBClose")

End Sub

'2007/10/30 Added by N.Kigaku
Public Function GF_RemoveAllBindParameter() As Boolean
''--------------------------------------------------------------------------------
'' @(f)
'' 機能概要　:Oracleﾊﾞｲﾝﾄﾞ変数全削除
''
'' 引数　　　:
''
'' 戻り値　　:
''--------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim intCnt As Integer

    GF_RemoveAllBindParameter = False

    For intCnt = gOraParam.Count - 1 To 0 Step -1
        gOraParam.Remove intCnt
    Next intCnt

    GF_RemoveAllBindParameter = True

ErrHandler:
End Function
