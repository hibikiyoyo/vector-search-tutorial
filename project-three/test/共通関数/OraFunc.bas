Attribute VB_Name = "basOraFunc"
' @(h) OraFunc.bas  ver1.00 ( 2002/05/07 N.Kigaku )
'------------------------------------------------------------------------------
' @(s)
'　プロジェクト名　: TLFﾌﾟﾛｼﾞｪｸﾄ
'　モジュール名　　: basOraFunc
'　ファイル名　　　: OraFunc.bas
'　Version　　　　: 1.00
'　機能説明　　　　: オラクルのデータベースに関する共通関数
'　作成者　　　　　: N.Kigaku
'　作成日　　　　　: 2002/05/07
'　備考　　　　　　:
'　修正履歴　　　　: 2006/12/05 N.Kigaku ｵﾗｸﾙ8.1.7 Nocache対応 検索時、ReadOnlyからNocacheに変更
'                  : 2012/06/11 J.Yamaoka SQL文内のLikeエスケープ,リテラル置換追加
'　　　　　　　　　: 2015/08/21  NIC 王  LFDB更新 SESSIONID修正
'　　　　　　　　　: 2017/02/08 D.Ikeda K545 CSプロセス改善
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 環境宣言
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' パブリック変数宣言
'------------------------------------------------------------------------------

'Oracle ALL_COLL_TYPES Fields
Public Type OraAllCollType
    tOwner          As String     'ｵｰﾅ
    tTypeName       As String     'Type型名称
    tCollType       As String     'ｺﾚｸｼｮﾝﾀｲﾌﾟ
    tUpperBound     As String     '配列数
    tElemTypeName   As String     '型
    tLength         As String     '型のｻｲｽﾞ
End Type

''2012/06/11 J.Yamaoka Add
Private Const mstrLikeEscape As String = "\"        ''Like指定時のエスケープ文字



Public Function GF_GetSYSDATE(ByRef strSysDate As String, Optional ByVal intFormatKbn As Integer = 1) As Boolean
'------------------------------------------------------------------------------
' @(f)
'　機能名　: ＤＢサーバのシステム日付取得
'　機能　　: ＤＢサーバのシステム日付をオラクル経由で取得する
'　引数　　:　 strSysDate As String     (out)  システム日付
'　　　　　:　 intFormatKbn As Integer  (in)   書式区分
'                0: なし（オラクルの日付形式に基づく）
'                1: "yyyy/mm/dd hh24:mi:ss"        (省略時)
'                2: "yyyy/mm/dd"
'　戻り値　:　True = 成功 / False = 失敗
'　備考　　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim intRet            As Integer
    Dim strSQL            As String
    Dim oraDyna           As OraDynaset
    Dim strFormatSysdate  As String
    Dim strMsg            As String
    
    GF_GetSYSDATE = False
       
    Select Case intFormatKbn
    Case 0
        strFormatSysdate = "SYSDATE"
    Case 1
        strFormatSysdate = "TO_CHAR(SYSDATE,'YYYY/MM/DD HH24:MI:SS') ""SYSDATE"""
    Case 2
        strFormatSysdate = "TO_CHAR(SYSDATE,'YYYY/MM/DD') ""SYSDATE"""
    Case Else
        strFormatSysdate = "SYSDATE"
    End Select
       
    '引合情報OPTから今回使用する連番を取得する
    strSQL = ""
    strSQL = strSQL & "SELECT " & strFormatSysdate & " FROM DUAL"
'    strSQL = strSQL & "SELECT SYSDATE FROM DUAL"
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oraDyna.EOF Then
        'システム日付の取得に失敗しました。
        strMsg = GF_GetMsg("WTG027")
        Err.Raise Number:=vbObjectError, Description:=strMsg
    Else
        strSysDate = CStr(oraDyna![SYSDATE])
    End If
    Set oraDyna = Nothing
       
    GF_GetSYSDATE = True
    
    Exit Function
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_GetSYSDATE", strSQL)
End Function

Public Function GF_GetColumLength(ByRef intColumLen As Integer _
                                , ByVal strTBL_NAME As String _
                                , ByVal strCOLUM_NAME As String _
                                , Optional ByVal strOWNER As String = "LFSYS") As Boolean
'------------------------------------------------------------------------------
' @(f)
'　機能名　: フィールドサイズの取得
'　機能　　: 特定のテーブルのフィールドサイズを取得する
'　引数　　:　 intColumLen As Integer   (out)  フィールドサイズ
'　　　　　:　 strTBL_NAME As String    (in)   テーブル名
'　　　　　:　 strCOLUM_NAME As String  (in)   フィールド名
'　　　　　:　 strOWNER As String       (in)   所有者　（省略時:LFSYS）
'　戻り値　:　True = 成功 / False = 失敗
'　備考　　:  オラクル関数のVSIZEと同じ
'               例）SELECT VSIZE(SYNO) FROM THJMR
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL            As String
    Dim oraDyna           As OraDynaset
    Dim strMsg            As String
    
    GF_GetColumLength = False
       
    intColumLen = 0
       
    strSQL = ""
    strSQL = strSQL & "SELECT DATA_LENGTH FROM DBA_TAB_COLUMNS"
    strSQL = strSQL & " WHERE OWNER='" & strOWNER & "'"
    strSQL = strSQL & "   AND TABLE_NAME='" & strTBL_NAME & "'"
    strSQL = strSQL & "   AND COLUMN_NAME='"" & strCOLUM_NAME & " '"
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oraDyna.EOF Then
        strMsg = "テーブルのフィールドサイズの取得に失敗しました。"
        Err.Raise Number:=vbObjectError, Description:=strMsg
    Else
        intColumLen = CInt(oraDyna![DATA_LENGTH])
    End If
    Set oraDyna = Nothing
       
    GF_GetColumLength = True
    
    Exit Function
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_GetColumLength", strSQL)
End Function


Public Function GF_GetSessionID(ByRef lngSessionID As Long) As Boolean
'------------------------------------------------------------------------------
' @(f)
'　機能名　: セッションIDの取得
'　機能　　: セッションIDを取得する
'　引数　　:　 lngSessionID As Long   (out)  セッションID
'　戻り値　:　True = 成功 / False = 失敗
'　備考　　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL            As String
    Dim oraDyna           As OraDynaset
    Dim strMsg            As String
    
    GF_GetSessionID = False
       
    lngSessionID = 0
    '<LFDB更新 SESSIONID修正> del Start NIC 王  2015/08/21
    'strSQL = ""
    'strSQL = strSQL & "SELECT USERENV('SESSIONID') SESSIONID FROM DUAL"
    ''ﾀﾞｲﾅｾｯﾄの生成
    'Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    'If oraDyna.EOF Then
    '    strMsg = "セッションIDの取得に失敗しました。"
    '    Err.Raise Number:=vbObjectError, Description:=strMsg
    'Else
    '    lngSessionID = CLng(oraDyna![SESSIONID])
    'End If
    'Set oraDyna = Nothing
    '<LFDB更新 SESSIONID修正> del end NIC 王  2015/08/21
    '<LFDB更新 SESSIONID修正> ADD Start NIC 王  2015/08/21
    If GF_GetSessionID_Func(lngSessionID, strMsg) = False Then
        strMsg = "セッションIDの取得に失敗しました。"
        Err.Raise Number:=vbObjectError, Description:=strMsg
        Exit Function
        End If
    '<LFDB更新 SESSIONID修正> ADD end NIC 王  2015/08/21
    GF_GetSessionID = True
    
    Exit Function
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_GetSessionID", strSQL)
End Function

Public Sub Init_OraAllCollType(typCollType As OraAllCollType)
'------------------------------------------------------------------------------
' @(f)
'　機能名　: ALL_COLL_TYPES構造体の初期化
'　機能　　:
'　引数　　: typCollType As OraAllCollType   (in/out)  ALL_COLL_TYPES ﾃｰﾌﾞﾙのﾌｨｰﾙﾄﾞ
'　戻り値　: なし
'　備考　　:
'------------------------------------------------------------------------------
    With typCollType
        .tOwner = ""
        .tTypeName = ""
        .tCollType = ""
        .tUpperBound = ""
        .tElemTypeName = ""
        .tLength = ""
    End With
End Sub

Public Function GF_GetAllCollType(strTypeName As String, typCollType As OraAllCollType) As Boolean
'------------------------------------------------------------------------------
' @(f)
'　機能名　: Type型(ALL_COLL_TYPES)情報取得
'　機能　　: Type型(ALL_COLL_TYPES)の内容を取得する
'　引数　　: strTypeName As String            (in)  Type型名称
'            typCollType As OraAllCollType   (out)  ALL_COLL_TYPES ﾃｰﾌﾞﾙのﾌｨｰﾙﾄﾞ
'　戻り値　:　True = 成功 / False = 失敗
'　備考　　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strSQL    As String
    Dim oraDyna   As OraDynaset
    
    GF_GetAllCollType = False
    
    '構造体の初期化
    Call Init_OraAllCollType(typCollType)
    
    '型名称が空の時は正常終了とする
    If Len(Trim(strTypeName)) = 0 Then
        GF_GetAllCollType = True
        Exit Function
    End If
    
    'Type型の情報を取得する
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "       OWNER"
    strSQL = strSQL & "      ,TYPE_NAME"
    strSQL = strSQL & "      ,COLL_TYPE"
    strSQL = strSQL & "      ,UPPER_BOUND"
    strSQL = strSQL & "      ,ELEM_TYPE_NAME"
    strSQL = strSQL & "      ,LENGTH"
    strSQL = strSQL & "  FROM ALL_COLL_TYPES"
    strSQL = strSQL & " WHERE TYPE_NAME = '" & strTypeName & "'"
    
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oraDyna.EOF = False Then
        With typCollType
            'ｵｰﾅ
            .tOwner = IIf(IsNull(oraDyna![Owner]), "", RTrim(oraDyna![Owner]))
            'Type型名称
            .tTypeName = IIf(IsNull(oraDyna![TYPE_NAME]), "", RTrim(oraDyna![TYPE_NAME]))
            'ｺﾚｸｼｮﾝﾀｲﾌﾟ
            .tCollType = IIf(IsNull(oraDyna![COLL_TYPE]), "", RTrim(oraDyna![COLL_TYPE]))
            '配列数
            .tUpperBound = IIf(IsNull(oraDyna![UPPER_BOUND]), "", RTrim(oraDyna![UPPER_BOUND]))
            '型
            .tElemTypeName = IIf(IsNull(oraDyna![ELEM_TYPE_NAME]), "", RTrim(oraDyna![ELEM_TYPE_NAME]))
            '型のｻｲｽﾞ
            .tLength = IIf(IsNull(oraDyna![length]), "", RTrim(oraDyna![length]))
        End With
    End If
    Set oraDyna = Nothing
    
    GF_GetAllCollType = True
    
    Exit Function
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_GetAllCollType", strSQL)
    
End Function

''2012/06/11 J.Yamaoka Add
Public Function GF_ReplaceSQLLikeEscape(ByVal strCondition As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　: SQL文内のLikeエスケープ,リテラル置換
' 機能　　　:
' 引数　　　: strCondition As String     ''置換対象文字列
' 戻り値　　: 置換した文字列
' 機能説明　:
'------------------------------------------------------------------------------
    strCondition = Replace(strCondition, mstrLikeEscape, String(2, mstrLikeEscape))
    strCondition = Replace(strCondition, "%", mstrLikeEscape & "%", , , vbBinaryCompare)
' 2017/02/08 ▼ D.Ikeda K545 CSプロセス改善  DEL
'    strCondition = Replace(strCondition, "％", mstrLikeEscape & "％", , , vbBinaryCompare)
' 2017/02/08 ▲ D.Ikeda K545 CSプロセス改善  DEL
    strCondition = Replace(strCondition, "_", mstrLikeEscape & "_", , , vbBinaryCompare)
' 2017/02/08 ▼ D.Ikeda K545 CSプロセス改善  DEL
'    strCondition = Replace(strCondition, "＿", mstrLikeEscape & "＿", , , vbBinaryCompare)
' 2017/02/08 ▲ D.Ikeda K545 CSプロセス改善  DEL
    GF_ReplaceSQLLikeEscape = strCondition
End Function
''2012/06/11 J.Yamaoka Add
Public Property Get SQLLikeEscape() As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:Likeエスケープ文字
' 機能　　　:
' 引数　　　:なし
' 戻り値　　:Likeエスケープ文字
' 機能説明　:
'------------------------------------------------------------------------------
    SQLLikeEscape = mstrLikeEscape
End Property

'<LFDB更新 SESSIONID修正> ADD Start NIC 王  2015/08/21
Public Function GF_GetSessionID_Func(plngSessionID As Long, pstrErrMsg As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名　　:シーケンス取得関数呼び出し
' 機能　　　:
' 引数　　　:plngSessionID As Long  セッションID
' 　　　　　:pstrErrMsg As String       エラーメッセージ
' 機能説明　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL              As String
    Dim lclsOraClass        As New clsOraClass  ''Oracle関連用クラス
    Dim strErrMsg           As String
    
    GF_GetSessionID_Func = False
        
    'ストアド用オブジェクト宣言
    Set lclsOraClass = New clsOraClass
    Set lclsOraClass.OraDataBase_Strcall = gOraDataBase
    'サーバーエラーのリセット
    Call lclsOraClass.ErrReset_Strcall
    'バインド変数追加
    lclsOraClass.Add_Binds ORAPARM_OUTPUT, ORATYPE_NUMBER, "SESSION", plngSessionID       'セッションID
    lclsOraClass.Add_Binds ORAPARM_OUTPUT, ORATYPE_VARCHAR2, "OUTMSG", pstrErrMsg      '結果エラーメッセージ
    If (lclsOraClass.ErrCode_Strcall <> 0 Or lclsOraClass.ErrText_Strcall <> "") Then
        Err.Raise Number:=lclsOraClass.ErrCode_Strcall, Description:=lclsOraClass.ErrText_Strcall
    End If
    
    strSQL = ""
    strSQL = strSQL & "BEGIN "
    strSQL = strSQL & ":sql_code:=LFSYS.GLOBAL_SESSIONID_GET "
    strSQL = strSQL & " (:SESSION,   "
    strSQL = strSQL & "  :OUTMSG   "
    strSQL = strSQL & " );  "
    strSQL = strSQL & " END;"

    'サーバーエラーのリセット
    Call lclsOraClass.ErrReset_Strcall
    'SQL実行
    Call lclsOraClass.ExecSql_Strcall(strSQL)

    'チェック結果メッセージ取得
    strErrMsg = GF_VarToStr(gOraParam!OUTMSG)

    If (gOraParam!sql_code = -1) Then
        'システム異常
        GF_GetSessionID_Func = False
        '書き込み失敗
        'ログ出力
        Call GF_GetMsg_Addition("WTK009", , False, True)
        Exit Function
    Else
        plngSessionID = GF_VarToStr(gOraParam!Session)
        GF_GetSessionID_Func = True
    End If

    'パラメーターの全解放
    lclsOraClass.RemoveAll

    Set lclsOraClass = Nothing
    
    Exit Function

ErrHandler:
    'ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_GetSessionID_Func", strSQL)
    
End Function
'<LFDB更新 SESSIONID修正> ADD end NIC 王  2015/08/21


