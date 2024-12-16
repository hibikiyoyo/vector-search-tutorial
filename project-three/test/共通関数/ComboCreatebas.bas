Attribute VB_Name = "basComboCreate"
' @(h)  ComboCreate.BAS           ver 1.0    ( 2000/12/04 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   プロジェクト名  :   TLFﾌﾟﾛｼﾞｪｸﾄ
'   モジュール名    :   basComboCreate
'   ファイル名      :   ComboCreate.BAS
'   Ｖｅｒｓｉｏｎ  :   1.00
'   機能説明        :   各種コンボボックスの作成関数
'   作成者          :   T.Fukutani
'   作成日          :   2000/12/04
'   修正履歴        :   2001/04/29 N.Kigaku 販売店ｺﾝﾎﾞﾎﾞｯｸｽ･ﾘｽﾄﾎﾞｯｸｽ作成関数に引数1つ追加
'　 　　　　　　　　：  2001/11/26 N.Kigaku GF_CreateEigyoCombo,GF_MatchComboを追加
'　 　　　　　　　　：  2001/12/06 N.Kigaku GF_Com_CtlAdditem修正
'　 　　　　　　　　：  2001/12/12 N.Kigaku GF_Com_CtlAdditem2,GF_CreateCifCombo2 追加
'　 　　　　　　　　：  2002/01/09 N.Kigaku GF_CreateCifCombo2に引数2つ追加,GF_CreateMastCombo追加
'　 　　　　　　　　：  2002/01/24 N.Kigaku GF_CreateSyasyuCombo追加
'　 　　　　　　　　：  2002/10/10 GF_CreateBunruiComboとGF_CreateBunruiCombo2の分類1,2,3ｺﾝﾎﾞ作成SQL文にｸﾞﾙｰﾌﾟ化を追加
'　 　　　　　　　　：  2005/09/19 N.Kigaku GF_CreateDistCombo追加
'　 　　　　　　　　：  2005/10/25 N.Kigaku GF_CreateDistCombo修正,GF_CreateGroupCombo追加
'　 　　　　　　　　：  2005/11/02 N.Kigaku GF_CreateGroupCombo修正
'　 　　　　　　　　：  2005/11/04 N.Kigaku GF_CreateDistCombo修正
'　 　　　　　　　　：  2005/11/16 N.Kigaku GF_CreateDistCombo修正
'                   ：  2006/02/10 N.Kigaku GF_CreateGroupCombo 大文字で検索するように修正
'                   ：  2006/12/05 N.Kigaku ｵﾗｸﾙ8.1.7 Nocache対応 検索時、ReadOnlyからNocacheに変更
'                   ：  2016/12/15 M.Tanaka K545 CSプロセス改善 GF_CreateGroupList追加
'                   ：  2018/05/07 M.Kawamura K545 CSプロセス改善
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 環境宣言
'------------------------------------------------------------------------------
Option Explicit

' 2016/12/15 ▼ M.Tanaka K545 CSプロセス改善  追加
'------------------------------------------------------------------------------
' パブリック定数宣言
'------------------------------------------------------------------------------
'機種担当グループリストボックス作成用列挙型の宣言
Public Enum CGL_KaitoFlg     '回答部署フラグ
    CGL_InquiryKaitoFlg = 1  '引合回答部署フラグ
    CGL_DeliveryKaitoFlg = 2 '引合納期回答部署フラグ
    CGL_EDFKaitoFlg = 3      '仕決回答部署フラグ
' 2018/05/07 ▼ M.Kawamura K545 CSプロセス改善
    CGL_DeliEDFKaitoFlg = 4  '引合納期・仕決回答部署フラグ
' 2018/05/07 ▲ M.Kawamura K545 CSプロセス改善
End Enum
Public Enum CGL_HonkiAttKbn  '本機ATT区分
    CGL_HonkiAttAll = 0      '本機ATT区分の条件をつけない
    CGL_Att = 1              'ATT
    CGL_Honki = 2            '本機
    CGL_Sonota = 3           'その他
End Enum
Public Enum CGL_IdArrayKbn   'ID用配列内容区分
    CGL_Id = 1               'グループIDのみ
    CGL_IdAndModelKaito = 2  'グループID & ',' & 機種マスタ回答部署区分
End Enum
Public Enum CGL_EigyoDispKbn '営業表示区分
    CGL_EigyoAll = 0         '国内、海外の条件をつけない
    CGL_Kokunai = 1          '国内
    CGL_Kaigai = 2           '海外
End Enum
' 2016/12/15 ▲ M.Tanaka K545 CSプロセス改善  追加

Public Function GF_Com_CtlAdditem(cntControl As Control, strCombo As String, Optional intIndex As Integer = 0, _
            Optional intOption As Integer = 0, Optional bolCDFlg As Boolean = True) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   コンボボックス・リストボックスの作成
' 機能      :   ｺﾝﾎﾞﾎﾞｯｸｽﾏｽﾀより引合ｺﾝﾎﾞﾎﾞｯｸｽを作成する
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               strCombo   As String     コンボ or リストボックスの名前
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
'               bolCDFlg   As Boolean    コードの表示(省略時)/非表示    (True:表示、False:非表示)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String    'コントロール名
    Dim strCode      As String
    Dim strName      As String
    
    GF_Com_CtlAdditem = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    strComboName = Trim(strCombo)
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT CDVAL,"
    strSQL = strSQL & "       CDNAME"
    strSQL = strSQL & "   FROM THJCMBXMR"
    strSQL = strSQL & "   WHERE CMBNAME = '" & strComboName & "'"
    strSQL = strSQL & "   ORDER BY SEQNO"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![CDVAL])
            strName = GF_VarToStr(oraDyna![CDNAME])
            
            'コード表示/非表示
            If bolCDFlg = True Then
                '表示
                cntControl.AddItem strCode & "：" & strName
            Else
                '非表示
                cntControl.AddItem strName
            End If
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_Com_CtlAdditem = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_Com_CtlAdditem", strSQL)

End Function

Public Function GF_Com_CtlAdditem2(cntControl As Control _
                                , strCombo As String _
                                , Optional intIndex As Integer = 0 _
                                , Optional intOption As Integer = 0 _
                                , Optional intHyojiKbn As Integer = 1 _
                                , Optional intSpace As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   コンボボックス・リストボックスの作成
' 機能      :   ｺﾝﾎﾞﾎﾞｯｸｽﾏｽﾀより引合ｺﾝﾎﾞﾎﾞｯｸｽを作成する
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               strCombo   As String     コンボ or リストボックスの名前
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
'               intHyojiKbn as Integer   表示内容区分 (省略時 ｺｰﾄﾞ:名称)
'                  1 = ｺｰﾄﾞ:名称
'                  2 = 名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
'                  3 = 名称
'               intSpace  As Integer     名称とｺｰﾄﾞとの間隔(省略時0)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String    'コントロール名
    Dim strCode      As String
    Dim strName      As String
    
    GF_Com_CtlAdditem2 = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    strComboName = Trim(strCombo)
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT CDVAL,"
    strSQL = strSQL & "       CDNAME"
    strSQL = strSQL & "   FROM THJCMBXMR"
    strSQL = strSQL & "   WHERE CMBNAME = '" & strComboName & "'"
    strSQL = strSQL & "   ORDER BY SEQNO"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![CDVAL])
            strName = GF_VarToStr(oraDyna![CDNAME])
            
            If intHyojiKbn = 1 Then
            'ｺｰﾄﾞ:名称
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "：" & strName)
                
            ElseIf intHyojiKbn = 2 Then
            '名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
                cntControl.AddItem strName & Space(intSpace) & "：" & strCode
            
            ElseIf intHyojiKbn = 3 Then
            '名称
                cntControl.AddItem strName
                
            End If
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_Com_CtlAdditem2 = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_Com_CtlAdditem2", strSQL)

End Function

Public Function GF_Com_CboGetText(cntControl As Control) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   コンボボックス・リストボックスのテキスト切り出し
' 機能      :   コンボボックスに設定されているテキストの文字部（：より右）を取り出す
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
' 戻り値    :   String      抜き出した文字部
' 備考      :   コンボ未選択の場合や、空白テキストを選択の場合は Null を返却
'------------------------------------------------------------------------------
    GF_Com_CboGetText = ""
    
    If cntControl.ListIndex >= 0 Then
        If (InStr(1, cntControl.Text, "：") - 1) > 0 Then
            GF_Com_CboGetText = Right(cntControl.Text, Len(cntControl.Text) - InStr(1, cntControl.Text, "："))
        End If
    End If
    
End Function

Public Function GF_Com_CboGetCode(cntControl As Control) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   コンボボックス・リストボックスのテキスト切り出し
' 機能      :   コンボボックスに設定されているテキストの文字部（：より左）を取り出す
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
' 戻り値    :   String      抜き出した文字部
' 備考      :   コンボ未選択の場合や、空白テキストを選択の場合は Null を返却
'------------------------------------------------------------------------------
    GF_Com_CboGetCode = ""
    
    If cntControl.ListIndex >= 0 Then
        If (InStr(1, cntControl.Text, "：") - 1) > 0 Then
            GF_Com_CboGetCode = Left(cntControl.Text, InStr(1, cntControl.Text, "：") - 1)
        End If
    End If
    
End Function

Public Function GF_CreateTantoCombo(cntControl As Control, Optional intPostFlg As Integer, _
        Optional intIndex As Integer = 0, Optional intOption As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   担当者コンボボックス・リストボックスの作成
' 機能      :   社員ﾏｽﾀよりｺﾝﾎﾞﾎﾞｯｸｽを作成する
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               intPostFlg As Integer    部署区分(1：CS設計、2:開発四室、3:営業、4:生管)
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String    'コントロール名
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateTantoCombo = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT SYAINCD,"
    strSQL = strSQL & "       NAME"
    strSQL = strSQL & "   FROM THJUSRMR"
    If intPostFlg = 1 Then
        strSQL = strSQL & "   WHERE CS = '1'"
    ElseIf intPostFlg = 2 Then
        strSQL = strSQL & "   WHERE S4 = '1'"
    ElseIf intPostFlg = 3 Then
        strSQL = strSQL & "   WHERE EIGYO = '1'"
    ElseIf intPostFlg = 4 Then
        strSQL = strSQL & "   WHERE SEIKAN = '1'"
    End If
    strSQL = strSQL & "   ORDER BY SYAINCD"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![SYAINCD])
            strName = GF_VarToStr(oraDyna![Name])
            cntControl.AddItem strCode & "：" & strName
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateTantoCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateTantoCombo", strSQL)

End Function

Public Function GF_CreateCifCombo(cntControl As Control, Optional intIndex As Integer = 0, _
                                  Optional intOption As Integer = 0, Optional bolItemFlg As Boolean = False) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   販売店コンボボックス・リストボックスの作成
' 機能      :   販売店ﾏｽﾀよりｺﾝﾎﾞﾎﾞｯｸｽを作成する
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
'               bolItemFlg As Boolea     ItemDataに販売店ｺｰﾄﾞを設定するか否か(省略時否)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String    'コントロール名
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateCifCombo = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT CIFNO,"
    strSQL = strSQL & "       CIFNAME"
    strSQL = strSQL & "   FROM THJCIF"
    strSQL = strSQL & "   GROUP BY CIFNO,CIFNAME"
    strSQL = strSQL & "   ORDER BY CIFNO"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![CIFNO])
            strName = GF_VarToStr(oraDyna![CIFNAME])
            cntControl.AddItem strCode & "：" & strName
            If bolItemFlg = True Then
                cntControl.ItemData(cntControl.NewIndex) = strCode
            End If
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateCifCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateCifCombo", strSQL)

End Function

Public Function GF_CreateEigyoCombo(cntControl As Control _
                                  , Optional strCifNO As String = "" _
                                  , Optional intIndex As Integer = 0 _
                                  , Optional intOption As Integer = 0 _
                                  , Optional bolItemFlg As Boolean = False _
                                  ) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   営業所コンボボックス・リストボックスの作成
' 機能      :   販売店ﾏｽﾀよりｺﾝﾎﾞﾎﾞｯｸｽを作成する
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               strCifNO   As String     販売店コード(省略時空白)
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
'               bolItemFlg As Boolea     ItemDataに営業所ｺｰﾄﾞを設定するか否か(省略時否)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String    'コントロール名
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateEigyoCombo = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT EIGYONO,"
    strSQL = strSQL & "       EIGYONAME"
    strSQL = strSQL & "   FROM THJCIF"
    If Len(Trim(strCifNO)) > 0 Then
        strSQL = strSQL & "   WHERE CIFNO='" & strCifNO & "'"
    End If
    strSQL = strSQL & "   ORDER BY CIFNO"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![EIGYONO])
            strName = GF_VarToStr(oraDyna![EIGYONAME])
            cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "：" & strName)
            If bolItemFlg = True Then
                cntControl.ItemData(cntControl.NewIndex) = strCode
            End If
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateEigyoCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateEigyoCombo", strSQL)

End Function

Public Function GF_CreateEigyoCombo2(cntControl As Control _
                                  , Optional strCifNO As String = "" _
                                  , Optional intIndex As Integer = 0 _
                                  , Optional intOption As Integer = 0 _
                                  , Optional bolItemFlg As Boolean = False _
                                  , Optional intHyojiKbn As Integer = 1 _
                                  , Optional intSpace As Integer = 0 _
                                  , Optional blnDispNameFlg As Boolean = True _
                                  ) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   営業所コンボボックス・リストボックスの作成
' 機能      :   販売店ﾏｽﾀよりｺﾝﾎﾞﾎﾞｯｸｽを作成する
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               strCifNO   As String     販売店コード(省略時空白)
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
'               bolItemFlg As Boolea     ItemDataに営業所ｺｰﾄﾞを設定するか否か(省略時否)
'               intHyojiKbn as Integer     表示内容区分 (省略時 ｺｰﾄﾞ:名称)
'                  1 = ｺｰﾄﾞ:名称
'                  2 = 名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
'                  3 = 名称
'               intSpace  As Integer       名称とｺｰﾄﾞとの間隔(省略時0)
'               blnDispNameFlg As Boolean  名称が無い時に追加するか否か(省略時追加)   False:追加しない、True:追加
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String    'コントロール名
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateEigyoCombo2 = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT EIGYONO,"
    strSQL = strSQL & "       EIGYONAME"
    strSQL = strSQL & "   FROM THJCIF"
    If Len(Trim(strCifNO)) > 0 Then
        strSQL = strSQL & "   WHERE CIFNO='" & strCifNO & "'"
    End If
    strSQL = strSQL & "   ORDER BY CIFNO"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![EIGYONO])
            strName = GF_VarToStr(oraDyna![EIGYONAME])
            If intHyojiKbn = 1 Then
            'ｺｰﾄﾞ:名称
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "：" & strName)
                End If
                
            ElseIf intHyojiKbn = 2 Then
            '名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "：" & strCode)
                End If
                
            ElseIf intHyojiKbn = 3 Then
            '名称
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem strName
                End If
                
            End If

            If bolItemFlg = True Then
                cntControl.ItemData(cntControl.NewIndex) = strCode
            End If

            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateEigyoCombo2 = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateEigyoCombo2", strSQL)

End Function

Public Function GF_CreateCifCombo2(cntControl As Control _
                                , intCifKbn As Integer _
                                , Optional strCifNO As String = "" _
                                , Optional intIndex As Integer = 0 _
                                , Optional intOption As Integer = 0 _
                                , Optional intHyojiKbn As Integer = 1 _
                                , Optional intSpace As Integer = 0 _
                                , Optional blnDispNameFlg As Boolean = True _
                                , Optional intShiyuKbn As Integer = 0 _
                                , Optional intDispCS As Integer = 2) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   販売店／営業所コンボボックス・リストボックスの作成２
' 機能      :   販売店ﾏｽﾀよりｺﾝﾎﾞﾎﾞｯｸｽを作成する
' 引数      :   cntControl As Control      対象となるコンボコントロール及びリストコントロール
'               intCifKbn As Integer       販売店・営業所作成区分
'                  1 = 販売店
'                  2 = 営業所
'               strCifNO As String         販売店ｺｰﾄﾞ(販売店・営業所作成区分が営業所の時のみ)
'               intIndex   As Integer      デフォルト表示インデックス(省略時0)
'               intOption  As Integer      選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
'               intHyojiKbn as Integer     表示内容区分 (省略時 ｺｰﾄﾞ:名称)
'                  1 = ｺｰﾄﾞ:名称
'                  2 = 名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
'                  3 = 名称
'               intSpace  As Integer       名称とｺｰﾄﾞとの間隔(省略時0)
'               blnDispNameFlg As Boolean  名称が無い時に追加するか否か(省略時追加)   False:追加しない、True:追加
'               intShiyuKbn As Integer     市輸区分(省略時 国内)
'                  0 = 国内と海外
'                  1 = 国内のみ
'                  2 = 海外のみ
'               intDispCS As Integer       表示区分(省略時:0)  0:機台、1:環境機器、2:全て
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim strSQL_Srh   As String      '検索条件SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String    'コントロール名
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateCifCombo2 = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL文
    strSQL = ""
    strSQL_Srh = ""
    
    ''市輸区分条件作成
    Select Case intShiyuKbn
        Case 1
            strSQL_Srh = " SHIYUKBN = '1'"
        Case 2
            strSQL_Srh = " SHIYUKBN = '2'"
        Case Else
            strSQL_Srh = ""
    End Select
    
    ''表示区分条件作成
    Select Case intDispCS
        Case 0, 1
        
            If Len(strSQL_Srh) > 0 Then
                strSQL_Srh = strSQL_Srh & " AND "
            End If
            If intDispCS = 0 Then
                '機台
                strSQL_Srh = " CORDER_DISP_CS = '1'"
            ElseIf intDispCS = 1 Then
                '環境機器
                strSQL_Srh = " CENV_CARRY_CS = '1'"
            End If
        
        Case Else
    End Select
    
    If intCifKbn = 1 Then
        '販売店ｺﾝﾎﾞﾎﾞｯｸｽ･ﾘｽﾄﾎﾞｯｸｽの作成
        strSQL = strSQL & "SELECT CIFNO C_NO,"
        strSQL = strSQL & "       CIFNAME C_NAME"
        strSQL = strSQL & "  FROM THJCIF"
         strSQL = strSQL & IIf(Len(strSQL_Srh) = 0, "", " WHERE" & strSQL_Srh)
        strSQL = strSQL & " GROUP BY CIFNO,CIFNAME"
        strSQL = strSQL & " ORDER BY CIFNO"
    ElseIf intCifKbn = 2 Then
        '営業所ｺﾝﾎﾞﾎﾞｯｸｽ･ﾘｽﾄﾎﾞｯｸｽの作成
        strSQL = strSQL & "SELECT EIGYONO C_NO,"
        strSQL = strSQL & "       EIGYONAME C_NAME"
        strSQL = strSQL & "  FROM THJCIF"
        If Len(Trim(strCifNO)) > 0 Then
            strSQL = strSQL & "   WHERE CIFNO='" & strCifNO & "'"
            strSQL = strSQL & IIf(Len(strSQL_Srh) = 0, "", " AND" & strSQL_Srh)
        Else
            strSQL = strSQL & IIf(Len(strSQL_Srh) = 0, "", " WHERE" & strSQL_Srh)
        End If
        strSQL = strSQL & " GROUP BY EIGYONO,EIGYONAME"
        strSQL = strSQL & " ORDER BY EIGYONO"
    End If
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![C_NO])
            strName = GF_VarToStr(oraDyna![C_NAME])
            If intHyojiKbn = 1 Then
            'ｺｰﾄﾞ:名称
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "：" & strName)
                End If
                
            ElseIf intHyojiKbn = 2 Then
            '名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "：" & strCode)
                End If
'                cntControl.AddItem strName & Space(intSpace) & "：" & strCode
                
            ElseIf intHyojiKbn = 3 Then
            '名称
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem strName
                End If
                
            End If
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateCifCombo2 = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateCifCombo2", strSQL)

End Function

Public Function GF_CreateBunruiCombo(cntControl As Control, intBunruiFlg As Integer, _
                            Optional strBunrui1 As String, Optional strBunrui2 As String, _
                            Optional intOption As Integer = 1) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名　　:　 分類ｺﾝﾎﾞﾎﾞｯｸｽの作成
' 機能　　　:　 C-OPT分類ﾃｰﾌﾞﾙよりｺﾝﾎﾞﾎﾞｯｸｽを作成する。
' 引数　　　:　 cntControl     As Control    対象となるｺﾝﾎﾞｺﾝﾄﾛｰﾙ
' 　　　　　 　 intBunruiFlg   As Integer    分類区分(1：分類１, 2:分類２, 3:分類３)
' 　　　　　 　 strBunrui1　   AS String     分類ｺｰﾄﾞ1
' 　　　　　 　 strBunrui2　   AS String     分類ｺｰﾄﾞ2
' 　　　　　 　 intOption      As Integer    選択項目の先頭にﾇﾙ項目を追加する場合に１を指定する(省略時あり)
' 戻り値　　:　 True = 成功 / False = 失敗
' 備考　　　:
'------------------------------------------------------------------------------
    Dim strSQL       As String      'SQL文
    Dim oDynaset     As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String      'ｺﾝﾄﾛｰﾙ名
    Dim strCode      As String
    Dim strCode1     As String * 4
    Dim strCode2     As String * 4
    Dim strName      As String
    On Error GoTo ErrHandler
    GF_CreateBunruiCombo = False
    
'   ｺﾝﾄﾛｰﾙ初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
'   SQL文
    strSQL = ""
    Select Case intBunruiFlg
        Case 1
            strSQL = strSQL & " SELECT BUNRUI1     BUNRUI,"
            strSQL = strSQL & "        BUNRUINAME1 BUNRUINAME"
            strSQL = strSQL & "　 FROM THJBUNRUI1"
            strSQL = strSQL & "  GROUP BY BUNRUI1,BUNRUINAME1"
            strSQL = strSQL & "  ORDER BY BUNRUI1"
        Case 2
            If Trim(strBunrui1) = "" Then
                Exit Function
            End If
            strSQL = strSQL & " SELECT BUNRUI2     BUNRUI,"
            strSQL = strSQL & "        BUNRUINAME2 BUNRUINAME"
            strSQL = strSQL & "   FROM THJBUNRUI2  "
            strSQL = strSQL & "  WHERE BUNRUI1     = '" & Trim(strBunrui1) & "'"
            strSQL = strSQL & "  GROUP BY BUNRUI2,BUNRUINAME2"
            strSQL = strSQL & "  ORDER BY BUNRUI2 "
        Case 3
            If Trim(strBunrui1) = "" Or Trim(strBunrui2) = "" Then
                Exit Function
            End If
            strSQL = strSQL & " SELECT BUNRUI3     BUNRUI,"
            strSQL = strSQL & "        BUNRUINAME3 BUNRUINAME"
            strSQL = strSQL & "   FROM THJBUNRUI3  "
            strSQL = strSQL & "  WHERE BUNRUI1     = '" & Trim(strBunrui1) & "'"
            strSQL = strSQL & "    AND BUNRUI2     = '" & Trim(strBunrui2) & "'"
            strSQL = strSQL & "  GROUP BY BUNRUI3,BUNRUINAME3"
            strSQL = strSQL & "  ORDER BY BUNRUI3 "
    End Select
    Set oDynaset = Nothing
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
'   該当ﾚｺｰﾄﾞがない場合
    If oDynaset.EOF = True Then
        Exit Function
    End If
'   ｺﾝﾎﾞ作成
'   ﾇﾙ値設定
    If intOption <> 0 Then
        cntControl.AddItem ""
    End If
    Do Until oDynaset.EOF = True
        strCode = GF_VarToStr(oDynaset![BUNRUI])
        strName = GF_VarToStr(oDynaset![BUNRUINAME])
        cntControl.AddItem Trim(strCode) & "：" & Trim(strName)
        oDynaset.MoveNext
    Loop
    
    GF_CreateBunruiCombo = True
    Exit Function
    
ErrHandler:
'   ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateBunruiCombo", strSQL)

End Function

Public Function GF_CreateBunruiCombo2(cntControl As Control, intBunruiFlg As Integer, _
                            strKatashiki As String, Optional strBunrui1 As String, _
                            Optional strBunrui2 As String, Optional intOption As Integer = 1 _
                            ) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名　　:　 分類ｺﾝﾎﾞﾎﾞｯｸｽの作成
' 機能　　　:　 機種型式でC-OPTﾏｽﾀを検索し、一致する分類名称をC-OPT分類ﾃｰﾌﾞﾙより
'              取得してｺﾝﾎﾞﾎﾞｯｸｽを作成する。
' 引数　　　:　 cntControl     As Control    対象となるｺﾝﾎﾞｺﾝﾄﾛｰﾙ
' 　　　　　 　 intBunruiFlg   As Integer    分類区分(1：分類１, 2:分類２, 3:分類３)]
'              strKatashiki   As String     機種型式
' 　　　　　 　 strBunrui1　   AS String     分類ｺｰﾄﾞ1
' 　　　　　 　 strBunrui2　   AS String     分類ｺｰﾄﾞ2
' 　　　　　 　 intOption      As Integer    選択項目の先頭にﾇﾙ項目を追加する場合に１を指定する(省略時あり)
' 戻り値　　:　 True = 成功 / False = 失敗
' 備考　　　:
'------------------------------------------------------------------------------
    Dim strSQL       As String      'SQL文
    Dim oDynaset     As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String      'ｺﾝﾄﾛｰﾙ名
    Dim strCode      As String
    Dim strCode1     As String * 4
    Dim strCode2     As String * 4
    Dim strName      As String
    On Error GoTo ErrHandler
    GF_CreateBunruiCombo2 = False
    
'   ｺﾝﾄﾛｰﾙ初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
'   SQL文
    strSQL = ""
    Select Case intBunruiFlg
        Case 1
            strSQL = strSQL & "SELECT BUNRUI1     BUNRUI"
            strSQL = strSQL & "      ,BUNRUINAME1 BUNRUINAME"
            strSQL = strSQL & " FROM THJBUNRUI1"
            strSQL = strSQL & " WHERE BUNRUI1 IN"
            strSQL = strSQL & "  ("
            strSQL = strSQL & "   SELECT BUNRUI1 FROM THJCOPTMR"
            strSQL = strSQL & "    WHERE KATASHIKI ='" & RTrim(strKatashiki) & "'"
            strSQL = strSQL & "    GROUP BY BUNRUI1"
            strSQL = strSQL & "  )"
            strSQL = strSQL & " GROUP BY BUNRUI1,BUNRUINAME1"
            strSQL = strSQL & " ORDER BY BUNRUI1"
        Case 2
            If Trim(strBunrui1) = "" Then
                Exit Function
            End If
            strSQL = strSQL & "SELECT BUNRUI2     BUNRUI"
            strSQL = strSQL & "      ,BUNRUINAME2 BUNRUINAME"
            strSQL = strSQL & " FROM THJBUNRUI2"
            strSQL = strSQL & " WHERE BUNRUI2 IN"
            strSQL = strSQL & "  ("
            strSQL = strSQL & "   SELECT BUNRUI2 FROM THJCOPTMR"
            strSQL = strSQL & "    WHERE KATASHIKI ='" & RTrim(strKatashiki) & "'"
            strSQL = strSQL & "      AND BUNRUI1 = '" & RTrim(strBunrui1) & "'"
            strSQL = strSQL & "    GROUP BY BUNRUI2"
            strSQL = strSQL & "  )"
            strSQL = strSQL & " GROUP BY BUNRUI2,BUNRUINAME2"
            strSQL = strSQL & " ORDER BY BUNRUI2"
        Case 3
            If Trim(strBunrui1) = "" Or Trim(strBunrui2) = "" Then
                Exit Function
            End If
            strSQL = strSQL & "SELECT BUNRUI3     BUNRUI"
            strSQL = strSQL & "      ,BUNRUINAME3 BUNRUINAME"
            strSQL = strSQL & " FROM THJBUNRUI3"
            strSQL = strSQL & " WHERE BUNRUI3 IN"
            strSQL = strSQL & "  ("
            strSQL = strSQL & "   SELECT BUNRUI3 FROM THJCOPTMR"
            strSQL = strSQL & "    WHERE KATASHIKI ='" & RTrim(strKatashiki) & "'"
            strSQL = strSQL & "      AND BUNRUI1 = '" & RTrim(strBunrui1) & "'"
            strSQL = strSQL & "      AND BUNRUI2 = '" & RTrim(strBunrui2) & "'"
            strSQL = strSQL & "    GROUP BY BUNRUI3"
            strSQL = strSQL & "  )"
            strSQL = strSQL & " GROUP BY BUNRUI3,BUNRUINAME3"
            strSQL = strSQL & " ORDER BY BUNRUI3"
    End Select
    Set oDynaset = Nothing
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
'   該当ﾚｺｰﾄﾞがない場合
    If oDynaset.EOF = True Then
        Exit Function
    End If
'   ｺﾝﾎﾞ作成
'   ﾇﾙ値設定
    If intOption <> 0 Then
        cntControl.AddItem ""
    End If
    Do Until oDynaset.EOF = True
        strCode = GF_VarToStr(oDynaset![BUNRUI])
        strName = GF_VarToStr(oDynaset![BUNRUINAME])
        cntControl.AddItem Trim(strCode) & "：" & Trim(strName)
        oDynaset.MoveNext
    Loop
    
    GF_CreateBunruiCombo2 = True
    Exit Function
    
ErrHandler:
'   ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateBunruiCombo2", strSQL)

End Function

Public Function GF_CreateBunruiCombo2_2(cntControl As Control, intBunruiFlg As Integer, _
                            strKatashiki As String, Optional strBunrui1 As String, _
                            Optional strBunrui2 As String, Optional intOption As Integer = 1 _
                            ) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名　　:　 分類ｺﾝﾎﾞﾎﾞｯｸｽの作成
' 機能　　　:　 機種型式でC-OPTﾏｽﾀ2を検索し、一致する分類名称をC-OPT分類ﾃｰﾌﾞﾙより
'              取得してｺﾝﾎﾞﾎﾞｯｸｽを作成する。
' 引数　　　:　 cntControl     As Control    対象となるｺﾝﾎﾞｺﾝﾄﾛｰﾙ
' 　　　　　 　 intBunruiFlg   As Integer    分類区分(1：分類１, 2:分類２, 3:分類３)]
'              strKatashiki   As String     機種型式
' 　　　　　 　 strBunrui1　   AS String     分類ｺｰﾄﾞ1
' 　　　　　 　 strBunrui2　   AS String     分類ｺｰﾄﾞ2
' 　　　　　 　 intOption      As Integer    選択項目の先頭にﾇﾙ項目を追加する場合に１を指定する(省略時あり)
' 戻り値　　:　 True = 成功 / False = 失敗
' 備考　　　:
'------------------------------------------------------------------------------
    Dim strSQL       As String      'SQL文
    Dim oDynaset     As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String      'ｺﾝﾄﾛｰﾙ名
    Dim strCode      As String
    Dim strCode1     As String * 4
    Dim strCode2     As String * 4
    Dim strName      As String
    On Error GoTo ErrHandler
    GF_CreateBunruiCombo2_2 = False
    
'   ｺﾝﾄﾛｰﾙ初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
'   SQL文
    strSQL = ""
    Select Case intBunruiFlg
        Case 1
            strSQL = strSQL & "SELECT BUNRUI1     BUNRUI"
            strSQL = strSQL & "      ,BUNRUINAME1 BUNRUINAME"
            strSQL = strSQL & " FROM THJBUNRUI1"
            strSQL = strSQL & " WHERE BUNRUI1 IN"
            strSQL = strSQL & "  ("
            strSQL = strSQL & "   SELECT BUNRUI1 FROM THJCOPTMR2"
            strSQL = strSQL & "    WHERE KATASHIKI ='" & RTrim(strKatashiki) & "'"
            strSQL = strSQL & "    GROUP BY BUNRUI1"
            strSQL = strSQL & "  )"
            strSQL = strSQL & " ORDER BY BUNRUI1"
        Case 2
            If Trim(strBunrui1) = "" Then
                Exit Function
            End If
            strSQL = strSQL & "SELECT BUNRUI2     BUNRUI"
            strSQL = strSQL & "      ,BUNRUINAME2 BUNRUINAME"
            strSQL = strSQL & " FROM THJBUNRUI2"
            strSQL = strSQL & " WHERE BUNRUI2 IN"
            strSQL = strSQL & "  ("
            strSQL = strSQL & "   SELECT BUNRUI2 FROM THJCOPTMR2"
            strSQL = strSQL & "    WHERE KATASHIKI ='" & RTrim(strKatashiki) & "'"
            strSQL = strSQL & "      AND BUNRUI1 = '" & RTrim(strBunrui1) & "'"
            strSQL = strSQL & "    GROUP BY BUNRUI2"
            strSQL = strSQL & "  )"
            strSQL = strSQL & " ORDER BY BUNRUI2"
        Case 3
            If Trim(strBunrui1) = "" Or Trim(strBunrui2) = "" Then
                Exit Function
            End If
            strSQL = strSQL & "SELECT BUNRUI3     BUNRUI"
            strSQL = strSQL & "      ,BUNRUINAME3 BUNRUINAME"
            strSQL = strSQL & " FROM THJBUNRUI3"
            strSQL = strSQL & " WHERE BUNRUI3 IN"
            strSQL = strSQL & "  ("
            strSQL = strSQL & "   SELECT BUNRUI3 FROM THJCOPTMR2"
            strSQL = strSQL & "    WHERE KATASHIKI ='" & RTrim(strKatashiki) & "'"
            strSQL = strSQL & "      AND BUNRUI1 = '" & RTrim(strBunrui1) & "'"
            strSQL = strSQL & "      AND BUNRUI2 = '" & RTrim(strBunrui2) & "'"
            strSQL = strSQL & "    GROUP BY BUNRUI3"
            strSQL = strSQL & "  )"
            strSQL = strSQL & " ORDER BY BUNRUI3"
    End Select
    Set oDynaset = Nothing
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
'   該当ﾚｺｰﾄﾞがない場合
    If oDynaset.EOF = True Then
        Exit Function
    End If
'   ｺﾝﾎﾞ作成
'   ﾇﾙ値設定
    If intOption <> 0 Then
        cntControl.AddItem ""
    End If
    Do Until oDynaset.EOF = True
        strCode = GF_VarToStr(oDynaset![BUNRUI])
        strName = GF_VarToStr(oDynaset![BUNRUINAME])
        cntControl.AddItem Trim(strCode) & "：" & Trim(strName)
        oDynaset.MoveNext
    Loop
    
    GF_CreateBunruiCombo2_2 = True
    Exit Function
    
ErrHandler:
'   ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateBunruiCombo2_2", strSQL)

End Function

Public Function GF_MatchCombo(cntControl As Control, strCheck As String _
                            , Optional blnSpaceCheck As Boolean = False _
                            , Optional intLRCheckCode As Integer = 1 _
                            ) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   コンボボックス・リストボックスの表示内容設定
' 機能      :   コンボボックス・リストボックスで一致するテキストに設定する
' 引数      :   cntControl As Control   対象となるコンボコントロール及びリストコントロール
'               strCheck As String      検索対象データ
'               blnSpaceCheck As Boolean 空白時に同一チェックを行うかどうか
'                                         True = ﾁｪｯｸする、 False = ﾁｪｯｸしない
'               intLRCheckCode As Integer 左右どちらのｺｰﾄﾞを取り出すか
'                                         1 = 左 、2 = 右
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    Dim intloop As Integer
    Dim strCode As String
    
    On Error GoTo ErrHandler
    
    GF_MatchCombo = False
    
    'コントロール初期化
    cntControl.ListIndex = -1
    
    For intloop = 0 To cntControl.ListCount - 1
        If (InStr(1, cntControl.List(intloop), "：") - 1) > 0 Then
            If intLRCheckCode = 1 Then
                '：の左側のコードを取り出す
                strCode = Left(cntControl.List(intloop), InStr(1, cntControl.List(intloop), "：", vbTextCompare) - 1)
            Else
                '：の右側のコードを取り出す
                strCode = Mid(cntControl.List(intloop), InStrRev(cntControl.List(intloop), "：", -1, vbTextCompare) + 1)
            End If
            If strCode = strCheck Then
                cntControl.ListIndex = intloop
                Exit For
            End If
        Else
            If blnSpaceCheck = True Then
                If Trim(cntControl.List(intloop)) = strCheck Then
                    cntControl.ListIndex = intloop
                End If
            End If
        End If
    Next intloop
    
    GF_MatchCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_MatchCombo")

End Function


Public Function GF_SetCifCombo(cntControl As Control, strCheck As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   販売店コンボボックス・リストボックスの表示内容設定
' 機能      :   販売店コンボボックス・リストボックスの表示内容を設定する
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               strCheck As String      検索対象データ
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :   ItemDataをﾏｯﾁﾝｸﾞ対象とするため、あらかじめItemDataにﾃﾞｰﾀを入れておく
'               数値型のみ(文字列不可)
'------------------------------------------------------------------------------
    Dim intloop As Integer
    
    On Error GoTo ErrHandler
    
    GF_SetCifCombo = False
    
    'コントロール初期化
    cntControl.ListIndex = -1
    
    For intloop = 0 To cntControl.ListCount - 1
        If cntControl.ItemData(intloop) = strCheck Then
            cntControl.ListIndex = intloop
            Exit For
        End If
    Next intloop
    
    GF_SetCifCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_SetCifCombo")

End Function

Public Function GF_CreateGrpCombo(cntControl As Control, _
        Optional intIndex As Integer = 0, Optional intOption As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   GRPコンボボックス・リストボックスの作成
' 機能      :   C-OPTﾏｽﾀよりｺﾝﾎﾞﾎﾞｯｸｽを作成する
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    
    GF_CreateGrpCombo = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT GRP"
    strSQL = strSQL & " FROM THJCOPTMR"
    strSQL = strSQL & " GROUP BY GRP"
    strSQL = strSQL & " ORDER BY GRP"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If (oraDyna.EOF = True) Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If (intOption <> 0) Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            If (GF_VarToStr(oraDyna![GRP]) <> "") Then
                cntControl.AddItem GF_VarToStr(oraDyna![GRP])
            End If
            
            oraDyna.MoveNext
        Loop Until (oraDyna.EOF = True)
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateGrpCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateGrpCombo", strSQL)
    
End Function

Public Function GF_CreateMastCombo(cntControl As Control, _
        Optional intIndex As Integer = 0, Optional intOption As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   Mastコンボボックス・リストボックスの作成
' 機能      :
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    
    GF_CreateMastCombo = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT MASTTYPE"
    strSQL = strSQL & " FROM THJMSTTAIOMR"
    strSQL = strSQL & " GROUP BY MASTTYPE"
    strSQL = strSQL & " ORDER BY MASTTYPE"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If (oraDyna.EOF = True) Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If (intOption <> 0) Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            If (GF_VarToStr(oraDyna![MASTTYPE]) <> "") Then
                cntControl.AddItem GF_VarToStr(oraDyna![MASTTYPE])
            End If
            
            oraDyna.MoveNext
        Loop Until (oraDyna.EOF = True)
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateMastCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateMastCombo", strSQL)
    
End Function

Public Function GF_CreateSyasyuCombo(cntControl As Control, _
        Optional intIndex As Integer = 0, Optional intOption As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   車種コードコンボボックス・リストボックスの作成
' 機能      :
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    
    GF_CreateSyasyuCombo = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & " SELECT SERIESCD "
    strSQL = strSQL & "   FROM THJHNKTYPEMR"
    strSQL = strSQL & "  WHERE SHIYUKBN = ' '"
    strSQL = strSQL & "  GROUP BY SERIESCD"
    strSQL = strSQL & "  ORDER BY SERIESCD "
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If (oraDyna.EOF = True) Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If (intOption <> 0) Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            If (GF_VarToStr(oraDyna![SERIESCD]) <> "") Then
                cntControl.AddItem GF_VarToStr(oraDyna![SERIESCD])
            End If
            
            oraDyna.MoveNext
        Loop Until (oraDyna.EOF = True)
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateSyasyuCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateSyasyuCombo", strSQL)
    
End Function

Public Function GF_CreatePRFCombo(cntControl As Control _
                                , Optional intIndex As Integer = 0 _
                                , Optional intOption As Integer = 0 _
                                , Optional bolItemFlg As Boolean = False _
                                , Optional intHyojiKbn As Integer = 1 _
                                , Optional intSpace As Integer = 0 _
                                , Optional blnDispNameFlg As Boolean = True _
                                ) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   県マスターコンボボックス・リストボックスの作成
' 機能      :   県マスターよりｺﾝﾎﾞﾎﾞｯｸｽを作成する
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
'               bolItemFlg As Boolea     ItemDataに県ｺｰﾄﾞを設定するか否か(省略時否)
'               intHyojiKbn as Integer     表示内容区分 (省略時 ｺｰﾄﾞ:名称)
'                  1 = ｺｰﾄﾞ:名称
'                  2 = 名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
'                  3 = 名称
'               intSpace  As Integer       名称とｺｰﾄﾞとの間隔(省略時0)
'               blnDispNameFlg As Boolean  名称が無い時に追加するか否か(省略時追加)   False:追加しない、True:追加
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String    'コントロール名
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreatePRFCombo = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT CPREFECTURE_CD,"
    strSQL = strSQL & "       VCPREFECTURE_NAME"
    strSQL = strSQL & "  FROM M23_PREFECTURE"
    strSQL = strSQL & " ORDER BY NLIST_SEQ"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![CPREFECTURE_CD])
            strName = GF_VarToStr(oraDyna![VCPREFECTURE_NAME])
            If intHyojiKbn = 1 Then
            'ｺｰﾄﾞ:名称
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "：" & strName)
                End If
                
            ElseIf intHyojiKbn = 2 Then
            '名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "：" & strCode)
                End If
                
            ElseIf intHyojiKbn = 3 Then
            '名称
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem strName
                End If
                
            End If
            
            If bolItemFlg = True Then
                cntControl.ItemData(cntControl.NewIndex) = strCode
            End If
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreatePRFCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreatePRFCombo", strSQL)

End Function

Public Function GF_CreateDistCombo(cntControl As Control, _
                                   Optional vntGroupCD As Variant = "", _
                                   Optional intIndex As Integer = 0, _
                                   Optional intOption As Integer, _
                                   Optional intHyojiKbn As Integer = 1, _
                                   Optional intSpace As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   仕向先コンボボックス・リストボックスの作成
' 機能      :   取引先ﾏｽﾀよりｺﾝﾎﾞﾎﾞｯｸｽを作成する(ｸﾞﾙｰﾌﾟｺｰﾄﾞ(M30_GROUP)指定可能)
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               vntGroupCD As Variant    グループコード(M30_GROUPに対応)（添字は0から開始）
'                                        1件の時は配列でなくてもOK
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
'               intHyojiKbn as Integer     表示内容区分 (省略時 ｺｰﾄﾞ:名称)
'                  1 = ｺｰﾄﾞ:名称
'                  2 = 名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
'                  3 = 名称
'               intSpace  As Integer       名称とｺｰﾄﾞとの間隔(省略時0)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String      'コントロール名
    Dim strCode      As String
    Dim strName      As String
    Dim intMsgCount As Integer
    Dim i           As Integer
    Dim strWHERE    As String
        
    GF_CreateDistCombo = False
    
    
    '' 配列の数を数える
    If IsArray(vntGroupCD) = True Then
        intMsgCount = UBound(vntGroupCD) + 1
    Else
        intMsgCount = 0
    End If

    strWHERE = ""
    
    '条件文を作成
    If intMsgCount > 0 Then
        '配列時
        For i = 0 To intMsgCount - 1
            If Trim(strWHERE) = "" Then
                strWHERE = strWHERE & "IN ( '" & vntGroupCD(i) & "'"
            Else
                strWHERE = strWHERE & ",'" & vntGroupCD(i) & "'"
            End If
        Next
        strWHERE = strWHERE & ")"
    
    ElseIf (Len(Trim(vntGroupCD)) > 0) Then
        '配列以外
        strWHERE = " = '" & vntGroupCD & "'"
    End If
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    strSQL = ""
    strSQL = strSQL & " SELECT CDIST_CD,VCDIST_NAME "
    strSQL = strSQL & " FROM M05_DIST "
    If strWHERE <> "" Then
        strSQL = strSQL & " WHERE  CGROUP_CD " & strWHERE
        strSQL = strSQL & " OR    CGROUP2_CD " & strWHERE
        strSQL = strSQL & " OR    CGROUP3_CD " & strWHERE
        strSQL = strSQL & " OR    CGROUP4_CD " & strWHERE
        strSQL = strSQL & " OR    CGROUP5_CD " & strWHERE
    End If
    strSQL = strSQL & " GROUP BY CDIST_CD,VCDIST_NAME"
    strSQL = strSQL & " ORDER BY CDIST_CD"
    
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![CDIST_CD])
            strName = GF_VarToStr(oraDyna![VCDIST_NAME])
            
            If intHyojiKbn = 1 Then
            'ｺｰﾄﾞ:名称
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "：" & strName)
                
            ElseIf intHyojiKbn = 2 Then
            '名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "：" & strCode)
                
            ElseIf intHyojiKbn = 3 Then
            '名称
                cntControl.AddItem strName
                
            End If
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateDistCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateDistCombo", strSQL)

End Function

Public Function GF_CreateGroupCombo(cntControl As Control, _
                                   Optional intIndex As Integer = 0, _
                                   Optional intOption As Integer, _
                                   Optional intHyojiKbn As Integer = 1, _
                                   Optional intSpace As Integer = 0, _
                                   Optional strUserID As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   グループコードコンボボックス・リストボックスの作成
' 機能      :　　ﾕｰｻﾞIDに対応するｸﾞﾙｰﾌﾟｺｰﾄﾞをｸﾞﾙｰﾌﾟｽﾀ(M30_GROUP)よりｺﾝﾎﾞﾎﾞｯｸｽを作成する
'　　　　　　　　ﾕｰｻﾞIDの権限がadmin("1")の場合、全ｸﾞﾙｰﾌﾟ対象。
' 引数      :   cntControl As Control    対象となるコンボコントロール及びリストコントロール
'               intIndex   As Integer    デフォルト表示インデックス(省略時0)
'               intOption  As Integer    選択項目の先頭にヌル項目を追加する場合に１を指定する(省略時なし)
'               intHyojiKbn as Integer     表示内容区分 (省略時 ｺｰﾄﾞ:名称)
'                  1 = ｺｰﾄﾞ:名称
'                  2 = 名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
'                  3 = 名称
'               intSpace  As Integer       名称とｺｰﾄﾞとの間隔(省略時0)
'               strUserID As String        ﾕｰｻﾞID
'
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim oraDyna      As OraDynaset  'ﾀﾞｲﾅｾｯﾄ
    Dim strComboName As String      'コントロール名
    Dim strCode      As String
    Dim strName      As String
    Dim strWHERE     As String
    Const strAdminFlg As String = "1"
    
    GF_CreateGroupCombo = False
    
    If Trim(strUserID) <> "" Then
    
    ''ﾕｰｻﾞｰﾏｽﾀｰを検索し、管理ﾌﾗｸﾞを取得
        strSQL = ""
        strSQL = strSQL & " SELECT CADMIN_FLG"
        strSQL = strSQL & " FROM M29_USER "
        strSQL = strSQL & " WHERE UPPER(CUSER_ID) = UPPER('" & Trim(strUserID) & "')"
        
        Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
       
        If oraDyna.EOF = True Then
            ''該当ﾃﾞｰﾀなし
            Exit Function
        Else
            If GF_VarToStr(oraDyna![CADMIN_FLG]) = strAdminFlg Then
                ''管理者の場合、全ｸﾞﾙｰﾌﾟ対象
                strWHERE = ""
            Else
                ''管理者以外、ﾕｰｻﾞｰｸﾞﾙｰﾌﾟﾏｽﾀ登録ｸﾞﾙｰﾌﾟが対象
                strWHERE = ""
                strWHERE = strWHERE & "SELECT CGROUP_CD FROM M31_GROUP_USER"
                strWHERE = strWHERE & " WHERE UPPER(CUSER_ID) = UPPER('" & Trim(strUserID) & "')"
            End If
        
        End If
    
    End If
    
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    
    strSQL = ""
    strSQL = strSQL & " SELECT CGROUP_CD,VCGROUP_NAME "
    strSQL = strSQL & " FROM M30_GROUP "
    If Trim(strWHERE) <> "" Then
        strSQL = strSQL & " WHERE CGROUP_CD IN ( "
        strSQL = strSQL & strWHERE
        strSQL = strSQL & " ) "
    End If
    strSQL = strSQL & " ORDER BY CGROUP_CD"
    
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = True Then
        '''該当データなし
        'ﾒｯｾｰｼﾞ表示
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        'ヌル項目設定あり？
        If intOption <> 0 Then
            'ヌル項目追加
            cntControl.AddItem ""
        End If
        
        '項目設定
        Do
            strCode = GF_VarToStr(oraDyna![CGROUP_CD])
            strName = GF_VarToStr(oraDyna![VCGROUP_NAME])
        
            If intHyojiKbn = 1 Then
            'ｺｰﾄﾞ:名称
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "：" & strName)
                
            ElseIf intHyojiKbn = 2 Then
            '名称 ｽﾍﾟｰｽ :ｺｰﾄﾞ
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "：" & strCode)
                
            ElseIf intHyojiKbn = 3 Then
            '名称
                cntControl.AddItem strName
                
            End If
            
            oraDyna.MoveNext
        
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateGroupCombo = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateGroupCombo", strSQL)

End Function

' 2016/12/15 ▼ M.Tanaka K545 CSプロセス改善  追加
Public Function GF_CreateGroupList(ByRef cntControl As Control, ByRef strGroupIdArray() As String, _
        ByVal intKaitoFlg As CGL_KaitoFlg, ByVal intHonkiAttKbn As CGL_HonkiAttKbn, _
        ByVal intIdArrayKbn As CGL_IdArrayKbn, ByVal intEigyoDispKbn As CGL_EigyoDispKbn) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   機種担当グループリストボックスの作成
' 機能      :
' 引数      :   cntControl        As Control   対象となるリストコントロール
'               srtGroupIDArray() As String    グループID用配列  (グループID、機種マスタ回答部署区分格納用配列)
'               strKaitoFlg       As Integer   回答部署フラグ種類(1:引合回答部署フラグ、2:引合納期回答部署フラグ、3:仕決回答部署フラグ、4:引合納期・仕決回答部署フラグ)
'               strHonkiAttKbn    As Integer   本機ATT区分       (0:本機ATT区分の条件をつけない、1:ATT、2:本機、3:その他)
'               intIdArrayKbn     As Integer   ID用配列内容区分  (1:グループIDのみ、2:グループID & ',' & 機種マスタ回答部署区分)
'               intEigyoDispKbn   As Integer   営業表示区分      (0:国内、海外の条件をつけない、1:国内、2:海外)
' 戻り値    :   True = 成功 / False = 失敗
' 備考      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL文
    Dim strSQL_WHERE As String      'SQL文WHERE句
    Dim oraDynaCount As OraDynaset  'ﾀﾞｲﾅｾｯﾄ(件数)
    Dim oraDynaData  As OraDynaset  'ﾀﾞｲﾅｾｯﾄ(データ)
    Dim strKaigaiEigyoId  As String '海外営業グループID
    Dim strKokunaiEigyoId As String '国内営業グループID
    Dim intCnt       As Integer     '件数
    Dim intIndex     As Integer     'インデックス
    
    GF_CreateGroupList = False
    
    'コントロール初期化
    cntControl.Clear
    cntControl.ListIndex = -1
    strKaigaiEigyoId = ""
    strKokunaiEigyoId = ""
    intCnt = 0
    intIndex = 0
    ReDim strGroupIdArray(0)
  
    'INIファイルマスタより国営、海営のグループIDを取得
    '営業表示区分
    Select Case intEigyoDispKbn
        Case CGL_EigyoAll '国内、海外の条件をつけない場合
            '何もしない
        Case CGL_Kokunai  '国内の場合
            '海営のグループIDを取得
            strKaigaiEigyoId = LF_GetIniTable("KAITOEG_KAIGAI_GROUPID", 1)
            '取得できなかった場合
            If Len(strKaigaiEigyoId) = 0 Then
                'エラーメッセージ表示
                Call GF_GetMsg_Addition("WTK785", "海営グループID", True)
                Exit Function
            End If
        Case CGL_Kaigai   '海外の場合
            '国営のグループIDを取得
            strKokunaiEigyoId = LF_GetIniTable("KAITOEG_KOKUNAI_GROUPID", 1)
            '取得できなかった場合
            If Len(strKokunaiEigyoId) = 0 Then
                'エラーメッセージ表示
                Call GF_GetMsg_Addition("WTK785", "国営グループID", True)
                Exit Function
            End If
    End Select
    
    '件数を取得する
    'SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT COUNT(NGROUP_ID) Cnt"
    strSQL = strSQL & " FROM TCS_GROUP"
    
    'SQL文WHERE句
    strSQL_WHERE = ""
    strSQL_WHERE = strSQL_WHERE & " WHERE 1 = 1"
    '回答部署フラグ
    Select Case intKaitoFlg
        Case CGL_InquiryKaitoFlg  '引合回答部署フラグの場合
            strSQL_WHERE = strSQL_WHERE & " AND CINQUIRY_KAITO_FLG = '1'"
        Case CGL_DeliveryKaitoFlg '引合納期回答部署フラグの場合
            strSQL_WHERE = strSQL_WHERE & " AND CDELIVERY_KAITO_FLG = '1'"
        Case CGL_EDFKaitoFlg      '仕決回答部署フラグの場合
            strSQL_WHERE = strSQL_WHERE & " AND CEDF_KAITO_FLG = '1'"
' 2018/05/07 ▼ M.Kawamura K545 CSプロセス改善
        Case CGL_DeliEDFKaitoFlg  '引合納期・仕決回答部署フラグの場合
            strSQL_WHERE = strSQL_WHERE & " AND ( CDELIVERY_KAITO_FLG = '1'"
            strSQL_WHERE = strSQL_WHERE & "  OR   CEDF_KAITO_FLG = '1' )"
' 2018/05/07 ▲ M.Kawamura K545 CSプロセス改善
    End Select
    '本機ATT区分
    Select Case intHonkiAttKbn
        Case CGL_HonkiAttAll    '本機ATT区分の条件をつけない
            '条件なし
        Case CGL_Att            'ATTの場合
            strSQL_WHERE = strSQL_WHERE & " AND HONKIATTKBN = '1'"
        Case CGL_Honki          '本機の場合
            strSQL_WHERE = strSQL_WHERE & " AND HONKIATTKBN = '2'"
        Case CGL_Sonota         'その他の場合
            strSQL_WHERE = strSQL_WHERE & " AND HONKIATTKBN IS NULL"
    End Select
    '営業表示区分
    Select Case intEigyoDispKbn
        Case CGL_EigyoAll '国内、海外の条件をつけない
            '条件なし
        Case CGL_Kokunai  '国内の場合
            strSQL_WHERE = strSQL_WHERE & " AND NGROUP_ID <> '" & strKaigaiEigyoId & "'"
        Case CGL_Kaigai   '海外の場合
            strSQL_WHERE = strSQL_WHERE & " AND NGROUP_ID <> '" & strKokunaiEigyoId & "'"
    End Select
    
    'SQL文にWHERE句追加
    strSQL = strSQL & strSQL_WHERE
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDynaCount = gOraDataBase.CreateDynaset(strSQL, ORADYN_READONLY)
    
     'ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDynaCount.EOF = False Then
        '1件以上の場合
        If 1 <= GF_VarToNum(oraDynaCount![Cnt]) Then
            '件数セット
            intCnt = GF_VarToNum(oraDynaCount![Cnt])
        '該当データなしの場合
        Else
            'エラーにしない
            GF_CreateGroupList = True
            Exit Function
        End If
    End If
    
    Set oraDynaCount = Nothing
    
    'SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "  NGROUP_ID,"
    strSQL = strSQL & "  CGROUP,"
    strSQL = strSQL & "  CMODEL_KAITO"
    strSQL = strSQL & " FROM TCS_GROUP"
    strSQL = strSQL & strSQL_WHERE        'WHERE句追加
    strSQL = strSQL & " ORDER BY NDISPNO"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDynaData = gOraDataBase.CreateDynaset(strSQL, ORADYN_READONLY)

    '配列の要素数定義
    ReDim strGroupIdArray(intCnt - 1)
    
    '項目設定
    Do
        cntControl.AddItem GF_VarToStr(oraDynaData![CGROUP])
        'ID用配列内容区分
        Select Case intIdArrayKbn
            Case CGL_Id              'グループIDのみの場合
                strGroupIdArray(intIndex) = GF_VarToStr(oraDynaData![NGROUP_ID])
            Case CGL_IdAndModelKaito 'グループIDと機種マスタ回答部署区分)の場合
                'カンマつなぎで格納
                strGroupIdArray(intIndex) = GF_VarToStr(oraDynaData![NGROUP_ID]) & "," & GF_VarToStr(oraDynaData![CMODEL_KAITO])
        End Select
        
        intIndex = intIndex + 1
        oraDynaData.MoveNext
    Loop Until (oraDynaData.EOF = True)
    
    Set oraDynaData = Nothing
    
    '最後にもListIndex初期化
    cntControl.ListIndex = -1
    
    GF_CreateGroupList = True
    
    Exit Function
    
ErrHandler:
    'ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CreateGroupList", strSQL)
    
End Function

' 2016/12/15 ▲ M.Tanaka K545 CSプロセス改善  追加

' 2016/12/15 ▼ M.Tanaka K545 CSプロセス改善  追加
Private Function LF_GetIniTable(ByVal strKeyCd As String, ByVal intNumber As Integer) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:　設定値の取得
' 機能　　　:　INIファイルマスタから設定値を取得する
' 引数　　　:　strKeyCd As String       キーコード
' 　　　　　　 intNumber As Integer     順序
' 戻り値　　:　取得した値
'
' 機能説明　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim oraDyna     As OraDynaset
    Dim sSQL        As String

    LF_GetIniTable = ""

    'SQL文生成
    sSQL = ""
    sSQL = sSQL & "SELECT VCSET_CD FROM M68_INI_TABLE"
    sSQL = sSQL & " WHERE VCKEY_CD = '" & strKeyCd & "'"
    sSQL = sSQL & "   AND NNUMBER = " & intNumber

    'ﾀﾞｲﾅｾｯﾄ生成
    Set oraDyna = gOraDataBase.CreateDynaset(sSQL, ORADYN_READONLY Or ORADYN_NOCACHE)
    
    If oraDyna.EOF = False Then
        LF_GetIniTable = GF_VarToStr(oraDyna![VCSET_CD])
    End If
    
    Set oraDyna = Nothing

    Exit Function

ErrHandler:
    Call GS_ErrorHandler("LF_GetIniTable", sSQL)
End Function
' 2016/12/15 ▲ M.Tanaka K545 CSプロセス改善  追加

