Attribute VB_Name = "basGeneral"
' @(h) General.bas  ver1.00 ( 2000/09/28 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   プロジェクト名  : TLFﾌﾟﾛｼﾞｪｸﾄ
'   モジュール名    : basGeneral
'   ファイル名      : General.bas
'   Version        : 1.00
'   機能説明       ： 一般共通関数
'   作成者         ： T.Fukutani
'   作成日         ： 2000/09/28
'   修正履歴       ： 2001/03/19 T.Fukutani ｺﾓﾝﾀﾞｲｱﾛｸﾞ表示関数追加
'   　　　　       ： 2001/04/24 N.Kigaku FormLoad処理関数追加
'   　　　　       ： 2001/04/27 N.Kigaku 仕様設定NO採番処理,本機ＡＴＴ表示順取得関数追加
'   　　　　       ： 2001/11/30 N.Kigaku GF_GetShiyoKbn修正
'   　　　　       ： 2001/12/19 N.Kigaku GF_GetNextMitsumoriNo追加
'   　　　　       ： 2001/12/20 Takashi.Kato GF_FileOpenDialogに引数strFileTitle追加
'   　　　　       ： 2001/12/21 N.Kigaku GF_GetShiyoKbn_CIF追加
'   　　　　       ： 2002/01/23 N.Kigaku GS_Com_NextCntl,GF_FormInit修正
'   　　　　       ： 2002/02/27 N.Kigaku GF_ShowHelp追加,GF_GetNextMitsumoriNoとGF_NumberingShiyoNo修正
'   　　　　       ： 2002/03/12 N.Kigaku GF_GetNextMitsumoriNoとGF_NumberingShiyoNo修正
'   　　　　       ： 2002/03/30 T.Nono GF_FormInitとGS_Com_NextCntl修正
'   　　　　       ： 2005/06/17 N.KIGAKU GF_FileCopy追加
'                  ： 2006/12/05 N.Kigaku ｵﾗｸﾙ8.1.7 Nocache対応 検索時、ReadOnlyからNocacheに変更
'                  ： 2018/05/15 T.Nakayama K545 CSプロセス改善
'                  : 2021/04/26 R.Kozasa コメントコピー機能追加
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 環境宣言
'------------------------------------------------------------------------------
Option Explicit

Declare Function IsDBCSLeadByte Lib "kernel32" (ByVal bTestChar As Byte) As Long

'------------------------------------------------------------------------------
' パブリック定数宣言
'------------------------------------------------------------------------------
Public Const gINPUTCOLOR    As Long = &HC0FFC0   ''入力可能ｾﾙの色
Public Const gNOTINPUTCOLOR As Long = &H80000005 ''入力不可能ｾﾙの色
Public Const gLINECOLOR     As Long = &HFF0000   ''区切り線の色
Public Const gEXLSMAXROW    As Long = 60000      ''Excel抽出桁数
Public Const gTOTALCOLOR    As Long = &HC0FFFF   ''合計行の色

'------------------------------------------------------------------------------
' パブリック変数宣言
'------------------------------------------------------------------------------
Public gstrShiyoNo      As String   '仕様設定NO
Public gstrHonkiAttKbn  As String   '本機ATT区分
Public gstrRenban       As String   '連番

'------------------------------------------------------------------------------
' モジュール変数宣言
'------------------------------------------------------------------------------
Private mfrmFromName    As Form
'2021/04/26▼ R.Kozasa コメントコピー機能追加
'Private mcrtCntl(200)   As Control
Private mcrtCntl(210)   As Control
'2021/04/26▲ R.Kozasa コメントコピー機能追加
Private mstrFormName    As String


Public Sub GS_CenteringForm(frmMe As Form, Optional intOption As Integer = 0)
'------------------------------------------------------------------------------
' @(f)
' 機能名　　:　ﾌｫｰﾑｾﾝﾀﾘﾝｸﾞ
' 機能　　　:　ﾌｫｰﾑを画面の中心に移動する
' 引数　　　:　frmMe As Form   'ﾃｷｽﾄBOX
'                               0:画面中央
'                               1:上段中央
' 備考　　　:
'------------------------------------------------------------------------------
    
    Select Case intOption
    Case 0
        'ﾌｫｰﾑを中央に移動
        frmMe.Move (Screen.Width - frmMe.Width) / 2, (Screen.Height - frmMe.Height) / 2
    Case 1
        'ﾌｫｰﾑを上段中央に移動
        frmMe.Move (Screen.Width - frmMe.Width) / 2, 0
    End Select
End Sub


Public Function GF_FileOpenDialog(objObject As Object, strFilter As String, intFilterIndex As Integer, _
                                    strDir As String, strOpenFile As String, _
                                    Optional strFileTitle As String) As Integer
'------------------------------------------------------------------------------
' @(f)
' 機能名　　:　ﾌｧｲﾙ名指定ｺﾓﾝﾀﾞｲｱﾛｸﾞ表示
' 機能　　　:　ﾌｧｲﾙ名を指定して開く
' 引数　　　:　[i]objObject      As CommonDialog  'ｺﾓﾝﾀﾞｲｱﾛｸﾞｺﾝﾄﾛｰﾙ
' 　　　　　:　[i]strFilter      As String        'ﾌｧｲﾙﾌｨﾙﾀｰ文字列
' 　　　　　:　[i]intFilterIndex As Integer       'ﾌｨﾙﾀｰｲﾝﾃﾞｯｸｽ
' 　　　　　:　[i]strDir         As String        '初期表示ﾊﾟｽ
' 　　　　　:　[o]strOpenFile    As String        '場所(ﾊﾟｽ+ﾌｧｲﾙ名)
' 　　　　　:　[o]strFileTitle   As String        'ﾌｧｲﾙ名
' 戻り値　　:　[vbCancel] = ｷｬﾝｾﾙﾎﾞﾀﾝ押下時
' 備考　　　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    With objObject
        .DialogTitle = "ファイルを開く"
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
' 機能名　　:　ﾌｧｲﾙ名指定ｺﾓﾝﾀﾞｲｱﾛｸﾞ表示
' 機能　　　:　ﾌｧｲﾙ名を指定して保存
' 引数　　　:　[i]oObject As CommonDialog  'ｺﾓﾝﾀﾞｲｱﾛｸﾞｺﾝﾄﾛｰﾙ
' 　　　　　:　[i]sDir As String           '初期表示ﾊﾟｽ
' 　　　　　:　[i]sFile As String          '初期表示ﾌｧｲﾙ名
' 　　　　　:　[o]sSaveFile As String      '保存場所(ﾊﾟｽ+ﾌｧｲﾙ名)
' 戻り値　　:　[vbCancel] = ｷｬﾝｾﾙﾎﾞﾀﾝ押下時
' 備考　　　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    With oObject

        oObject.DialogTitle = "出力先指定"
        oObject.Filter = "CSVﾌｧｲﾙ(.CSV)|*.CSV"
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
' 機能名 : コントロール情報の取得
' 機能   : 指定フォームのコントロール情報を取得します
'          ※複数の画面が存在する場合は画面が切り替わる度に GF_FormInit を再び Call して下さい
' 引数   : frmForm As Form  フォームコントロール
' 備考   : GF_FormInitを複数回 Call しても何もしません。但し、コントロール配列を使っている場合は不可。
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
    For intM_Idx = 0 To (mfrmFromName.Count - 1)        '配列にコントロールをTabIndex順に設定します
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
                    Set mcrtCntl(intCount) = mfrmFromName.Controls(intloop)      'TabIndex順に内部コントロール配列を設定します
                    intCount = intCount + 1
                    Exit For
                End If
            End Select
        Next intloop
    Next intM_Idx
    Set mcrtCntl(intCount) = Nothing        '内部コントロール配列の最終情報を設定します
    
End Sub

Public Sub GS_Com_NextCntl(crtControl As Control)
'------------------------------------------------------------------------------
' @(f)
' 機能名 : フォーカス移動
' 機能   : GF_FormInit関数で指定したフォームにて、指定するコントロール(crtControl)の次(TabIndex順)のコントロールにfocusをあわせます
' 引数   : crtControl As Control  元になるコントロール
' 備考   :
'------------------------------------------------------------------------------
    Dim strControlName As String
    Dim intMK_Idx      As Integer
    Dim intCount       As Integer
    Dim intLoopExit  As Integer
    
    If mfrmFromName Is Nothing = True Then
        SendKeys "{TAB}"       'TAB ｷ- SEND(Next Field Cursol)
        Exit Sub
    Else
        If mstrFormName <> Screen.ActiveForm.Name Then
            SendKeys "{TAB}"       'TAB ｷ- SEND(Next Field Cursol)
            Exit Sub
        End If
    End If
            
    intMK_Idx = 0
    Do While mcrtCntl(intMK_Idx) Is Nothing = False
        If mcrtCntl(intMK_Idx).Name = crtControl.Name Then
            intCount = intMK_Idx
'2002/01/23 Delete N.Kigaku
'ｺﾝﾄﾛｰﾙ配列もﾁｪｯｸするためｺﾒﾝﾄ
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
        If mcrtCntl(intCount) Is Nothing = True Then       '最後のコントロールだったらもう一度最初のコントロールからループ
            If intLoopExit <> 1 Then                     '永久ループ抑止対応
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
' 機能名    :   3月までの月数算出
' 機能      :   同年度の3月までの月数を求める
' 引数      :   strMonth As String  月(YYYY/MM/DD or YYYY/MM)
' 戻り値    :   Integer   月数
' 備考      :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strYear  As String
    Dim strMarch As String
    
    GF_MonthCount = 0
    
    '年取得
    strYear = Left(strMonth, 4)
    
    '1日日付に
    strMonth = Left(strMonth, 7) & "/01"
    
    If CInt(Mid(strMonth, 6, 2)) > 3 Then
        strMarch = CStr(CInt(strYear) + 1) & "/03/01"
    Else
        strMarch = strYear & "/03/01"
    End If
    
    GF_MonthCount = DateDiff("M", strMonth, strMarch)
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_MonthCount")
    
End Function

Public Function GF_Year(Optional strYear As String = "NoDate", Optional strMonth As String = "NoDate") As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   任意年月の年度算出(引数省略時はｼｽﾃﾑ日付の年度を返す)
' 機能      :
' 引数      :   strYear As String      '年(省略可)
'               strMonth As String     '月(省略可)
' 戻り値    :   String     年度
' 備考      :   2000/12/11
'------------------------------------------------------------------------------
    Dim intYear   As Integer
    Dim intMonth  As Integer
    Dim strDate   As String
    
    '引数省略時、運用年月の年度算出
    If strYear = "NoDate" Or strMonth = "NoDate" Then
        strDate = Screen.ActiveForm.lblNowDate
        strYear = Left(strDate, 4)
        strMonth = Mid(strDate, 6, 2)
    End If
    
    '年度算出処理
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
' 機能名　　:　小数点切上げ関数
' 機能　　　:　小数点以下指定桁を切上げる
' 引数　　　:　strNum   切上げ対象値
' 　　　　　:　intPoint 切上げ桁位置(小数点以下切上げの場合 = 0)
' 戻り値　　:　切上げ後値(切上げ桁位置に2を指定した場合は0.621が0.63)
' 備考　　　:　[有効値] 切上げ桁位置+精度の範囲内の小数点値
' 　　　　　:　[限界値] 整数部小数部合計29桁を超えるとｵｰﾊﾞｰﾌﾛｰします
' 　　　　　:　精度を上げるため変数は文字列型を使用し、内部処理形式は10進型(DECIMAL)
'------------------------------------------------------------------------------
    Dim strTemp    As String          '作業領域
    Const intSeido As Integer = 5     '精度(切上げ桁以降の有効範囲)

    On Error GoTo ErrHandler
    
    GF_CutUp = strNum
    
    ''数値ﾁｪｯｸ
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
    'ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CutUp")

End Function

Public Function GF_WCardChenge(strString As String, intLength As Integer) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   ワイルドカードの作成
' 機能      :
' 引数      :   strString As String      '文字列
'               intLength As Integer     '文字数
' 戻り値    :   String     変換後の文字列
' 備考      :   ｢*｣ を｢_｣ に変化する、文字数が足りないものは足りない文字数分「_」を追加
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
' 機能名 : 分類名称取得関数
' 機能　 : 分類コードに該当する分類名称を取得する。
' 引数　 : strBunruiFlg  As String     分類区分(1：分類１, 2:分類２, 3:分類３)
' 　　　   strBunrui1    As String     分類ｺｰﾄﾞ1
' 　　　   strBunrui2    AS String     分類ｺｰﾄﾞ2
' 　　　   strBunrui3    AS String     分類ｺｰﾄﾞ3
' 戻り値 : String   分類名称
' 備考   :
'------------------------------------------------------------------------------
'   変数定義
    Dim strSQL        As String
    Dim oDynaset      As OraDynaset
    On Error GoTo ErrHandler
'   ｸﾘｱ
    GF_Com_BunRuiName = ""
'   検索対象DB設定
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
'   ﾚｺｰﾄﾞがない場合ｴﾗｰ
    If oDynaset.EOF = True Then
        Exit Function
    End If
'   分類名称ｾｯﾄ
    GF_Com_BunRuiName = GF_VarToStr(oDynaset![BUNRUINAME])
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_Com_BunName", strSQL)
    GF_Com_BunRuiName = " "
    
End Function

Public Function GF_LoadFormProcess(frm As Form, objMe As Object)
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:FormLoad処理
' 機能　　　:フォームロード(画面遷移)処理
' 引数　　　:frm    As Form     フォームオブジェクト
'　　　　　 :objMe  As Object　　Meオブジェクト
' 機能説明　:フォームロード(画面遷移)処理
'------------------------------------------------------------------------------
    Load frm
    'フォームロード完了判定
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
' 機能名　　:ｼﾝｸﾞﾙｸｫｰﾃｰｼｮﾝの2重化
' 機能　　　:
' 引数　　　:strString As String (in)    文字列
' 戻り値　　:ｼﾝｸﾞﾙｸｫｰﾃｰｼｮﾝを2重化した後の文字列
' 機能説明　:
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
' 機能名　　:仕様設定NO採番処理
' 機能　　　:仕様設定NO採番処理
' 引数　　　:strShiyoNo As String (out)    仕様設定NO
'　　　　　　intShiyuKbn As Integer (in)   市輸区分　1:国内、2:海外
' 機能説明　:仕様設定NOの採番を行う (注意：トランザクションは呼び元で行う)
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
    
    strMsgTitle = "仕様設定NO採番処理"
    
    ''◆① システム日付の取得◆
    strSQL = ""
    strSQL = strSQL & "SELECT TRUNC(SYSDATE) SYSTEMDATE"
    strSQL = strSQL & "  FROM DUAL"
    
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oDynaset.EOF Then
        'システム日付の取得に失敗しました。
        Call GF_MsgBoxDB(strMsgTitle, "WTG027", "OK", "E")
        Exit Function
    Else
        strWareki = Format(CStr(oDynaset![SYSTEMDATE]), "e")
        strSysSeireki = Format(CStr(oDynaset![SYSTEMDATE]), "yyyy")
    End If
    
    ''◆② 仕様設定NOと年度の取得◆
    strSQL = ""
    strSQL = strSQL & "SELECT NVL(SYNODOME,0)  SYNODOME"
'    strSQL = strSQL & "      ,NVL(SYNOFORE,0)  SYNOFORE"
    strSQL = strSQL & "      ,NVL(NENDO,'0')   NENDO"
    strSQL = strSQL & "  FROM THJSIYOSETNO"
    
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oDynaset.EOF Then
        ''該当ﾃﾞｰﾀが無い時は追加する
        strSQL = ""
        strSQL = strSQL & "INSERT INTO THJSIYOSETNO"
'        strSQL = strSQL & " (SYNODOME,SYNOFORE,NENDO,SEIREKI)"
'        strSQL = strSQL & " (SYNODOME,MTNO,NENDO,SEIREKI)"
        strSQL = strSQL & " (SYNODOME,NENDO)"
        strSQL = strSQL & "VALUES"
'        If intShiyuKbn = 2 Then
'            strSQL = strSQL & " (0,1,'" & strWareki & "','" & strSysSeireki & "')"   '海外
'        Else
            strSQL = strSQL & " (1,'" & strWareki & "')"   '国内
'        End If
        Set oraDSQLStmt = gOraDataBase.CreateSql(strSQL, ORADYN_NO_AUTOBIND)
        If oraDSQLStmt.RecordCount = 0 Then
            '更新失敗
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
        
        ''◆③ 採番処理◆
        '採番テーブルの和暦・西暦よりシステムの和暦が大きい時または仕様設定ＮＯが最大値に達した時に
        '仕様設定NOをクリアしてシステムの和暦を登録する
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
        '        '国内
                strField = "SYNODOME=SYNODOME+1"
        '    Else
        '        '海外
        '        strField = "SYNOFORE=SYNOFORE+1"
        '    End If
            strSQL = ""
            strSQL = strSQL & "UPDATE THJSIYOSETNO SET " & strField
            strSQL = strSQL & " WHERE NENDO='" & strNendo & "'"
        End If
        
        Set oraDSQLStmt = gOraDataBase.CreateSql(strSQL, ORADYN_NO_AUTOBIND)
        If oraDSQLStmt.RecordCount = 0 Then
            '更新失敗
            Call GF_MsgBoxDB(strMsgTitle, "WTG028", "OK", "E")
            Exit Function
        End If
        
        ''◆④ 採番後の仕様設定Noを取得する
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
            '更新失敗
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
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_NumberingShiyoNo", strSQL)
End Function

Public Function GF_GetHonkiAttHyojiNo(ByVal strSYNO As String, ByRef strHyoji As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:本機ＡＴＴ表示順取得
' 機能　　　:本機ＡＴＴの表示順を取得する
' 引数　　　:strSYNO  As String  (in)    仕様設定NO
'           strHyoji As String (out)    表示順
' 機能説明　:本機ＡＴＴの表示順を取得する
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
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_GetHonkiAttHyojiNo", strSQL)
End Function

Public Function GF_GetNextMitsumoriNo(ByRef strMITSUMORINO As String, ByVal strHanbaitenNo As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:見積NO取得処理
' 機能　　　:見積NO取得処理
' 引数　　　:strMitsumoriNo As String (out)  見積No
'　　　　　 :strHanbaitenNo As String (in)   販売店ｺｰﾄﾞ(5桁) + 営業所ｺｰﾄﾞ(2桁)
' 機能説明　:販売店を受取り、次に使用する見積NOを返す
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
    
    strMsgTitle = "見積NO取得処理"
    
    ''◆① システム日付の取得◆
    strSQL = ""
    strSQL = strSQL & "SELECT TRUNC(SYSDATE) SYSTEMDATE"
    strSQL = strSQL & "  FROM DUAL"
    
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oDynaset.EOF Then
        'システム日付の取得に失敗しました。
        Call GF_MsgBoxDB(strMsgTitle, "WTG027", "OK", "E")
        Exit Function
    Else
        strWareki = Format(CStr(oDynaset![SYSTEMDATE]), "e")
        strSysSeireki = Format(CStr(oDynaset![SYSTEMDATE]), "yyyy")
    End If
    
    ''◆② 見積NOと年度の取得◆
    strSQL = ""
    strSQL = strSQL & "SELECT NVL(MTNO,0)      MTNO"
    strSQL = strSQL & "      ,NVL(SEIREKI,'0') SEIREKI"
    strSQL = strSQL & "  FROM THJSIYOSETNO"
    
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oDynaset.EOF Then
        ''該当ﾃﾞｰﾀが無い時は追加する
        strSQL = ""
        strSQL = strSQL & "INSERT INTO THJSIYOSETNO"
        strSQL = strSQL & " (MTNO,SEIREKI)"
        strSQL = strSQL & "VALUES"
        strSQL = strSQL & " (1,'" & strSysSeireki & "')"
        Set oraDSQLStmt = gOraDataBase.CreateSql(strSQL, ORADYN_NO_AUTOBIND)
        If oraDSQLStmt.RecordCount = 0 Then
            '更新失敗
            Call GF_MsgBoxDB(strMsgTitle, "WTG028", "OK", "E")
            Exit Function
        End If
        strMITSUMORINO = strHanbaitenNo & strSysSeireki & "00001"
        GF_GetNextMitsumoriNo = True
        Exit Function
        
    Else
        strSeireki = oDynaset![SEIREKI]
        strWkMitsumori = oDynaset![MTNO]
        
        ''◆③ 採番処理◆
        
        '採番テーブルの西暦よりシステムの西暦が大きい時または見積Noが最大値に達した時に
        '見積Noをクリアしてシステムの西暦を登録する
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
            '更新失敗
            Call GF_MsgBoxDB(strMsgTitle, "WTG028", "OK", "E")
            Exit Function
        End If
        
        ''◆④ 採番後の見積Noを取得する
        strSQL = ""
        strSQL = strSQL & "SELECT NVL(MTNO,'1') MTNO"
        strSQL = strSQL & "      ,NVL(SEIREKI,'1') SEIREKI"
        strSQL = strSQL & "  FROM THJSIYOSETNO"
        Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
        If oDynaset.EOF Then
            '更新失敗
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
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_GetNextMitsumoriNo", strSQL)
End Function

Public Function GF_GetShiyoKbn(ByVal strShiyoNo As String, ByVal strMsgTitle As String, _
                               ByRef strAndShiyuKbn As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:  市輸区分取得
' 機能　　　:  仕様設定Noから市輸区分を取得する
' 引数　　　:  strShiyoNo      As String  仕様設定No
'             strMsgTitle     As String  エラーメッセージタイトル
'             strAndShiyuKbn  As String  市輸区分
' 機能説明　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strSQL     As String      ''SQL文
    Dim oraDyna    As OraDynaset  ''ﾀﾞｲﾅｾｯﾄ
    Dim intRet     As Integer
    
    GF_GetShiyoKbn = False
    
    strAndShiyuKbn = ""
    
    ''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT CIF.SHIYUKBN FROM THJMR MR,THJCIF CIF"
    strSQL = strSQL & "    WHERE MR.CIFNO = CIF.CIFNO"
    strSQL = strSQL & "      AND MR.EIGYONO = CIF.EIGYONO"
'2001/11/30 Added By Kigaku
    strSQL = strSQL & "      AND MR.SYNO = '" & strShiyoNo & "'"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = False Then
        strAndShiyuKbn = GF_VarToStr(oraDyna![SHIYUKBN])
        GF_GetShiyoKbn = True
        Exit Function
    End If
    
    'ﾀﾞｲﾅｾｯﾄの解放
    Set oraDyna = Nothing
    
'2001/12/27 Delete 海外は使用しないため一時的に削除
'    ''SQL文
'    strSQL = ""
'    strSQL = strSQL & "SELECT CIF.SHIYUKBN FROM THJYSMR MR,THJCIF CIF"
'    strSQL = strSQL & "    WHERE MR.CIFNO = CIF.CIFNO"
'    strSQL = strSQL & "      AND MR.EIGYONO = CIF.EIGYONO"
''2001/11/30 Added By Kigaku
'    strSQL = strSQL & "      AND MR.SYNO = '" & strShiyoNo & "'"
'
'    'ﾀﾞｲﾅｾｯﾄの生成
'    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
'
'    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
'    If oraDyna.EOF = False Then
'        strAndShiyuKbn = GF_VarToStr(oraDyna![SHIYUKBN])
'        GF_GetShiyoKbn = True
'        Exit Function
'    End If
'
'    'ﾀﾞｲﾅｾｯﾄの解放
'    Set oraDyna = Nothing
    
    If strAndShiyuKbn = "" Then
        intRet = GF_MsgBoxDB(strMsgTitle, "WTG001", "OK", "E")
        Exit Function
    End If

    GF_GetShiyoKbn = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_GetShiyoKbn", strSQL)

End Function


Public Function GF_GetShiyoKbn_CIF(ByVal strCifNO As String, ByVal strEigyoNo As String _
                                , ByRef strShiyuKbn As String, Optional ByVal strMsgTitle As String = "市輸区分取得") As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名　　:   市輸区分取得
' 機能　　　:   販売店ｺｰﾄﾞと営業所ｺｰﾄﾞから市輸区分を取得する
' 引数　　　:  strCifNo As String   販売店ｺｰﾄﾞ
'             strEigyoNo As String 営業所ｺｰﾄﾞ
'             strShiyuKbn  As String  市輸区分   '市輸区分 '1'ｽﾍﾟｰｽ:国内、'2'海外
'             strMsgTitle As String   ﾒｯｾｰｼﾞﾀｲﾄﾙ
' 機能説明　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strSQL     As String      ''SQL文
    Dim oraDyna    As OraDynaset  ''ﾀﾞｲﾅｾｯﾄ
    Dim intRet     As Integer
    Dim strMsg     As String
    
    GF_GetShiyoKbn_CIF = False
    
    strShiyuKbn = ""
    
    ''SQL文
    strSQL = ""
    strSQL = strSQL & "SELECT SHIYUKBN FROM THJCIF"
    strSQL = strSQL & " WHERE CIFNO = '" & strCifNO & "'"
    strSQL = strSQL & "   AND EIGYONO = '" & strEigyoNo & "'"
    
    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''ﾃﾞｰﾀ存在ﾁｪｯｸ
    If oraDyna.EOF = False Then
        strShiyuKbn = GF_VarToStr(oraDyna![SHIYUKBN])
        GF_GetShiyoKbn_CIF = True
        Exit Function
    Else
'        intRet = GF_MsgBoxDB(strMsgTitle, "WTH134", "OK", "E")
        strMsg = GF_GetMsg("WTH134")
        strMsg = strMsg & vbCr & "販売店ｺｰﾄﾞ：" & strCifNO
        strMsg = strMsg & vbCr & "営業所ｺｰﾄﾞ：" & strEigyoNo
        intRet = GF_MsgBox(strMsgTitle, strMsg, "OK", "E")
        Exit Function
    End If
    
    'ﾀﾞｲﾅｾｯﾄの解放
    Set oraDyna = Nothing
    
    GF_GetShiyoKbn_CIF = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_GetShiyoKbn_CIF", strSQL)

End Function

Public Function GF_CutCharLength(ByRef strMoji As String, ByVal intCutLngth As Integer) As String
''------------------------------------------------------------------------------
'' @(f)
''
'' 機能名    :   文字列を指定バイト数で切り取る
'' 機能      :
'' 引数      :   strMoji As String      (in/out) '文字列
''               intCutLngth As Integer (in)     '文字数
'' 戻り値    :   String     切り出した後の文字列
'' 備考      :   strMoji には切り出した文字以降の文字列が戻る
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
'' 機能名    :　　ヘルプ画面表示
'' 機能      :
'' 引数      :   strLinkID As String      (in/out) '画面リンクID
'' 戻り値    :
'' 備考      :
''------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim lngRet    As Long
    Dim blnErrFlg As Boolean
    
    GF_ShowHelp = False
    
    blnErrFlg = False
    
    On Error Resume Next
    
    'ヘルプ起動プログラムの有無チェック
    If Dir(App.Path & "\help.exe") = "" Then
        blnErrFlg = True
    End If
    If Err.Number <> 0 Then
        blnErrFlg = True
    End If
    Err.Clear
    If blnErrFlg = True Then
        lngRet = GF_MsgBoxDB("ヘルプ", "WTG040", "OK", "E")
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    'へルプの起動
    lngRet = Shell(App.Path & "\help.exe " & strLinkID, vbNormalFocus)
    DoEvents
    
    GF_ShowHelp = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_ShowHelp")
End Function

Public Function GF_FileCopy(ByVal strSource As String _
                              , ByVal strDestination As String) As Boolean
''--------------------------------------------------------------------------------
'' @(f)
''
' 機能名　　:　ﾌｧｲﾙｺﾋﾟｰ
' 機能　　　:　ﾌｧｲﾙをｺﾋﾟｰする
' 引数　　　:　strSource        As String     ''ｺﾋﾟｰ元ﾌｧｲﾙ名
' 　　　　　 　strDestination   As String     ''ｺﾋﾟｰ先ﾌｧｲﾙ名
'
'' 戻り値   : TRUE：正常 FALSE:エラー Boolean
''--------------------------------------------------------------------------------
On Error GoTo ErrHandler
    
    GF_FileCopy = False
    
    Call FileCopy(strSource, strDestination)
    
    GF_FileCopy = True
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_FileCopy", "ｺﾋﾟｰ元:" & strSource & " , ｺﾋﾟｰ先:" & strDestination)
End Function

' 2018/05/15 ▼ T.Nakayama K545 CSプロセス改善
Public Function GF_Chk_AutoModelWork(ByVal strAutoTypeFlg As String _
                                    , ByVal strTSDRNo As String _
                                    , ByVal strAcceptNo As String _
                                    , ByVal strDeliveryNo As String _
                                    , ByVal strHonkiAttKubun As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'　機能名　: 自動適用範囲ﾜｰｸ存在ﾁｪｯｸ
'　機能　　:
'　引数　　: strAutoTypeFlg     As String      (in)  自動種別ﾌﾗｸﾞ
'　    　　: strTsdrNo          As String      (in)  仕様設定NO
'　    　　: strAcceptNo        As String      (in)  受注NO
'　    　　: strDeliveryNo      As String      (in)  引合納期システムNO
'　    　　: strHonkiAttKubun   As String      (in)  本機ATT区分
'　戻り値　: True = 有り / False = 無し
'　備考　　:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL          As String
    Dim oraDyna         As OraDynaset
    '対象テーブル
    Dim strTaisyoTable  As String

    GF_Chk_AutoModelWork = False
    
    '対象テーブルの設定
    Select Case strAutoTypeFlg
    '仕様設定の場合
    Case 1
        Select Case strHonkiAttKubun
        'ATTの場合
        Case 1
            '対象テーブル:自動適用範囲ﾜｰｸ(仕様設定)(自動)(ATT)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_THJ_ATT"
        '本機の場合
        Case 2
            '対象テーブル:自動適用範囲ﾜｰｸ(仕様設定)(自動)(本機)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_THJ_H"
        End Select
    '引合納期の場合
    Case 2
        Select Case strHonkiAttKubun
        'ATTの場合
        Case 1
            '対象テーブル:自動適用範囲ﾜｰｸ(引合納期)(自動)(ATT)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_INQ_ATT"
        '本機の場合
        Case 2
            '対象テーブル:自動適用範囲ﾜｰｸ(引合納期)(自動)(本機)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_INQ_H"
        End Select
    '仕決処理の場合
    Case 3
        Select Case strHonkiAttKubun
        'ATTの場合
        Case 1
            '対象テーブル:自動適用範囲ﾜｰｸ(自動)(ATT)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_ATT"
        '本機の場合
        Case 2
            '対象テーブル:自動適用範囲ﾜｰｸ(自動)(本機)
            strTaisyoTable = "  FROM TCS_AUTO_WORK_H"
        End Select
    End Select

    '受注NO
    strAcceptNo = IIf(strAcceptNo <> "", strAcceptNo, "            ")
    '引合納期システムNO
    strDeliveryNo = IIf(strDeliveryNo <> "", strDeliveryNo, "          ")

    strSQL = ""
    strSQL = strSQL & " SELECT VCTSDR_NO"
    strSQL = strSQL & strTaisyoTable
    strSQL = strSQL & "  WHERE VCTSDR_NO = '" & Trim(strTSDRNo) & "'"
    strSQL = strSQL & "    AND CACCEPTNO = '" & strAcceptNo & "'"
    strSQL = strSQL & "    AND NDELIVERY_NO = '" & strDeliveryNo & "'"

    'ﾀﾞｲﾅｾｯﾄの生成
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oraDyna.EOF = False Then
        GF_Chk_AutoModelWork = True
    End If
    Set oraDyna = Nothing

    Exit Function
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_Chk_AutoModelWork", strSQL)
End Function
' 2018/05/15 ▲ T.Nakayama K545 CSプロセス改善
