Attribute VB_Name = "basInputCheck"
' @(h) basInputCheck.bas  ver 1.00 ( 2000/08/29 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   プロジェクト名  : L&Fﾌﾟﾛｼﾞｪｸﾄ
'   モジュール名    : basInputCheck
'   ファイル名      : basInputCheck.bas
'   Version        : 1.00
'   機能説明       ： 入力ﾁｪｯｸに関する共通関数
'   作成者         ： T.Fukutani
'   作成日         ： 2000/12/01
'   修正履歴　　　  ： 2001/05/15  GF_Com_KeyPressに英小文字→英大文字を追加
'   　　　　　　　     2001/12/11  GF_Com_KeyPress14,15を追加 <T.Matsui>
'   　　　　　　　     2002/01/08  GF_Com_KeyPressを使用するGF_Com_CheckStringを作成 <T.Matsui>
'   　　　　　　　     2002/01/10  GF_ChangeQuateSingを作成 <N.Kigaku>
'   　　　　　　　     2002/01/17  GF_ReplaceAmperを作成 <N.Kigaku>
'   　　　　　　　     2002/04/09  GF_Com_KeyPress16を追加 <N.Kigaku>
'   　　　　　　　     2002/07/11  GS_Com_TxtGotFocusに'chk'を追加 <N.Kigaku>
'                      2002/08/27  GF_MinusCheckを追加 <N.Kigaku>
'                      2002/10/23  GF_FileNameRestrinctionを追加 <N.Kigaku>
'                      2005/01/11  GF_DateConv追加 <N.Kigaku>
'                      2005/12/28  GF_THJCMBXMR_CHK追加 <N.Kigaku>
'                      2006/01/06  GF_CheckNumber2,GF_ChkDeci追加 <N.Kigaku>
'                      2006/01/19  GF_Com_KeyPressに許可ﾊﾟﾀｰﾝ"17"とﾘﾀｰﾝﾁｪｯｸﾌﾗｸﾞを追加,
'                                  GF_OptFormatChk,定数[ｵﾌﾟｼｮﾝ桁数,ｻｲｽﾞｵﾌﾟｼｮﾝ桁数]追加 <N.Kigaku>
'                      2006/01/31  GF_DateConvに全角ﾁｪｯｸ追加 <N.Kigaku>
'                      2006/07/07  GF_UndoAmperを追加 <N.Kigaku>
'                      2006/12/05  ｵﾗｸﾙ8.1.7 Nocache対応 検索時、ReadOnlyからNocacheに変更 <N.Kigaku >
'                      2006/12/08  改行ｺｰﾄﾞﾁｪｯｸ関数[GF_CheckLinefeed]追加 <N.Kigaku>
'                      2006/12/11  GF_CheckEngNumMark追加 <N.Kigaku>
'                      2008/02/21  GF_CharPermitChekの許可ﾊﾟﾀｰﾝ12のﾊｲﾌﾝﾁｪｯｸ修正
'                      2008/05/27  GF_ChangeQuateDouble関数追加 <N.Kigaku>
'                      2009/03/16  GF_Com_KeyPressに許可ﾊﾟﾀｰﾝ"18"を追加 <N.Kigaku>
'                      2011/05/20  GF_CharPermitChekに許可ﾊﾟﾀｰﾝ"15","16"を追加 <N.Kigaku>
'                      2016/11/14  GF_Com_KeyPressに制御ｺｰﾄﾞ有効ﾊﾟﾀｰﾝを追加、GF_Com_CheckStringにﾘﾀｰﾝﾁｪｯｸﾌﾗｸﾞを追加
'                      2017/01/11  M.Tanaka K545 CSプロセス改善 GF_CheckStartToEnd追加
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 環境宣言
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
'  定数宣言
'------------------------------------------------------------------------------
Private Const mlngOptLength = 4         'ｵﾌﾟｼｮﾝ桁数
Private Const mlngSizeOptLength = 8     'ｻｲｽﾞｵﾌﾟｼｮﾝ桁数

' 2017/01/11 ▼ M.Tanaka K545 CSプロセス改善  ADD
Public Enum CSTE_ChkKbn 'チェック区分
    CSTE_Year = 1           '年
    CSTE_Month = 2        '月
    CSTE_Date = 3           '日
End Enum
' 2017/01/11 ▲ M.Tanaka K545 CSプロセス改善  ADD

Public Function GF_Com_CheckString(intPatan As Integer, _
                                   strCharCode As String, _
                          Optional bolAsterisk As Boolean = False, _
                          Optional bolEnterChkFlg As Boolean = True) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 文字列コードチェック
' 機能　 : 許可されたコード以外が文字列に含まれていないかチェックを行う
' 引数　 : intPatan As Integer        ''許可ﾊﾟﾀｰﾝ
' 　　　                                (GF_Com_KeyPressに準拠)
' 　　　   strCharCode As String      ''文字列(チェックOKの場合、大文字化など変換後の値を返す)
' 　　　   bolAsterisk As Boolean     ''ｱｽﾀﾘｽｸ許可ﾌﾗｸﾞ(省略時False=不可)
'          bolEnterChkFlg As Boolean  ''ﾘﾀｰﾝﾁｪｯｸﾌﾗｸﾞ(省略時True=有)
' 戻り値 : Boolean                    ''True･･･チェックOK     False･･･チェックNG
' 備考　 : ※常にGF_Com_KeyPressに準拠していないと意味がない(引数なども)
'------------------------------------------------------------------------------
    Dim l                           As Long
    Dim intRet                      As Integer
    Dim intAscii                    As Integer
    Dim sString                     As String
    
    GF_Com_CheckString = True
    
    '1文字1文字、文字列の長さ分チェック
    sString = ""
    For l = 1 To Len(strCharCode)
        
        'アスキーコードに変換
        intAscii = Asc(Mid$(strCharCode, l, 1))
        
        '入力キーコードチェック
'2016/11/14 LF1667_TIK担当メールアドレス欄の桁拡張 START <<<
'        intRet = GF_Com_KeyPress(intPatan, intAscii, bolAsterisk)
        intRet = GF_Com_KeyPress(intPatan, intAscii, bolAsterisk, bolEnterChkFlg)
'2016/11/14 LF1667_TIK担当メールアドレス欄の桁拡張 END >>>
        If (intRet = 0) Then
            '無効な文字が見つかった場合
            GF_Com_CheckString = False
            Exit Function
        End If
        
        '文字に戻して格納
        sString = sString & Chr$(intAscii)
    Next l
    
    '大文字化など変換後文字列を返す
    strCharCode = sString
    
End Function

Public Function GF_Com_KeyPress(intPatan As Integer, _
                                intKeyAscii As Integer, _
                       Optional bolAsterisk As Boolean = False, _
                       Optional bolEnterChkFlg As Boolean = True, _
                       Optional intCtrlPatan As Integer = 0) As Integer
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 入力キーコードチェック
' 機能   : 入力許可されたキーコード以外のものをはじく
' 引数   : intPatan As Integer     ''許可ﾊﾟﾀｰﾝ
'                0  - Keypress Code Non Check
'                1  - 数字  Code Non Check  "0,1,2,～9"
'                2  - 数字＋ﾋﾟﾘｵﾄﾞ Code Non Check   "0,1,2,～9,.,"
'                3  - 数字＋ﾋﾟﾘｵﾄﾞ＋ﾏｲﾅｽ Code Non Check   "0,1,2,～9,.,-"
'                4  - 数字＋英字 Code Non Check   "0,1,2,～9,A～Z"
'                5  - カナ漢字
'                6  - 数字＋ﾏｲﾅｽ Code Non Check   "0,1,2,～9,-"
'                7  - 数字＋英字 Code Non Check   "0,1,2,～9,A～Z,a～z,"
'                8  - '!' ～ '}' までOK    (ｺｰﾄﾞだと 33 ～ 125まで)
'                9  - 英数字＋ﾌﾟﾗｽ＋ﾏｲﾅｽ＋"*" Code Non Check   "0,1,2,～9,A～Z,a～z,+,-,*"
'                10 - 英大文字 Code Non Check   "A～Z"
'                11 - 数字＋ﾊｲﾌﾝ＋ｽﾗｯｼｭ Code Non Check   "0,1,2,～9,-,/"
'                12 - 数字＋ﾊｲﾌﾝ＋ｶｯｺ Code Non Check   "0,1,2,～9,-,(,)"
'                13 - 英小文字 → 英大文字      "a～z""
'                14 - 数字＋英字＋ﾊｲﾌﾝ Code Non Check   "0,1,2,～9,A～Z","-"
'                15 - 数字＋英字＋ﾌﾞﾗﾝｸ Code Non Check   "0,1,2,～9,A～Z"," "
'                16 - 数字＋ﾌﾞﾗﾝｸ Code Non Check   "0,1,2,～9," "
'                17 - ASCIIｺｰﾄﾞ(0～127)の制御文字(ﾘﾀ-ﾝ,ﾗｲﾝﾌｨ-ﾙﾄﾞ除く)以外全て許可
'                18 - ASCIIｺｰﾄﾞ(0～127)+拡張ASCIIｺｰﾄﾞ(128～255)の制御文字以外全て許可
'          intKeyAscii As Integer      ''ｱｽｷｰｺｰﾄﾞ
'          bolAsterisk As Boolean      ''ｱｽﾀﾘｽｸ許可ﾌﾗｸﾞ(省略時False=不可)
'          bolEnterChkFlg              ''ﾘﾀｰﾝﾁｪｯｸﾌﾗｸﾞ(省略時True=有)
'          intCtrlPatan As Integer     ''制御ｺｰﾄﾞ有効ﾊﾟﾀｰﾝ
'                 0 - ﾁｪｯｸ無し
'                 1 - Ctrl+C,Ctrl+V
'                 2 - Ctrl+C,Ctrl+V,Ctrl+X
'                 3 - Ctrl+C,Ctrl+V,Ctrl+X,Ctrl+Z
' 戻り値 : Integer          ''0･･･ｱｽｷｰｺｰﾄﾞは無効 　 1･･･ｱｽｷｰｺｰﾄﾞは有効
' 備考   :
'------------------------------------------------------------------------------
    GF_Com_KeyPress = 0    'ｷｰｺｰﾄﾞは無効

    ''ﾀﾌﾞ(9)  ﾊﾞｯｸｽﾍﾟ-ｽ(8) ｷ- ﾁｪｯｸ   ===>制御無し
    If (intKeyAscii = 9) Then
        Exit Function
    ElseIf (intKeyAscii = 8) Then
        GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
        Exit Function
    End If
    ''ﾘﾀ-ﾝ(13) ﾗｲﾝﾌｨ-ﾙﾄﾞ(10) ｷ- ﾁｪｯｸ   ===>制御無し
    If (bolEnterChkFlg = True) And (intKeyAscii = 13 Or intKeyAscii = 10) Then
        intKeyAscii = 0     'こうしないとBEEP音が鳴るため
        Exit Function
    End If
    
    'ｱｽﾀﾘｽｸ入力許可時、*(42)ｷｰﾁｪｯｸ
    If bolAsterisk = True And intKeyAscii = 42 Then
        GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
        Exit Function
    End If

'2016/11/14 LF1667_TIK担当メールアドレス欄の桁拡張 START <<<
    '制御コードで以下のものは有効とする
    Select Case intCtrlPatan
        Case 0
        Case 1
            'Ctrl+C(3), Ctrl+V(22)
            If (intKeyAscii = 3) Or (intKeyAscii = 22) Then
                GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
                Exit Function
            End If
        
        Case 2
            'Ctrl+C(3), Ctrl+V(22), Ctrl+X(24)
            If (intKeyAscii = 3) Or (intKeyAscii = 22) Or (intKeyAscii = 24) Then
                GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
                Exit Function
            End If
        
        Case 3
            'Ctrl+C(3), Ctrl+V(22), Ctrl+X(24), Ctrl+Z(26)
            If (intKeyAscii = 3) Or (intKeyAscii = 22) Or (intKeyAscii = 24) Or (intKeyAscii = 26) Then
                GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
                Exit Function
            End If
    End Select
'2016/11/14 LF1667_TIK担当メールアドレス欄の桁拡張 END >>>

    '許可ﾊﾟﾀｰﾝ ﾁｪｯｸ
    Select Case intPatan
        Case 0          '' Check Non
        
        Case 1          '' 0-9
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                intKeyAscii = 0
            End If
            
        Case 2          '' 0-9 or .(46)
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                If intKeyAscii <> 46 Then
                   intKeyAscii = 0
                End If
            End If

        Case 3          '' 0-9 or -(45) or .(46)
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                If intKeyAscii <> 45 And intKeyAscii <> 46 Then
                   intKeyAscii = 0
                End If
            End If
            
        Case 4          '' 0-9 or A-Z
            Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''小文字⇒大文字に変換
            End Select
                    
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
               ((intKeyAscii < 65) Or (intKeyAscii > 90)) Then
               intKeyAscii = 0
            End If
                
        Case 5          '' ここは Non Check
            
        Case 6          '' 0-9 or -(45)
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And (intKeyAscii <> 45) Then
                   intKeyAscii = 0
            End If
            
        Case 7          '' 0-9 or A-Z or a-z
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
               ((intKeyAscii < 65) Or (intKeyAscii > 90)) And _
               ((intKeyAscii < 97) Or (intKeyAscii > 122)) Then
               intKeyAscii = 0
            End If
            
        Case 8          ''  "!" ～ "}"
            If intKeyAscii < 33 Or intKeyAscii > 125 Then
               intKeyAscii = 0
            End If
            
        Case 9          '' 0-9 or A-Z or a-z or +(43) or -(45) or *(42)
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
                    ((intKeyAscii < 65) Or (intKeyAscii > 90)) And _
                    ((intKeyAscii < 97) Or (intKeyAscii > 122)) Then
                If intKeyAscii = 43 Or intKeyAscii = 45 Or intKeyAscii = 42 Then
                
                Else
                    intKeyAscii = 0
                End If
            End If
        Case 10          '' A-Z
            Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''小文字⇒大文字に変換
            End Select
                    
            If intKeyAscii < 65 Or intKeyAscii > 90 Then
               intKeyAscii = 0
            End If
        Case 11          '' 0-9 or -(45) or /(47)
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                If intKeyAscii <> 45 And intKeyAscii <> 47 Then
                   intKeyAscii = 0
                End If
            End If
        
        Case 12          '' 0-9 or -(45) or ((40) or )(41)
            If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                If intKeyAscii <> 45 And intKeyAscii <> 40 And intKeyAscii <> 41 Then
                   intKeyAscii = 0
                End If
            End If
        Case 13
           Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''小文字⇒大文字に変換
            End Select
            
        Case 14          '' 0-9 or A-Z or "-"
            
            'ﾊｲﾌｫﾝ入力許可時
            If (intKeyAscii = 45) Then
                GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
                Exit Function
            End If
            
            Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''小文字⇒大文字に変換
            End Select
                    
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
               ((intKeyAscii < 65) Or (intKeyAscii > 90)) Then
               intKeyAscii = 0
            End If
            
        Case 15          '' 0-9 or A-Z or " "
            
            'ﾌﾞﾗﾝｸ入力許可時
            If (intKeyAscii = 32) Then
                GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
                Exit Function
            End If
            
            Select Case Chr(intKeyAscii)
                Case "a" To "z"
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii))) ''小文字⇒大文字に変換
            End Select
                    
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) And _
               ((intKeyAscii < 65) Or (intKeyAscii > 90)) Then
               intKeyAscii = 0
            End If
            
        Case 16          '' 0-9 or A-Z or " "
            
            'ﾌﾞﾗﾝｸ入力許可時
            If (intKeyAscii = 32) Then
                GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
                Exit Function
            End If
            
            If ((intKeyAscii < 48) Or (intKeyAscii > 57)) Then
               intKeyAscii = 0
            End If
            
        Case 17
            
            'ASCIIｺｰﾄﾞ(0～127)の制御文字(一部除く)以外全て許可
            ''許可制御文字 (ﾘﾀ-ﾝ、ﾗｲﾝﾌｨ-ﾙﾄﾞ)
            If (bolEnterChkFlg = False) And (intKeyAscii = 13 Or intKeyAscii = 10) Then
                GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
                Exit Function
            End If
            If (intKeyAscii < 32) Or (intKeyAscii > 126) Then
                intKeyAscii = 0
            End If

'2009/03/16 Added by N.Kigaku Start --------------------------------------------------------
        Case 18
            'ASCIIｺｰﾄﾞ(0～127)+拡張ASCIIｺｰﾄﾞ(128～255)の制御文字(一部除く)以外全て許可
            ''許可制御文字 (ﾘﾀ-ﾝ、ﾗｲﾝﾌｨ-ﾙﾄﾞ)
            If (bolEnterChkFlg = False) And (intKeyAscii = 13 Or intKeyAscii = 10) Then
                GF_Com_KeyPress = 1   ''ｷｰｺｰﾄﾞは有効
                Exit Function
            End If
            If ((intKeyAscii < 32) Or (intKeyAscii > 126)) And _
               ((intKeyAscii < 160) Or (intKeyAscii > 223)) Then
                intKeyAscii = 0
            End If
'2009/03/16 End ----------------------------------------------------------------------------

    End Select

    If intKeyAscii = 0 Then
        GF_Com_KeyPress = 0    'ｷｰｺｰﾄﾞは無効
    Else
        GF_Com_KeyPress = 1    'ｷｰｺｰﾄﾞは有効
    End If
    
End Function
Public Function GF_Com_CutNumber(strIn_txt As String) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 半角数字の切り出し
' 機能   : 文字列中の半角数字のみを切り出す
' 引数   : strIn_txt As String  '' 文字列
' 戻り値 : String          切り出された半角数字
' 備考   :
'------------------------------------------------------------------------------
    Dim intLoop_c As Integer    'ﾙｰﾌﾟｶｳﾝﾀ
    Dim strChk_chr As String    '文字ﾁｪｯｸ用
    Dim strOut_chr As String    '文字列作成用
    '1文字目から順にﾁｪｯｸし、半角数字の場合は切り出す
    For intLoop_c = 1 To Len(strIn_txt)
        strChk_chr = Mid(strIn_txt, intLoop_c, 1)
        Select Case strChk_chr
            Case "0" To "9"
                strOut_chr = strOut_chr & strChk_chr
        End Select
    Next intLoop_c
    GF_Com_CutNumber = strOut_chr
End Function

Public Function GF_Com_CutAlfNum(strIn_txt As String) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 半角英数字の切り出し
' 機能   : 文字列中の半角英数字のみを切り出す
' 引数   : strIn_txt As String  '' 文字列
' 戻り値 : String          切り出された半角英数字
' 備考   :
'------------------------------------------------------------------------------
    Dim intLoop_c As Integer    'ﾙｰﾌﾟｶｳﾝﾀ
    Dim strChk_chr As String    '文字ﾁｪｯｸ用
    Dim strOut_chr As String    '文字列作成用
    '1文字目から順にﾁｪｯｸし、半角英数字の場合は切り出す
    For intLoop_c = 1 To Len(strIn_txt)
        strChk_chr = Mid(strIn_txt, intLoop_c, 1)
        Select Case strChk_chr
            Case "0" To "9", "A" To "Z", "a" To "z"
                strOut_chr = strOut_chr & strChk_chr
        End Select
    Next intLoop_c
    GF_Com_CutAlfNum = strOut_chr
End Function

Public Sub GS_TextSelect(txtControl As TextBox)
'------------------------------------------------------------------------------
' @(f)
' 機能名 : テキスト全選択
' 機能   : ﾃｷｽﾄBOXのテキストを全選択する
' 引数   : txtControl As TextBox   'ﾃｷｽﾄBOX
' 備考   :
'------------------------------------------------------------------------------
    
    ''ﾌｫｰｶｽを取得したら全選択状態にする
    txtControl.SelStart = 0
    txtControl.SelLength = Len(txtControl)
    
End Sub

Public Function GF_CheckHostStr(strTemp As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : ﾎｽﾄ受信可能文字変換
' 機能   : ﾎｽﾄにて登録不能文字を登録可能文字に変換する。
' 引数   : strTemp   As String      ''対象文字列
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    Dim nCount    As Integer
    Dim intLength As Integer
    
    GF_CheckHostStr = True
    
    intLength = Len(strTemp)
    
    For nCount = 1 To intLength

        Select Case Mid(strTemp, nCount, 1)
        Case "ｱ" To "ﾝ"
        Case "ﾞ"
        Case "ﾟ"
        Case "ｧ"
            strTemp = Left(strTemp, nCount - 1) & "ｱ" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "ｨ"
            strTemp = Left(strTemp, nCount - 1) & "ｲ" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "ｩ"
            strTemp = Left(strTemp, nCount - 1) & "ｳ" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "ｪ"
            strTemp = Left(strTemp, nCount - 1) & "ｴ" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "ｫ"
            strTemp = Left(strTemp, nCount - 1) & "ｵ" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "ｯ"
            strTemp = Left(strTemp, nCount - 1) & "ﾂ" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "ｬ"
            strTemp = Left(strTemp, nCount - 1) & "ﾔ" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "ｭ"
            strTemp = Left(strTemp, nCount - 1) & "ﾕ" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "ｮ"
            strTemp = Left(strTemp, nCount - 1) & "ﾖ" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "ｰ"
            strTemp = Left(strTemp, nCount - 1) & "-" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case Chr(34)    'ﾀﾞﾌﾞﾙｸｵｰﾃｰｼｮﾝ
            strTemp = Left(strTemp, nCount - 1) & Chr(39) & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        Case "\"
        Case "･"
        Case Chr(32) To Chr(90)
        Case Else
            strTemp = Left(strTemp, nCount - 1) & "?" & Mid(strTemp, nCount + 1, intLength - nCount)
            GF_CheckHostStr = False
        End Select
     
    Next nCount
    
End Function

Public Function GF_ConvertWide(strChar As String) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 文字列の全角変換
' 機能   : 文字列中の半角文字を全角に変換する
' 引数   : strChar As String   '文字列
' 戻り値 : String (変換後の文字列)
' 備考   :
'------------------------------------------------------------------------------
    GF_ConvertWide = strConv(strChar, vbWide)
    
End Function

Public Function GF_ConvertNarrow(strChar As String) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 文字列の半角変換
' 機能   : 文字列中の全角文字を半角に変換する
' 引数   : strChar As String   '文字列
' 戻り値 : String (変換後の文字列)
' 備考   :
'------------------------------------------------------------------------------
    GF_ConvertNarrow = strConv(strChar, vbNarrow)
    
End Function

Public Function GF_CutWide(strIn_txt As String) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 全角文字の切り出し
' 機能   : 文字列中の全角文字のみを切り出す
' 引数   : strIn_txt As String  '' 文字列
' 戻り値 : String          切り出された全角文字列
' 備考   :
'------------------------------------------------------------------------------
    Dim intLoop_c As Integer    'ﾙｰﾌﾟｶｳﾝﾀ
    Dim strChk_chr As String    '文字ﾁｪｯｸ用
    Dim strOut_chr As String    '文字列作成用
    strOut_chr = ""
    '1文字目から順にﾁｪｯｸし、半角英数字の場合は切り出す
    For intLoop_c = 1 To Len(strIn_txt)
        strChk_chr = Mid(strIn_txt, intLoop_c, 1)
        If LenB(strConv(strChk_chr, vbFromUnicode)) = 2 Then
            strOut_chr = strOut_chr & strChk_chr
        End If
    Next intLoop_c
    GF_CutWide = strOut_chr
End Function

Public Function GF_CutHalf(strIn_txt As String) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 半角ｶﾅ英数字の切り出し
' 機能   : 文字列中の半角ｶﾅ英数字のみを切り出す
' 引数   : strIn_txt As String  '' 文字列
' 戻り値 : String          切り出された半角ｶﾅ英数字
' 備考   :
'------------------------------------------------------------------------------
    Dim intLoop_c As Integer    'ﾙｰﾌﾟｶｳﾝﾀ
    Dim strChk_chr As String    '文字ﾁｪｯｸ用
    Dim strOut_chr As String    '文字列作成用
    Dim intAsc     As Integer
    strOut_chr = ""
    '1文字目から順にﾁｪｯｸし、半角英数字の場合は切り出す
    For intLoop_c = 1 To Len(strIn_txt)
        strChk_chr = Mid(strIn_txt, intLoop_c, 1)
        intAsc = Asc(strChk_chr)
        If ((intAsc >= 48) And (intAsc <= 57)) Or _
           ((intAsc >= 65) And (intAsc <= 90)) Or _
           ((intAsc >= 97) And (intAsc <= 122)) Or _
           ((intAsc >= 166) And (intAsc <= 223)) Then
            strOut_chr = strOut_chr & strChk_chr
        End If
    Next intLoop_c
    GF_CutHalf = strOut_chr
End Function

Public Sub GS_DecimalPointCheck(intKeyAscii As Integer, strString As String)
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 小数点入力制御
' 機能   : 小数'点'の入力制御
' 引数   : intKeyAscii As Integer   ''ｱｽｷｰｺｰﾄﾞ
'          strString   As String    ''編集中の文字列
' 備考   :
'------------------------------------------------------------------------------
    
    If intKeyAscii = Asc(".") Then
        ''1文字目が"."か2つめの小数点ならはじく
        If strString = "" Or InStr(strString, ".") > 0 Then intKeyAscii = 0
    End If
    
End Sub

Public Function GF_MinusCheck(strString As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : マイナスのみの入力チェック
' 機能   : マイナスのみの入力チェック
' 引数   :  strString   As String    ''文字列
' 戻り値 : True = マイナス以外 / False = マイナスのみ
' 備考   :
'------------------------------------------------------------------------------
    GF_MinusCheck = False
    
    If Trim(strString) <> "-" Then
        GF_MinusCheck = True
    End If
    
End Function

Public Function GF_StrCheck(strString As String, strConvert As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 半角文字ﾁｪｯｸ
' 機能   : 文字列を半角文字変換し、全角文字が残っていればｴﾗｰを返す
' 引数   :
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    
    Dim intStrLength    As Integer  '文字列の長さ(文字数)
    Dim intCount        As Integer  'ﾙｰﾌﾟ変数
    
    On Error GoTo Err_GF_StrCheck
    
    GF_StrCheck = False
    
    strConvert = ""
    
    ''半角に変換できる文字は変換する
    strConvert = strConv(strString, vbNarrow)
    '文字数取得
    intStrLength = Len(strConvert)
    '一文字ずつﾁｪｯｸする
    For intCount = 1 To intStrLength
        '1ﾊﾞｲﾄ以外があればｴﾗｰとして返す
        If LenB(strConv(Mid(strConvert, intCount, 1), vbFromUnicode)) _
                                                    <> 1 Then Exit Function
    Next
        
    GF_StrCheck = True
    
    Exit Function
    
Err_GF_StrCheck:
    
    Call GS_ErrorHandler("GF_StrCheck", "")
    
End Function

Public Function GF_HalfOrFullSizeCheck(ByVal strString As String, ByVal intCheck As Integer) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 半角全角文字チェック
' 機能   : 文字列が半角(または全角)かをチェックする
' 引数   : strString As String  文字列
'          intCheck As Integer  フラグ    1：半角チェック、2：全角
' 戻り値 : True = 成功 / False = 失敗
' 備考   :
'------------------------------------------------------------------------------
    
    Dim intStrLength    As Integer  '文字列の長さ(文字数)
    Dim intCount        As Integer  'ﾙｰﾌﾟ変数
    
    On Error GoTo Err_GF_HalfOrFullSizeCheck
    
    GF_HalfOrFullSizeCheck = False
    
    '文字数取得
    intStrLength = Len(strString)
    '一文字ずつﾁｪｯｸする
    For intCount = 1 To intStrLength
        '1ﾊﾞｲﾄ以外があればｴﾗｰとして返す
        If LenB(strConv(Mid(strString, intCount, 1), vbFromUnicode)) _
                                                    <> intCheck Then Exit Function
    Next
        
    GF_HalfOrFullSizeCheck = True
    
    Exit Function
    
Err_GF_HalfOrFullSizeCheck:
    
    Call GS_ErrorHandler("GF_HalfOrFullSizeCheck", "")
    
End Function

Public Function GF_Com_KeyPressNum(KeyAscii As Integer, crtControl As Control, strTxtData As String, intPatan As Integer, intMaxlen As Integer, Optional dblMaxVal As Double = 9999999999999#) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 数値（金額・利率）のチェック関数
' 機能   : 有効範囲内であれば strTxtData に設定値を保存する
' 引数   : KeyAscii As Integer  13(ﾘﾀｰﾝｺｰﾄﾞ） を指定すると次のフィールドにフォーカスが移動する
'                               0 を指定するとフォーカスの移動なしでチェックが行われる
'                               13(ﾘﾀｰﾝｺｰﾄﾞ）及び 0 の場合、エラー発生時は入力域にフォーカスを合わせる
'          crtControl As Control  ｺﾝﾄﾛｰﾙ
'          strTxtData  As String   ﾃｷｽﾄ
'          intPatan As Integer
'               0 - 金額 (#,##9)        桁数は intMaxlen に依存する  有効範囲の下限値 『0 以上』　上限値は dblMaxVal迄
'               1 - 金額 (#,##9)        桁数は intMaxlen に依存する  有効範囲の下限値 『1 以上』　上限値は dblMaxVal迄
'               2 - 利率 (#9.99999999)  桁数は ２桁．８桁の固定      有効範囲は 『0 より大 100 未満』の固定
'               3 - 金額 (#,##9.99)     桁数は intMaxlen に依存する  有効範囲の下限値 『0    以上』　上限値は dblMaxVal迄
'               4 - 金額 (#,##9.99)     桁数は intMaxlen に依存する  有効範囲の下限値 『0.01 以上』　上限値は dblMaxVal迄
'               5 - 利率 (#9.99999)     桁数は ２桁．５桁の固定      有効範囲は 『0 より大 100 未満』の固定
'               6 - 利率 (#9.999)       桁数は ２桁．３桁の固定      有効範囲は 『0 以上　 100 未満』の固定
'               7 - 利率 (#9.99999999)  桁数は ２桁．８桁の固定      有効範囲は 『0 以上　 100 未満』の固定
'          intMaxlen As Integer 数値の入力可能桁数（小数点含まず）
'          dblMaxVal As Double  有効範囲の上限値　※intPatan が 2 , 5 , 6 の場合を除く
' 戻り値 : エラーメッセージ
' 備考   :
'------------------------------------------------------------------------------
    Dim strMsg       As String
    Dim strCvtTxt    As String

    strMsg = ""
    If (KeyAscii = 13) Or (KeyAscii = 0) Then           'ﾘﾀｰﾝｺｰﾄﾞか確認用ｺｰﾄﾞならﾁｪｯｸ
        If IsNumeric(crtControl.Text) = True Then
            If (intPatan = 0) Or (intPatan = 1) Then
                strCvtTxt = Format(crtControl.Text, "#,##0")
                crtControl.Text = strCvtTxt
            ElseIf (intPatan = 2) Or (intPatan = 7) Then
                strCvtTxt = GF_Com_DecCut((crtControl.Text), 8)
                strCvtTxt = Format(strCvtTxt, "#0.00000000")
            ElseIf (intPatan = 3) Or (intPatan = 4) Then
                strCvtTxt = Format(crtControl.Text, "#,##0.00")
            ElseIf intPatan = 6 Then
                strCvtTxt = GF_Com_DecCut((crtControl.Text), 3)
                strCvtTxt = Format(strCvtTxt, "#0.000")
            ElseIf intPatan = 5 Then
                strCvtTxt = GF_Com_DecCut((crtControl.Text), 5)
                strCvtTxt = Format(strCvtTxt, "#0.00000")
            End If
            
            If ((intPatan = 0) And (CDbl(strCvtTxt) < 0)) Or ((intPatan = 1) And (CDbl(strCvtTxt) < 1)) Or _
               ((intPatan = 3) And (CDbl(strCvtTxt) < 0)) Or ((intPatan = 4) And (CDbl(strCvtTxt) < 0.01)) Then
                strMsg = GF_GetMsg("WTG020")
                Call GS_Com_TxtGotFocus(crtControl)
            ElseIf ((intPatan = 0) Or (intPatan = 1) Or (intPatan = 3) Or (intPatan = 4)) And (CDbl(strCvtTxt) > dblMaxVal) Then
                strMsg = GF_GetMsg("WTG019")
                Call GS_Com_TxtGotFocus(crtControl)
            ElseIf (intPatan = 7) And ((CDbl(strCvtTxt) < 0) Or (CDbl(strCvtTxt) >= 100)) Then
                strMsg = GF_GetMsg("WTG021")
                Call GS_Com_TxtGotFocus(crtControl)
            ElseIf ((intPatan = 2) Or (intPatan = 6) Or (intPatan = 5)) And ((CDbl(strCvtTxt) <= 0) Or (CDbl(strCvtTxt) >= 100)) Then
                strMsg = GF_GetMsg("WTG021")
                Call GS_Com_TxtGotFocus(crtControl)
            Else
                crtControl.Text = strCvtTxt
                strTxtData = GF_Com_CutNumber(strCvtTxt)
                If (KeyAscii = 13) Then            'ﾘﾀｰﾝｺｰﾄﾞならﾁｪｯｸ
                    If (Len(GF_Com_CutNumber(crtControl.Text)) - Len(GF_Com_CutNumber(crtControl.SelText))) <= intMaxlen Then
                        Call GF_Com_KeyPress(0, KeyAscii)
                        Call GS_Com_NextCntl(crtControl)
                    ElseIf Not (KeyAscii = 9 Or KeyAscii = 8) Then
                        KeyAscii = 0
                    End If
                End If
            End If
        Else
            If (intPatan = 0) Or (intPatan = 1) Or (intPatan = 3) Or (intPatan = 4) Then
                strMsg = GF_GetMsg("ITH004") & GF_GetMsg("ITH001")
            ElseIf (intPatan = 2) Or (intPatan = 7) Or (intPatan = 6) Or (intPatan = 5) Then
                strMsg = GF_GetMsg("ITH003") & GF_GetMsg("ITH001")
            End If
            If crtControl.Enabled = True Then
                Call GS_Com_TxtGotFocus(crtControl)
            Else
'                MsgBox "処理に問題はありませんが調査依頼をして下さい" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "モジュール名 : GF_Com_KeypressNum" & Chr(13) & Chr(10) & "プログラム識別 : " & G_APL_Job1 & G_APL_Job2 & Chr(13) & Chr(10) & "コントロール : " & crtControl.Name & Chr(13) & Chr(10) & "補足 : " & strMsg & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "OKボタンをクリックし処理を続行して下さい"
            End If
        End If
    Else
        If (Len(GF_Com_CutNumber(crtControl.Text)) - Len(GF_Com_CutNumber(crtControl.SelText))) <= intMaxlen Then
            Select Case intPatan
            '数値のみ
            Case 0, 1
                Call GF_Com_KeyPress(1, KeyAscii)
            '数値・小数点
            Case 2, 7, 3, 4, 6, 5
                Call GF_Com_KeyPress(2, KeyAscii)
                Call GS_DecimalPointCheck(KeyAscii, crtControl.Text)
            '数値・小数点・マイナス
            Case Else
                Call GF_Com_KeyPress(3, KeyAscii)
                Call GS_DecimalPointCheck(KeyAscii, crtControl.Text)
            End Select
        ElseIf Not (KeyAscii = 9 Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End If
    
    GF_Com_KeyPressNum = strMsg

End Function

Public Function GF_Com_DecCut(strTxt As String, intCut As Integer) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 小数部を切り捨て
' 機能   : 指定された以降の小数部を切り捨てる
' 引数   : strTxt As String     入力値
'          intCut As Integer    小数第何位を指定(そこまで有効）
' 戻り値 : String   算出した値
' 備考   :
'------------------------------------------------------------------------------
    Dim intCheckChar As Integer
    Dim strOutChar As String
    
    strTxt = GF_Com_CnvSisu(strTxt)
    
    intCheckChar = InStr(1, strTxt, ".")
    
    If intCheckChar <> 0 Then
        strOutChar = Mid(strTxt, 1, (intCheckChar + intCut))
    Else
        strOutChar = strTxt
    End If
    
    GF_Com_DecCut = strOutChar
End Function

Public Function GF_Com_CnvSisu(strTxt As String) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 指数表現の数値文字列変換
' 機能   : 指定された指数表現の数値文字列（小数点含む）を指数表現を使わない文字列に変換する
' 引数   : strTxt As String     入力値
' 戻り値 : String   数値文字列
' 備考   :
'------------------------------------------------------------------------------
    Dim intEPoint    As Integer  '指数表現 'E'の位置
    Dim intTenPoint  As Integer  '小数点位置
    Dim intAddZerosu As Integer
    Dim strFugoChar  As String   '指数部の符号
    Dim strFugoTxt   As String   '文字列先頭の符号（-）
    Dim strHenTxt    As String
    Dim strZeroTxt   As String
    Dim intSisu      As Integer  '指数部格納
    Dim strTenMaeTxt As String   '小数点より前の数字
    Dim strTenAtoTxt As String   '小数点より後の数字
    Dim intKugirisu  As Integer
    
    '"E"の位置算出
    intEPoint = InStr(1, strTxt, "E")
    '小数点"."の位置算出
    intTenPoint = InStr(1, strTxt, ".")
    
    '指数部表現を含んでいないものまたは小数点を含まないものは、入力値を戻り値にして抜ける
    If (intEPoint = 0) Or (intTenPoint = 0) Then
        GF_Com_CnvSisu = strTxt
        Exit Function
    End If
    '指数部格納
    intSisu = CInt(Mid(strTxt, intEPoint + 2, 2))
    
    '指数部以前の文字列及び、後の文字列を抜き出す
    strHenTxt = Mid(strTxt, 1, intEPoint - 1)
    strTenMaeTxt = Mid(strHenTxt, 1, intTenPoint - 1)
    '指数部以前の文字列がマイナスを含んでいるか
    If (InStr(1, strTenMaeTxt, "-")) <> 0 Then
        strTenMaeTxt = Mid(strTenMaeTxt, 2)
        strFugoTxt = "-"
    Else
        strFugoTxt = ""
    End If
    strTenAtoTxt = Mid(strHenTxt, intTenPoint + 1)
    
    '指数部以降の符号によって処理振り分け
    strFugoChar = Mid(strTxt, intEPoint + 1, 1)
    'マイナスの場合
    If strFugoChar = "-" Then
        '小数点をずらすだけの場合
        If intSisu < Len(strTenMaeTxt) Then
            intKugirisu = Len(strTenMaeTxt) - intSisu
            strHenTxt = Mid(strTenMaeTxt, 1, intKugirisu) & "." & Mid(strTenMaeTxt, intKugirisu + 1)
            strHenTxt = strHenTxt & strTenAtoTxt
        Else
        '0.0‥をつける場合
            intAddZerosu = intSisu - 1
            strZeroTxt = "0."
            Do While intAddZerosu > 0
                strZeroTxt = strZeroTxt & "0"
                intAddZerosu = intAddZerosu - 1
            Loop
            strHenTxt = strZeroTxt & strTenMaeTxt & strTenAtoTxt
        End If
    'プラスの場合
    Else
        '小数点をずらすだけの場合
        If intSisu < Len(strTenAtoTxt) Then
            strHenTxt = Mid(strTenAtoTxt, 1, intSisu) & "." & Mid(strTenAtoTxt, intSisu + 1)
            strHenTxt = strTenMaeTxt & strHenTxt
        Else
        '00‥をつける場合
            intAddZerosu = intSisu - Len(strTenAtoTxt)
            strZeroTxt = ""
            Do While intAddZerosu > 0
                strZeroTxt = strZeroTxt & "0"
                intAddZerosu = intAddZerosu - 1
            Loop
            strHenTxt = strTenMaeTxt & strTenAtoTxt & strZeroTxt
        End If
    End If
    
    GF_Com_CnvSisu = strFugoTxt & strHenTxt
End Function

Public Sub GS_Com_TxtGotFocus(crtControl As Control)
'------------------------------------------------------------------------------
' @(f)
' 機能名 : テキストコントロールの入力内容を選択状態（反転表示）にする
' 機能   :
' 引数   : crtControl As Control  対象となるテキスト
'                                 コンボボックス
'                                 チェックボックス
' 備考   :
'------------------------------------------------------------------------------
    Screen.ActiveForm.Enabled = True
    If (crtControl.Visible = True) And (crtControl.Enabled = True) Then
        Select Case LCase(Left(crtControl.Name, 3))
            Case "txt"
                crtControl.SetFocus
                crtControl.SelStart = 0
                crtControl.SelLength = Len(crtControl.Text)
            Case "cbo"
                crtControl.SetFocus
            Case "chk"
                crtControl.SetFocus
        End Select
    End If
End Sub

Public Function GF_Com_KeyPressDate(KeyAscii As Integer, cntControl As Control, strTxtDat As String, _
                                            intInCheck As Integer, Optional cntControl2 As Control = Nothing, Optional bolDispFlg As Integer = 0) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 日付のチェック関数
' 機能   : 該当すれば strTxtDat に設定値を保存する
' 引数   : KeyAscii As Integer  13(ﾘﾀｰﾝｺｰﾄﾞ） を指定すると次のフィールドにフォーカスが移動する
'                               0 を指定するとフォーカスの移動なしでチェックが行われる
'                               13(ﾘﾀｰﾝｺｰﾄﾞ）及び 0 の場合、エラー発生時は入力域にフォーカスを合わせる
'          crtControl As Control  ｺﾝﾄﾛｰﾙ
'          strTxtData  As String  ﾃｷｽﾄ
'          intInCheck As Integer　 1 - 営業日ﾁｪｯｸ
'                                  2 - 許容起算日ﾁｪｯｸ
'                                  4 - 本日営業日を越えているﾁｪｯｸ
'                                  8 - 本日営業日以降ﾁｪｯｸ
'                                 16 - 未来日付（本日有効）ﾁｪｯｸ
'                                 32 - 未来日付（本日無効）
'                                 64 - 前日営業日
'          crtControl2 As Control  ｺﾝﾄﾛｰﾙ   曜日表示用ｺﾝﾄﾛｰﾙ(非営業日は赤くする)
'          bolDispFlg  As Integer  表示ﾌﾗｸﾞ　0:年月日、1:年月、2:月日、3:日
' 戻り値 : String          エラーメッセージ
' 備考   : 指定されたチェック条件(intInCheck)に関係無く営業日チェックし曜日の色を設定する
'------------------------------------------------------------------------------
    Dim strDate     As String
    Dim strMsg      As String
    Dim strYoubiTbl() As Variant      '曜日テーブル
    
    
    strYoubiTbl = Array("(日)", "(月)", "(火)", "(水)", "(木)", "(金)", "(土)")
    If Not (cntControl2 Is Nothing) Then
        cntControl2 = ""
        cntControl2.ForeColor = &H80000008              'デフォルト色表示
    End If
    
    strMsg = ""
    strDate = GS_Com_TxtCvtYmd(strTxtDat)     'YYYY/MM/DD に変換
    If (KeyAscii = 13) Or (KeyAscii = 0) Then           'ﾘﾀｰﾝｺｰﾄﾞか確認用ｺｰﾄﾞならﾁｪｯｸ
        Select Case strDate
            Case ""
                strMsg = GF_GetMsg("ITH002") & GF_GetMsg("ITH001")   '未入力
                Call GS_Com_TxtGotFocus(cntControl)
            Case "Not Date"
                strMsg = GF_GetMsg("WTG022")       '無効な日付
                Call GS_Com_TxtGotFocus(cntControl)
            Case "Less Than"
                strMsg = GF_GetMsg("WTG022")       '８桁未満
                Call GS_Com_TxtGotFocus(cntControl)
            Case Else
                'ﾃﾞｰﾀ表示
                Select Case bolDispFlg
                Case 0
                '年月日
                    cntControl.Text = strDate
                Case 1
                '年月
                    cntControl.Text = Format(strDate, "YYYY/MM")
                Case 2
                '月日
                    cntControl.Text = Format(strDate, "MM/DD")
                Case 3
                '日
                    cntControl.Text = Format(strDate, "DD")
                End Select
                
                If ((intInCheck And 1) = 1) And (GF_Com_CheckOpen(strDate) = False) Then    '営業日チェック
                    strMsg = GF_GetMsg("ITG004")                   '営業日でない
                    Call GS_Com_TxtGotFocus(cntControl)
'                ElseIf ((intInCheck And 2) = 2) And (CDbl(GF_Com_CutNumber(strDate)) < CDbl(G_rs_okkymd)) Then     '許容起算日チェック
'                    strMsg = GF_GetMsg("B249")                   '許容起算日を越えている（下回っている）
'                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 4) = 4) And (CDbl(GF_Com_CutNumber(strDate)) > CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then        '過去日付チェック（本日有効）
                    strMsg = GF_GetMsg("ITG005")                   '本日営業日を越えている
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 8) = 8) And (CDbl(GF_Com_CutNumber(strDate)) >= CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then     '過去日付チェック（本日無効）
                    strMsg = GF_GetMsg("ITG006")                   '本日営業日以降です
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 16) = 16) And (CDbl(GF_Com_CutNumber(strDate)) < CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then     '未来日付チェック（本日有効）
                    strMsg = GF_GetMsg("ITG007")                   '未来日付（本日有効）
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 32) = 32) And (CDbl(GF_Com_CutNumber(strDate)) <= CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then     '未来日付チェック（本日無効）
                    strMsg = GF_GetMsg("ITG008")                   '未来日付（本日無効）
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 64) = 64) And (CDbl(GF_Com_CutNumber(strDate)) <> CDbl(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate))) Then      '前日営業日チェック
                    strMsg = GF_GetMsg("ITG009")                   '前日営業日と同じ
                    Call GS_Com_TxtGotFocus(cntControl)
                ElseIf ((intInCheck And 128) = 128) And CDbl(Left(GF_Com_CutNumber(strDate), 6) < CDbl(Left(GF_Com_CutNumber(Screen.ActiveForm.lblNowDate), 6))) Then  '前月以降ﾁｪｯｸ
                    strMsg = GF_GetMsg("ITG010")                   '前月以降
                    Call GS_Com_TxtGotFocus(cntControl)
                Else
                    strTxtDat = GF_Com_CutNumber(strDate)
                    If (KeyAscii = 13) Then            'ﾘﾀｰﾝｺｰﾄﾞならﾁｪｯｸ
                        If (Len(GF_Com_CutNumber(strTxtDat)) - Len(GF_Com_CutNumber(strTxtDat))) < 8 Then
                            Call GF_Com_KeyPress(11, KeyAscii)
                        ElseIf Not (KeyAscii = 9 Or KeyAscii = 8) Then
                            KeyAscii = 0
                        End If
                        Call GS_Com_NextCntl(cntControl)
                    End If
                End If
                
                If Not (cntControl2 Is Nothing) Then
                    'チェック条件(intInCheck)に関係無く営業日チェックし曜日の色を設定する
                    If (GF_Com_CheckOpen(strDate) = False) Then
                        cntControl2.ForeColor = &HFF&              '赤色表示
                    End If
                    cntControl2 = strYoubiTbl(Weekday(strDate) - 1)    '曜日設定
                End If
        End Select
    ElseIf (Len(GF_Com_CutNumber(strDate)) - Len(GF_Com_CutNumber(strDate))) < 8 Then
        Call GF_Com_KeyPress(11, KeyAscii)
    ElseIf Not (KeyAscii = 9 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
    GF_Com_KeyPressDate = strMsg
End Function

Public Function GS_Com_TxtCvtYmd(strDate As String) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 文字列編集
' 機能   : 入力された文字列を YYYY/MM/DD に編集する
' 引数   : strDate As String     入力日付
' 戻り値 : String   編集後文字列
'                       "”          　入力無し
'                       "Not Date”　　実在しない年月の場合
'                       "Less Than"    8桁未満、西暦１０００年未満の場合
' 備考   :
'------------------------------------------------------------------------------
    Dim strInData As String    ' コントロールテキスト内容
    Dim strFmtData As String   ' 編集後入力データ
    
    GS_Com_TxtCvtYmd = ""
    
    If IsDate(strDate) = True Then
        strFmtData = Format(strDate, "YYYY/MM/DD")
    Else
        ' スラッシュのカット
        strInData = GF_Com_CutNumber(strDate)
        If Len(strInData) = 0 Then
            Exit Function
        ElseIf Len(strInData) <> 8 Then
            GS_Com_TxtCvtYmd = "Less Than"
            Exit Function
        End If
        
        strFmtData = Format(strInData, "####/##/##")
    End If
    
    '存在チェック
    If (IsDate(strFmtData)) = True Then
        '編集後入力データを返す
        GS_Com_TxtCvtYmd = strFmtData
    Else
        GS_Com_TxtCvtYmd = "Not Date"
    End If
   
End Function

Public Function GF_Com_CheckOpen(strInDate As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 営業日ﾁｪｯｸ
' 機能   : 営業日、非営業日を判別する
' 引数   : strDate As String     入力日付
' 戻り値 : Boolean   営業日:TRUE、非営業日：FALSE
' 備考   :
'------------------------------------------------------------------------------
    Dim strDate     As String    ' コントロールテキスト内容
    Dim strFmtData  As String   ' 編集後入力データ
    Dim strDataFlag As String

    GF_Com_CheckOpen = False
    
    If IsDate(strInDate) = True Then
        strFmtData = Format(strInDate, "YYYY/MM/DD")
    Else
        ' スラッシュのカット
        strDate = GF_Com_CutNumber(strInDate)
        If Len(strDate) <> 8 Then Exit Function
        
        strFmtData = Format(strDate, "####/##/##")
    End If
    
    '存在チェック
    If (IsDate(strFmtData)) = True Then
        Select Case Weekday(strFmtData)
            Case vbMonday, vbTuesday, vbWednesday, vbThursday, vbFriday
                GF_Com_CheckOpen = GF_Com_CheckHoriday(strFmtData)
        End Select
    End If
    
End Function

Public Function GF_Com_CheckHoriday(strInDate As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 営業日ﾁｪｯｸ
' 機能   : 営業日、非営業日をｶﾚﾝﾀﾞﾏｽﾀを見て判別する
' 引数   : strDate As String     入力日付
' 戻り値 : Boolean   営業日:TRUE、非営業日：FALSE
' 備考   :
'------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTableName As String
    Dim oraDyna    As OraDynaset
    
    GF_Com_CheckHoriday = True
    
    '祝祭日ファイル抽出
    strSQL = "SELECT * FROM TGCLMR"
    strSQL = strSQL & " WHERE YMD = '" & strInDate & "'"
    strSQL = strSQL & "   AND HRDFLG = '1'"
    Set oraDyna = gOraDataBase.dbcreatedynaset(strSQL, ORADYN_NOCACHE)
    
    'レコードが存在した場合実行
    If oraDyna.RecordCount > 0 Then
        GF_Com_CheckHoriday = False
    End If
    
End Function

Public Function GF_LenString(strChar As String) As Integer
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 文字数の長さをﾊﾞｲﾄ数で返す
' 機能   :
' 引数   : strChar As String    文字列
' 戻り値 : Integer              文字列のﾊﾞｲﾄ数
' 備考   :
'------------------------------------------------------------------------------
    GF_LenString = LenB(strConv(strChar, vbFromUnicode))
End Function

Public Function GF_Left(strChar As String, intCount As Integer) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 文字列の左から指定ﾊﾞｲﾄ数分取得する
' 機能   :
' 引数   : strChar As String    文字列
'          intCount As Integer  ﾊﾞｲﾄ数
' 戻り値 : String      指定ﾊﾞｲﾄ数分の文字列
' 備考   :
'------------------------------------------------------------------------------
    GF_Left = strConv(LeftB(strConv(strChar, vbFromUnicode), intCount), vbUnicode)
End Function

Public Function GF_Mid(strChar As String, intStrat As Integer, intCount As Integer) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 文字列の指定位置から指定ﾊﾞｲﾄ数分取得する
' 機能   :
' 引数   : strChar As String    文字列
'          intStrat As Integer  開始位置(ﾊﾞｲﾄ数)
'          intCount As Integer  ﾊﾞｲﾄ数
' 戻り値 : String      指定ﾊﾞｲﾄ数分の文字列
' 備考   :
'------------------------------------------------------------------------------
    GF_Mid = strConv(MidB(strConv(strChar, vbFromUnicode), intStrat, intCount), vbUnicode)
End Function

Public Function GF_Right(strChar As String, intCount As Integer) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 文字列の右から指定ﾊﾞｲﾄ数分取得する
' 機能   :
' 引数   : strChar As String    文字列
'          intCount As Integer  ﾊﾞｲﾄ数
' 戻り値 : String      指定ﾊﾞｲﾄ数分の文字列
' 備考   :
'------------------------------------------------------------------------------
    GF_Right = strConv(RightB(strConv(strChar, vbFromUnicode), intCount), vbUnicode)
End Function

Public Function GF_CheckContract(bolMoveKbn As Boolean, strYear As String, strMonth As String, strContractKbn As String) As String
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 約定日算出
' 機能   :
' 引数   : bolMoveKbn As Boolean   True：前だし、False：後だし
'          strYear  　As String    年
'          strMonth 　As String    月
'          strContractKbn As String　1～28：日付、38：月末-2、39：月末-1、40：月末
' 戻り値 : String          補正した日付(ｴﾗｰ時は"")
' 備考   :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strDate     As String
    Dim intStep     As Integer
    Dim intRet      As Integer
    
    GF_CheckContract = ""
    
    '前だし、後だしか
    If bolMoveKbn = False Then
        intStep = 1
    Else
        intStep = -1
    End If
    
    '日付取得
    Select Case strContractKbn
    Case "38"
        '月末 - 2
        strDate = Format(DateAdd("D", -3, DateAdd("M", 1, strYear & "/" & strMonth & "/01")), "YYYY/MM/DD")
    Case "39"
        '月末 - 1
        strDate = Format(DateAdd("D", -2, DateAdd("M", 1, strYear & "/" & strMonth & "/01")), "YYYY/MM/DD")
    Case "40"
        '月末
        strDate = Format(DateAdd("D", -1, DateAdd("M", 1, strYear & "/" & strMonth & "/01")), "YYYY/MM/DD")
    Case Else
        strDate = strYear & "/" & strMonth & "/" & strContractKbn
        
        '日付ﾁｪｯｸ
        If IsDate(strDate) = False Then
            intRet = GF_MsgBoxDB("", "WTG022", "OK", "E")
            Exit Function
        End If
        
        strDate = Format(strDate, "YYYY/MM/DD")
    End Select
    
    '土日、祝日ﾁｪｯｸ
    Do
        '営業日になったら抜ける
        If GF_Com_CheckOpen(strDate) = True Then Exit Do
        
        '1日ずらす
        strDate = Format(DateAdd("D", intStep, strDate), "YYYY/MM/DD")
    Loop
    
    GF_CheckContract = strDate
    
    Exit Function
    
ErrHandler:
    ''ｴﾗｰﾊﾝﾄﾞﾗ
    Call GS_ErrorHandler("GF_CheckContract")
    
End Function

Public Function GF_ChangeQuateSing(strString As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :  シングルクォーテーション変換
' 機能      :　シングルクォーテーションをオラクルデータベースに登録できる形式に変換する
' 引数     ： strString   As String (I)    判定する文字列
' 戻り値    :  変換後の文字列
' 備考      :
'------------------------------------------------------------------------------
    GF_ChangeQuateSing = Replace(strString, "'", "''", 1, , vbBinaryCompare)
End Function

'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :  アンパサンド(&)をラベルに表示できる形式にする
' 機能      :
' 引数     ： strChar   As String (I)    文字列
' 戻り値    :  文字列
' 備考      :
'------------------------------------------------------------------------------
Public Function GF_ReplaceAmper(strChar As String) As String
    GF_ReplaceAmper = Replace(strChar, "&", "&&", 1, , vbBinaryCompare)
End Function

'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :  ラベルに表示しているアンパサンド(&)を戻した形式にする
' 機能      :
' 引数     ： strChar   As String (I)    文字列
' 戻り値    :  文字列
' 備考      :
'------------------------------------------------------------------------------
Public Function GF_UndoAmper(strChar As String) As String
    GF_UndoAmper = Replace(strChar, "&&", "&", 1, , vbBinaryCompare)
End Function


Public Function GF_FileNameRestrinction(strName As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :   ﾌｧｲﾙ名制限ﾁｪｯｸ
' 機能      :   登録可能なﾌｧｲﾙ名かﾁｪｯｸする
' 引数      :   strName AS String   ﾌｧｲﾙ名
' 戻り値    : 　True:正常 / False:ﾌｧｲﾙ名制限ｴﾗｰ
' 備考      :
'------------------------------------------------------------------------------
    On Error Resume Next
    
    Dim i         As Integer
    Dim intAsc    As Integer
    
    GF_FileNameRestrinction = False
    
    For i = 1 To Len(strName)
        intAsc = Asc(Mid(strName, i, 1))
        
        '数字、英字(大文字,小文字)、ﾊｲﾌﾝ、ｱﾝﾀﾞｰﾊﾞｰ
        If (47 < intAsc And intAsc < 58) _
          Or (64 < intAsc And intAsc < 91) _
          Or (96 < intAsc And intAsc < 123) _
          Or (intAsc = 45) _
          Or (intAsc = 95) Then
            
        Else
            Exit Function
        End If
    Next i
    
    GF_FileNameRestrinction = True
End Function

Public Function GF_CheckNumeric(ByVal strNum As String, _
                                Optional ByVal blnIntegerFlg As Boolean = True _
                                ) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 数値チェック
' 機能   :
' 引数   : strNum As String             'チェック対象文字列
'          blnIntegerFlg As Boolean     '整数チェックフラグ
'                                         True:整数チェックあり ,False:なし
' 戻り値 : True:数値 / False:数値以外
' 備考   :
'------------------------------------------------------------------------------
    GF_CheckNumeric = False
    
    If blnIntegerFlg = True Then
        '整数チェックで小数点があるときはエラー
        If InStr(1, strNum, ".", vbTextCompare) > 0 Then Exit Function
    End If
    
    If IsNumeric(strNum) = True Then
        GF_CheckNumeric = True
    End If
End Function

''--------------------------------------------------------------------------------
'' @(f)
'' 機能概要 : 数値チェック(IsNumeric[3E4/3E+4/3E-4/(10)/\1,000/10.5/&12/12/12+/0001/&HFF/１２３(全角)]判定できない)
''
'' 引数     : ByVal strNumber As String         チェックデータ
''          : Optional blnFlg As Boolean        TRUE:マイナス可, FALSE:マイナス不可
''
'' 戻り値   : TRUE：正常 FALSE：異常  Boolean
''--------------------------------------------------------------------------------
Public Function GF_CheckNumber2(ByVal strNumber As String, Optional blnFlg As Boolean = True) As Boolean
    Dim intLen As Integer

    If Left(strNumber, 1) = "-" Then
        If blnFlg Then      ' マイナス可
            strNumber = Mid(strNumber, 2)
        Else                ' マイナス不可
            Exit Function
        End If
    End If

    intLen = Len(strNumber)

    ' 文字列に全角が含まれていたら関数を抜ける or 文字列が空なら抜ける
    If LenB(strConv(strNumber, vbFromUnicode)) <> intLen Or intLen = 0 Then Exit Function

    ' 文字列がすべて数字で構成されてるか?
    If strNumber Like String$(intLen, "#") Then GF_CheckNumber2 = True

End Function

Public Function GF_SearchCount(ByVal strChar As String, ByVal strSearch As String) As Integer
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :  対象の文字の件数を取得する
' 機能      :
' 引数     ： strChar   As String (I)    文字列
'            strSearch As String(I)     カウントする文字列
' 戻り値    :  文字数
' 備考      :
'------------------------------------------------------------------------------
    Dim i           As Integer
    Dim intPoint    As Integer
    Dim intCount    As Integer
    
    intCount = 0
    
    For i = 1 To Len(strChar)
        intPoint = InStr(i, strChar, strSearch, vbTextCompare)
        If intPoint > 0 Then
            intCount = intCount + 1
            i = intPoint
        Else
            Exit For
        End If
    Next i
    
    GF_SearchCount = intCount
    
End Function

Public Function GF_CharPermitChek(intPatan As Integer, ByVal strChar As String) As Integer
'------------------------------------------------------------------------------
' @(f)
' 機能名 : 許可文字列チェック
' 機能   : 許可された文字以外のものが含まれる場合はエラーとする
' 引数   : intPatan As Integer     ''許可ﾊﾟﾀｰﾝ
'                1  - 数字  Code Non Check  "0,1,2,～9"
'                2  - 数字＋ﾋﾟﾘｵﾄﾞ Code Non Check   "0,1,2,～9,.,"
'                3  - 数字＋ﾋﾟﾘｵﾄﾞ＋ﾏｲﾅｽ Code Non Check   "0,1,2,～9,.,-"
'                4  - 数字＋英字 Code Non Check   "0,1,2,～9,A～Z"
'                5  - 数字＋ﾏｲﾅｽ Code Non Check   "0,1,2,～9,-"
'                6  - 数字＋英字 Code Non Check   "0,1,2,～9,A～Z,a～z,"
'                7  - '!' ～ '}' までOK    (ｺｰﾄﾞだと 33 ～ 125まで)
'                8  - 英数字＋ﾌﾟﾗｽ＋ﾏｲﾅｽ＋"*" Code Non Check   "0,1,2,～9,A～Z,a～z,+,-,*"
'                9  - 英大文字 Code Non Check   "A～Z"
'                10 - 数字＋ﾊｲﾌﾝ＋ｽﾗｯｼｭ Code Non Check   "0,1,2,～9,-,/"
'                11 - 数字＋ﾊｲﾌﾝ＋ｶｯｺ Code Non Check   "0,1,2,～9,-,(,)"
'                12 - 数字＋英字＋ﾊｲﾌﾝ Code Non Check   "0,1,2,～9,A～Z","-"
'                13 - 数字＋英字＋ﾌﾞﾗﾝｸ Code Non Check   "0,1,2,～9,A～Z"," "
'                14 - 数字＋ﾌﾞﾗﾝｸ Code Non Check   "0,1,2,～9," "
'                15 - ASCIIｺｰﾄﾞ(32～126)の文字
'                16 - ASCIIｺｰﾄﾞ(32～126)+拡張ASCIIｺｰﾄﾞ(160～223)の文字
'          strChar As Integer       ''文字列
' 戻り値 : Integer          ''0･･･有効な文字列   0以外･･･無効な文字列が最初に見つかった位置
' 備考   :
'------------------------------------------------------------------------------
    Dim i       As Integer
    Dim intAsc  As Integer
    Dim intLen  As Integer
    
    GF_CharPermitChek = 0
    
    intLen = Len(strChar)
    
    If intLen = 0 Then
        Exit Function
    End If
    
    '許可ﾊﾟﾀｰﾝ ﾁｪｯｸ
    Select Case intPatan
        Case 1          ''1-9
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) Then
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
            
        Case 2          '' 0-9 or .(46)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If (intAsc < 48) Or (intAsc > 57) Then
                    If intAsc <> 46 Then
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
            
        Case 3          '' 0-9 or -(45) or .(46)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If (intAsc < 48) Or (intAsc > 57) Then
                    If intAsc <> 45 And intAsc <> 46 Then
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
            
        Case 4          '' 0-9 or A-Z
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   ((intAsc < 65) Or (intAsc > 90)) Then
                   GF_CharPermitChek = i
                   Exit Function
                End If
            Next i
            
        Case 5          '' 0-9 or -(45)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And (intAsc <> 45) Then
                       GF_CharPermitChek = i
                       Exit Function
                End If
            Next i
            
        Case 6          '' 0-9 or A-Z or a-z
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   ((intAsc < 65) Or (intAsc > 90)) And _
                   ((intAsc < 97) Or (intAsc > 122)) Then
                   GF_CharPermitChek = i
                   Exit Function
                End If
            Next i
            
        Case 7          ''  "!" ～ "}"
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If intAsc < 33 Or intAsc > 125 Then
                   GF_CharPermitChek = i
                   Exit Function
                End If
            Next i
            
        Case 8          '' 0-9 or A-Z or a-z or +(43) or -(45) or *(42)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                        ((intAsc < 65) Or (intAsc > 90)) And _
                        ((intAsc < 97) Or (intAsc > 122)) Then
                    If intAsc = 43 Or intAsc = 45 Or intAsc = 42 Then
                    
                    Else
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
                        
        Case 9          ''A-Z
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 65) Or (intAsc > 90)) Then
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
            
        Case 10          '' 0-9 or -(45) or /(47)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If (intAsc < 48) Or (intAsc > 57) Then
                    If intAsc <> 45 And intAsc <> 47 Then
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
            
        Case 11          '' 0-9 or -(45) or ((40) or )(41)
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If (intAsc < 48) Or (intAsc > 57) Then
                    If intAsc <> 45 And intAsc <> 40 And intAsc <> 41 Then
                        GF_CharPermitChek = i
                        Exit Function
                    End If
                End If
            Next i
            
        Case 12          '' 0-9 or A-Z or "-"
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))

'2008/02/21 Updated by N.Kigaku Start-------------------------
'ﾊｲﾌﾝのASCIIｺｰﾄﾞ修正
'                If ((intAsc < 48) Or (intAsc > 57)) And _
'                   ((intAsc < 65) Or (intAsc > 90)) And _
'                   (intAsc <> 32) Then
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   ((intAsc < 65) Or (intAsc > 90)) And _
                   (intAsc <> 45) Then
'2008/02/21 Update End ---------------------------------------
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
            
        Case 13          ''0-9 or A-Z or " "
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   ((intAsc < 65) Or (intAsc > 90)) And _
                   (intAsc <> 32) Then
                    
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
        
        Case 14          ''0-9 or " "
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 48) Or (intAsc > 57)) And _
                   (intAsc <> 32) Then
                    
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i

'2011/05/20 Added by N.Kigaku Start --------------------------------
        Case 15         ''ASCIIｺｰﾄﾞ(32～126)の文字
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))

                If (intAsc < 32) Or (intAsc > 126) Then
                    
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i

        Case 16         ''ASCIIｺｰﾄﾞ(32～126)+拡張ASCIIｺｰﾄﾞ(160～223)の文字
            For i = 1 To intLen
                intAsc = Asc(Mid(strChar, i, 1))
                
                If ((intAsc < 32) Or (intAsc > 126)) And _
                   ((intAsc < 160) Or (intAsc > 223)) Then
                    
                    GF_CharPermitChek = i
                    Exit Function
                End If
            Next i
'2011/05/20 End ----------------------------------------------------

        Case Else
    
    End Select
    
End Function

''--------------------------------------------------------------------------------
'' @(f)
'' 機能概要 : 日付型をチェック
''
'' 引数     : ByRef strDt As String             チェックデータ
''          : ByRef blnEnd As Boolean           True = '99999999'を許可, False = '99999999'は不正とする
''
'' 戻り値   : TRUE：正常 FALSE：異常  Boolean
''--------------------------------------------------------------------------------
Public Function GF_DateConv(strDt As String, Optional blnEnd As Boolean = True) As Boolean
    Dim strConv As String

    GF_DateConv = False

    '' 8桁ではない or 数値ではない
    If Len(strDt) <> 8 Or IsNumeric(strDt) = False Then Exit Function

    '' 全角文字が含まれる場合はエラー
    If GF_LenString(strDt) <> Len(strDt) Then Exit Function

    ''YYYY/MM/DD形式にする
    strConv = Mid(strDt, 1, 4) & "/" & Mid(strDt, 5, 2) & "/" & Mid(strDt, 7, 2)
   
    '' 99999999は使用可?
    If blnEnd = True Then
        If IsDate(strConv) = True Or strDt = "99999999" Then
            GF_DateConv = True
        End If
    ElseIf IsDate(strConv) = True Then
        GF_DateConv = True
    End If

End Function

Public Function GF_THJCMBXMR_CHK(ByVal strCMBNAME As String, ByVal strCDVAL As String) As Boolean
'--------------------------------------------------------------------------------
' @(f)
' 機能名　　: 共通表示ﾘｽﾄﾃｰﾌﾞﾙ検索
' 機能　　　: 共通表示ﾘｽﾄﾃｰﾌﾞﾙを検索し、該当ﾃﾞｰﾀがあるかﾁｪｯｸする
' 引数　　　: strCMBNAME        ''ﾘｽﾄ表示内容
' 　　　　　: strCDVAL          ''ﾘｽﾄ物理値
' 戻り値　　: True：ﾃﾞｰﾀ有り    False：ﾃﾞｰﾀ無し
' 機能説明　:
'--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strSQL   As String
    Dim Dynaset   As OraDynaset
        
    GF_THJCMBXMR_CHK = False
        
    strSQL = ""
    strSQL = strSQL & "SELECT ROWID"
    strSQL = strSQL & "  FROM THJCMBXMR"
    strSQL = strSQL & " WHERE CDVAL ='" & strCDVAL & "'"
    strSQL = strSQL & "   AND CMBNAME ='" & strCMBNAME & "'"
    
    Set Dynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    '共通表示ﾃｰﾌﾞﾙﾃﾞｰﾀ有無ﾁｪｯｸ
    If (Dynaset.EOF = True) Then
    'ﾃﾞｰﾀ無し
        GF_THJCMBXMR_CHK = False
    Else
    'ﾃﾞｰﾀ有り
        GF_THJCMBXMR_CHK = True
    End If
    
    Set Dynaset = Nothing
    
    Exit Function
        
ErrHandler:
    Call GS_ErrorHandler("GF_THJCMBXMR_CHK")

End Function

''--------------------------------------------------------------------------------
'' @(f)
'' 機能概要 : 小数点桁数チェック
''
'' 引数     : ByRef strDt As String             チェックデータ
''          : ByRef intInt As Integer           整数部桁数
''          : ByRef intDec As Integer           小数部桁数
''          : Optional blnFlg As Boolean        TRUE:マイナス可, FALSE:マイナス不可
''
'' 戻り値   : 0:正常, -1:桁が違う, -2:数値ではない
''--------------------------------------------------------------------------------
Public Function GF_ChkDeci(ByVal strDt As String, intInt As Integer, intDec As Integer, Optional blnFlg As Boolean = True) As Integer
    On Error GoTo ErrHandler
    
    Dim intLen As Integer, intCnt As Integer
    Dim strDec() As String

    GF_ChkDeci = 0

    '' マイナス?
    If Left(strDt, 1) = "-" Then
        If blnFlg Then  '' マイナス許可?
            strDt = Mid(strDt, 2)   '' マイナスを除去
        Else
            GF_ChkDeci = -2
            Exit Function
        End If
    End If

    strDec = Split(strDt, ".")
    If UBound(strDec) > 0 Then  '' 小数値?
        '' 桁数チェック
        If Len(strDec(0)) > intInt Then GF_ChkDeci = -1
        If Len(strDec(1)) > intDec Then GF_ChkDeci = -1

        '' 数値型チェック
        If GF_CheckNumber2(strDec(0) & strDec(1)) = False Then GF_ChkDeci = -2
    Else                        '' 整数値
        '' 桁数チェック
        If Len(strDt) > intInt Then GF_ChkDeci = -1

        '' 数値型チェック
        If GF_CheckNumber2(strDec(0)) = False Then GF_ChkDeci = -2
    End If

    Exit Function
        
ErrHandler:
    Call GS_ErrorHandler("GF_ChkDeci")
End Function

Public Function GF_OptFormatChk(strOpt As String, strSize As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名　　:OPTﾌｫｰﾏｯﾄﾁｪｯｸ
' 機能　　　:OPTﾌｫｰﾏｯﾄﾁｪｯｸ
' 引数　　　:strOpt :OPTｺｰﾄﾞ
' 　　　　　:strSize:ｻｲｽﾞ区分(0:ｻｲｽﾞ無,1:ｻｲｽﾞ在)
' 機能説明　:OPTﾌｫｰﾏｯﾄﾁｪｯｸ
'------------------------------------------------------------------------------
    ''変数定義
    Dim lngOptCount     As Long ''ｻｲｽﾞ無OPT桁数
    Dim lngOptSizeCount As Long ''ｻｲｽﾞ有OPT桁数
    Dim intIdx          As Integer ''ｶｳﾝﾀ
    Dim strChar         As String ''OPT1文字

    ''処理状態設定
    GF_OptFormatChk = False
    
    ''変数初期化
    lngOptCount = mlngOptLength
    lngOptSizeCount = mlngSizeOptLength
    intIdx = 1
    
    ''OPT桁数ﾁｪｯｸ
    If GF_LenString(Trim(strOpt)) = lngOptCount And strSize = "0" Then
    ElseIf GF_LenString(Trim(strOpt)) = lngOptSizeCount And strSize = "1" Then
    Else
        Exit Function
    End If
    
    ''ﾌｫｰﾏｯﾄﾁｪｯｸ
    Do Until intIdx > lngOptCount
        strChar = ""
        If intIdx = 2 Or intIdx = 3 Then
        ''2桁目,3桁目(数値)
            strChar = GF_Mid(Trim(strOpt), intIdx, 1)
            If GF_CheckNumeric(strChar) = False Then
                Exit Function
            End If
        Else
        ''1桁目,4桁目(英字)
            strChar = GF_Mid(Trim(strOpt), intIdx, 1)
            If GF_Com_CheckString(10, strChar) = False Then
                Exit Function
            End If
        End If
        intIdx = intIdx + 1
    Loop

    ''処理状態再設定
    GF_OptFormatChk = True

End Function

Public Function GF_CheckLinefeed(ByVal strChar As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' 機能名　　:改行ｺｰﾄﾞ入力ﾁｪｯｸ
' 機能　　　:
' 引数　　　:strChar As String      ﾁｪｯｸする文字列
' 機能説明　:
' 戻り値　　:False:改行あり / True:改行なし
'------------------------------------------------------------------------------
    If 0 < InStr(1, strChar, vbCrLf, vbBinaryCompare) Then
        GF_CheckLinefeed = False
    Else
        GF_CheckLinefeed = True
    End If
End Function

'2006/12/11 Added by N.Kigaku
''--------------------------------------------------------------------------------
'' @(f)
'' 機能概要 : 半角英数字記号チェック
''
'' 引数     : ByVal strDt As String             チェックデータ
''
'' 戻り値   : TRUE：正常 FALSE：異常  Boolean
''--------------------------------------------------------------------------------
Public Function GF_CheckEngNumMark(ByVal strDt As String, Optional blnCRLF_Flag As Boolean = False) As Boolean
    Dim intCnt As Integer
    Dim lngChk_Asc As Long

    '' 半角英数字記号
    For intCnt = 1 To Len(strDt)
        lngChk_Asc = Asc(Mid(strDt, intCnt, 1))
        
        If (blnCRLF_Flag = True) And (lngChk_Asc = 10 Or lngChk_Asc = 13) Then
            '改行ｺｰﾄﾞは対象外

        ElseIf Asc(" ") > lngChk_Asc Or Asc("~") < lngChk_Asc Then
            GF_CheckEngNumMark = False
            Exit Function

        End If
    Next

    GF_CheckEngNumMark = True

End Function

'2008/05/27 Added by N.Kigaku
Public Function GF_ChangeQuateDouble(strString As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :  ダブルクォーテーション変換
' 機能      :　ダブルクォーテーションを出力できる形式に変換する
' 引数     ： strString   As String (I)    判定する文字列
' 戻り値    :  変換後の文字列
' 備考      :　CSVﾌｫｰﾏｯﾄ出力時の変換など
'------------------------------------------------------------------------------
    GF_ChangeQuateDouble = Replace(strString, """", """""", 1, , vbBinaryCompare)
End Function


'Public Function GF_Com_KeypressCif(KeyAscii As Integer, crtControl As Control, strTxtData As String, crtOutControl As Control) As String
''------------------------------------------------------------------------------
'' @(f)
'' 機能名 : 取引先番号のチェック関数
'' 機能   : 該当すれば strTxtData に設定値を保存する。また、該当した顧客名称を crtOutControl に表示する
'' 引数   : KeyAscii As Integer      13(ﾘﾀｰﾝｺｰﾄﾞ） を指定すると次のフィールドにフォーカスが移動する
''                                   0 を指定するとフォーカスの移動なしでチェックが行われる
''                                   13(ﾘﾀｰﾝｺｰﾄﾞ）及び 0 の場合、エラー発生時は入力域にフォーカスを合わせる
''          crtControl As Control    取引先番号入力コントロール
''          strTxtData As String     入出力値
''          crtOutControl As Control 名称表示用コントロール
'' 戻り値 : String   エラーメッセージ
'' 備考   :
''------------------------------------------------------------------------------
'    Dim strMsg      As String
'    Dim strCifName  As String
'    Const intMaxlen As Integer = 5
'
'    strMsg = ""
'    If (KeyAscii = 13) Or (KeyAscii = 0) Then           'ﾘﾀｰﾝｺｰﾄﾞか確認用ｺｰﾄﾞならﾁｪｯｸ
'        If crtControl.Text <> "" Then
'            crtControl.Text = Format(crtControl.Text, "00000")
'            strCifName = GF_Com_Cifget(crtControl)
'
'            If Len(strCifName) = 0 Then                   '名称が取得できたか
'                strMsg = GF_GetMsg("WTH004")
'                Call GS_Com_TxtGotFocus(crtControl)
'            Else
'                crtOutControl.Caption = strCifName
'                strTxtData = crtControl.Text
'                If (KeyAscii = 13) Then            'ﾘﾀｰﾝｺｰﾄﾞならﾁｪｯｸ
'                    If (Len(crtControl.Text) - Len(crtControl.SelText)) < intMaxlen Then
'                        Call GF_Com_Keypress(1, KeyAscii)
'                    Else
'                        KeyAscii = 0
'                    End If
'                End If
'            End If
'        Else
'            If crtControl.Enabled = True Then
'                Call GS_Com_TxtGotFocus(crtControl)
'            Else
''                MsgBox "処理に問題はありませんが調査依頼をして下さい" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "モジュール名 : GF_Com_KeypressCif" & Chr(13) & Chr(10) & "プログラム識別 : " & G_APL_Job1 & G_APL_Job2 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "OKボタンをクリックし処理を続行して下さい"
'            End If
'            strMsg = GF_GetMsg("ITH005") & GF_GetMsg("ITH001")
'        End If
'    Else
'        If (Len(crtControl.Text) - Len(crtControl.SelText)) < intMaxlen Then
'            Call GF_Com_Keypress(1, KeyAscii)
'        Else
'            KeyAscii = 0
'        End If
'    End If
'
'    GF_Com_KeypressCif = strMsg
'
'End Function
'
'Public Function GF_Com_Cifget(vntControl As Control) As String
''------------------------------------------------------------------------------
'' @(f)
'' 機能名 : 取引先番号に該当する項目値を求める
'' 機能   :
'' 引数   : vntControl As Control   取引先番号コントロール
'' 戻り値 : String   取引先名称
'' 備考   :
''------------------------------------------------------------------------------
'    Dim strChar As String    '抜き出したキー
'
'    '引数がコントロールの場合
'    If IsObject(vntControl) Then
'        Select Case UCase(Left(vntControl.Name, 3))
'            Case "TXT"
'                strChar = vntControl.Text
'            Case "LBL"
'                strChar = vntControl.Caption
'            Case "CBO", "LST"
'                strChar = vntControl.List(vntControl.ListIndex)
'            Case Else
'                strChar = vntControl
'        End Select
'    Else
'        strChar = vntControl
'    End If
'
'    GF_Com_Cifget = GF_Com_CifName(strChar)
'
'End Function
'
' 2017/01/11 ▼ M.Tanaka K545 CSプロセス改善  ADD
Public Function GF_CheckStartToEnd(dtStart As Variant, dtEnd As Variant, intCheck As Integer, intKbn As Integer) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' 機能名    :  日付の期間チェック関数
' 機能      :　日付の期間をチェックする
' 引数      :  dtStart      As Variant      開始日
'              dtEnd        As Variant      終了日
'              intCheck     As integer      チェック期間
'              strKbn       As String       チェック区分 (1：年、2：月、3：日)
' 戻り値    :  True ： 一致 / False : 不一致
' 備考      :
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim dtStart_After As Date 'チェック期間加算後開始日
    
    GF_CheckStartToEnd = False
    
    If IsNull(dtStart) = True Or IsDate(dtStart) = False Then
        'NullもしくはDate型ではない場合falseで返す
        Exit Function
    End If
    
    If IsNull(dtEnd) = True Or IsDate(dtEnd) = False Then
        'NullもしくはDate型ではない場合falseで返す
        Exit Function
    End If
    
    Select Case intKbn
        Case CSTE_Year
            dtStart_After = DateTime.DateAdd("yyyy", intCheck, CDate(dtStart))
        Case CSTE_Month
            dtStart_After = DateTime.DateAdd("m", intCheck, CDate(dtStart))
        Case CSTE_Date
            dtStart_After = DateTime.DateAdd("d", intCheck, CDate(dtStart))
        Case Else
            GoTo ErrHandler
    End Select
    
    If dtStart_After = CDate(dtEnd) Then
        GF_CheckStartToEnd = True
    End If
    
    Exit Function
ErrHandler:
    Call GS_ErrorHandler("GF_CheckStartToEnd")
    Err.Raise Number:=vbObjectError, Description:="GF_CheckStartToEndでエラーが発生しました。"
End Function
' 2017/01/11 ▲ M.Tanaka K545 CSプロセス改善  ADD
