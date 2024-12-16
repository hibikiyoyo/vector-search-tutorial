Attribute VB_Name = "basMsgFunc"
' @(h) MsgFunc.bas  ver1.00 ( 2003/02/05 N.Kigaku )
'------------------------------------------------------------------------------
' @(s)
'   �v���W�F�N�g�� : TLF��ۼު��
'   ���W���[�����@ : basMsgFunc
'   �t�@�C�����@�@ : MsgFunc.bas
'   �o�[�W�����@�@ : 1.00
'   �@�\�����@�@�@ : ү���ޏ����֘A
'   �쐬�ҁ@�@�@�@ : N.Kigaku
'   �쐬���@�@�@�@ : 2003/02/05
'   �C�������@�@�@ : 2004/03/22 N.Kigaku GF_ExeLogOut �ǉ�
'                    2004/08/24 N.Kigaku GF_GetMsg_Addition��ү�����ޯ���\����
'                                        ���݂�ύX�ł���悤�ɏC���B
'                    2004/09/10 N.Kigaku GF_GetMsg_MasterMente �ǉ�
'                    2005/07/05 N.Kigaku GF_GetMsg_Addition, GF_GetMsg_MasterMente��On Error���ǋL
'                    2006/12/05 N.Kigaku �׸�8.1.7 Nocache�Ή� �������AReadOnly����Nocache�ɕύX
'                    2007/07/30 N.Kigaku GF_WriteLogData��۸ޏo�͂�հ��ID��7���Ő؂�悤�ɏC��
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ���錾
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ���W���[���萔�錾
'------------------------------------------------------------------------------
Private Const mstrCommonErrMsgCD As String = "WTK009"     ''���ʴװү���޺���

'------------------------------------------------------------------------------
' ���W���[���ϐ��錾
'------------------------------------------------------------------------------
Private mstrLogFile         As String       ''�߽�t��۸�̧�ٖ�      �i�����è�j
Private mblnMsgDispFlg      As Boolean      ''�װү�����ޯ���\���׸ށi�����è�j
Private mstrUserID          As String       ''հ��ID�i�����è�j
Private mstrPGMCD           As String       ''��۸���CD�i�����è�j
Private mstrTerminalCD      As String       ''�[��CD�i�����è�j

'ү���ފi�[�z�� START>>>>>
Private Type TYPE_MSG
    TYPE_MSG_CD     As String 'ү����CD
    TYPE_MSG_NAIYO  As String 'ү���ޓ��e
End Type
Private ADD_TYPE_MSG() As TYPE_MSG
'<<<<<END



Public Property Get LogFile() As String
'------------------------------------------------------------------------------
' �@�\���@�@: ۸�̧�ٖ������è
' �@�\�@�@�@:
' �����@�@�@: �Ȃ�
' �߂�l�@�@: �߽�t��۸�̧�ٖ�
' �@�\�����@: �����è�̒l��߂�
'------------------------------------------------------------------------------
    LogFile = mstrLogFile
End Property

Public Property Let LogFile(ByVal strLog As String)
'------------------------------------------------------------------------------
' �@�\���@�@: ۸�̧�ٖ������è
' �@�\�@�@�@:
' �����@�@�@: ByVal strLog As String    ''�߽�t��۸�̧�ٖ�
' �߂�l�@�@: �Ȃ�
' �@�\�����@: �����è�ɒl������
'------------------------------------------------------------------------------
    mstrLogFile = strLog
End Property

Public Property Get DispErrMsgFlg() As Boolean
'------------------------------------------------------------------------------
' �@�\���@�@: �װү�����ޯ���\���׸������è
' �@�\�@�@�@:
' �����@�@�@: �Ȃ�
' �߂�l�@�@: �װү�����ޯ���\���׸�
' �@�\�����@: �����è�̒l��߂�
'------------------------------------------------------------------------------
    DispErrMsgFlg = mblnMsgDispFlg
End Property

Public Property Let DispErrMsgFlg(ByVal blnFlg As Boolean)
'------------------------------------------------------------------------------
' �@�\���@�@: �װү�����ޯ���\���׸������è
' �@�\�@�@�@:
' �����@�@�@: ByVal blnFlg As Boolean    ''�װү�����ޯ���\���׸�
' �߂�l�@�@: �Ȃ�
' �@�\�����@: �����è�ɒl������
'------------------------------------------------------------------------------
    mblnMsgDispFlg = blnFlg
End Property

Public Property Get UserID() As String
'------------------------------------------------------------------------------
' �@�\���@�@: հ��ID�����è
' �@�\�@�@�@:
' �����@�@�@: �Ȃ�
' �߂�l�@�@: հ��ID
' �@�\�����@: �����è�̒l��߂�
'------------------------------------------------------------------------------
    UserID = mstrUserID
End Property

Public Property Let UserID(ByVal strUserID As String)
'------------------------------------------------------------------------------
' �@�\���@�@: հ��ID�����è
' �@�\�@�@�@:
' �����@�@�@: ByVal strUserID As String    ''հ��ID
' �߂�l�@�@: �Ȃ�
' �@�\�����@: �����è�ɒl������
'------------------------------------------------------------------------------
    mstrUserID = strUserID
End Property

Public Property Get PGMCD() As String
'------------------------------------------------------------------------------
' �@�\���@�@: ��۸���CD�����è
' �@�\�@�@�@:
' �����@�@�@: �Ȃ�
' �߂�l�@�@: ��۸���CD
' �@�\�����@: �����è�̒l��߂�
'------------------------------------------------------------------------------
    PGMCD = mstrPGMCD
End Property

Public Property Let PGMCD(ByVal strPGMCD As String)
'------------------------------------------------------------------------------
' �@�\���@�@: ��۸���ID�����è
' �@�\�@�@�@:
' �����@�@�@: ByVal strPGMCD As String    ''��۸���ID
' �߂�l�@�@: �Ȃ�
' �@�\�����@: �����è�ɒl������
'------------------------------------------------------------------------------
    mstrPGMCD = strPGMCD
End Property

Public Property Get TerminalCD() As String
'------------------------------------------------------------------------------
' �@�\���@�@: �[��CD�����è
' �@�\�@�@�@:
' �����@�@�@: �Ȃ�
' �߂�l�@�@: �[��CD
' �@�\�����@: �����è�̒l��߂�
'------------------------------------------------------------------------------
    TerminalCD = mstrTerminalCD
End Property

Public Property Let TerminalCD(ByVal strTerminalCd As String)
'------------------------------------------------------------------------------
' �@�\���@�@: �[��CD�����è
' �@�\�@�@�@:
' �����@�@�@: ByVal strTerminalCD As String    ''�[��CD
' �߂�l�@�@: �Ȃ�
' �@�\�����@: �����è�ɒl������
'------------------------------------------------------------------------------
    mstrTerminalCD = strTerminalCd
End Property



Public Function GF_MsgBox(sTITLE As String, sMSG As String, _
                            sBTN As String, sICON As String, _
                              Optional iDefBTN As Integer = 1) As Integer
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:�@ү�����ޯ���\��
' �@�\�@�@�@:�@ү�����ޯ����\������
' �����@�@�@:�@[I] sTITLE     As String    ''����
' �@�@�@�@�@ �@[I] sMSG       As String    ''ү����
' �@�@�@�@�@ �@[I] sBTN       As String    ''��������
' �@�@�@�@�@ �@[I] sICON      As String    ''��������
'             iDefBTN    As Integer   ''������̫��̫����ʒu
'             1 = ��1���� / 2 = ��2���� / 3 = ��3���� / 4 = ��4����
' �߂�l�@�@:�@1 = OK / 2 = CANCEL / 6 = YES / 7 = NO
' �@�@�@�@�@ �@0 = ERROR
' �@�\�����@:
'------------------------------------------------------------------------------
    
    Dim mlngStyle As Long
    Dim mintDefBtn As Integer
    
    On Error GoTo ErrHandler

    '[����]
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

    '[����]
    Select Case UCase(Trim(sICON))
        Case "C"    '[�x��]
                    mlngStyle = mlngStyle + vbCritical
                    '�ް�߉��炷
                    Beep
        Case "Q"    '[�₢���킹]
                    mlngStyle = mlngStyle + vbQuestion
        Case "E"    '[����]
                    mlngStyle = mlngStyle + vbExclamation
                    '�ް�߉��炷
                    Beep
        Case "I"    '[���]
                    mlngStyle = mlngStyle + vbInformation
    End Select
    
    '[��̫������]
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
    
    'ү�����ޯ���\��
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
' �@�\���@�@:�@ү���ޕ\���iDB�Łj
' �@�\�@�@�@:�@ү�����ޯ�����ͽð���ް�Ɏw��ү���ނ�\������
' �����@�@�@:�@[I] sTITLE     As String    ''����
' �@�@�@�@�@ �@[I] sMSGID     As String    ''ү����ID
' �@�@�@�@�@ �@[I] sBTN       As String    ''��������
' �@�@�@�@�@ �@[I] sICON      As String    ''��������
'             [I] iDefBTN    As Integer   ''������̫��̫����ʒu
'             1 = ��1���� / 2 = ��2���� / 3 = ��3���� / 4 = ��4����
' �߂�l�@�@:�@1 = OK / 2 = CANCEL / 6 = YES / 7 = NO / 9 = �ð���ް�֏o��
' �@�@�@�@�@ �@0 = ERROR
' �@�\�����@:�@ү���ނ��ް��ް�����擾
' �@�@�@�@�@ �@�Y�������ް����Ȃ��ꍇ�́A�װү���ނ�\��
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
    
    '[����]
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
    
    '[����]
    Select Case UCase(Trim(sICON))
        Case "C"    '[�x��]
                    mlngStyle = mlngStyle + vbCritical
                    '�ް�߉��炷
                    Beep
        Case "Q"    '[�₢���킹]
                    mlngStyle = mlngStyle + vbQuestion
        Case "E"    '[����]
                    mlngStyle = mlngStyle + vbExclamation
                    '�ް�߉��炷
                    Beep
        Case "I"    '[���]
                    mlngStyle = mlngStyle + vbInformation
    End Select
    
    '[��̫������]
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
    
    'SQL������
    sSQL = ""
    sSQL = "SELECT * FROM THJMSG WHERE MSGCD = '" & UCase(Trim(sMSGID)) & "'"
    
    '�޲ž�Đ���
    Set oDynaset = gOraDataBase.CreateDynaset(sSQL, ORADYN_NOCACHE)
    
    '�ް������������ꍇ
    If (oDynaset.EOF = False) Then
        mstrOutMsg = GF_VarToStr(oDynaset![MSGNAIYO])
        mintOutFlg = GF_VarToNum(oDynaset![OUTFLG])
    Else
        '�޲ž�ĉ��
        Set oDynaset = Nothing
        Beep
        intRet = MsgBox("ү����ð��قɓo�^����Ă��Ȃ�ү����ID���w�肳��܂����B" & vbCrLf & _
                        "MSGID -> [" & UCase(Trim(sMSGID)) & "]", vbOKOnly + vbExclamation + mintDefBtn, "GF_MsgBoxDB")
        GF_MsgBoxDB = 0
        Exit Function
    End If
    
    '�޲ž�ĉ��
    Set oDynaset = Nothing
    
    '�o�͐�I��
    If (mintOutFlg = 0) Then
        'ү�����ޯ���\��
        GF_MsgBoxDB = MsgBox(GF_CnvCtrChar(mstrOutMsg), mlngStyle + mintDefBtn, sTITLE)
    Else
        '�ð���ް�\��
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
' �@�\���@�@:�@ү���ގ擾
' �@�\�@�@�@:�@ү���ނ��ް��ް�����擾����
' �����@�@�@:�@[I] strMsgCD As String             ''ү����CD
' �߂�l�@�@:�@�擾����ү����
' �@�@�@�@�@ �@�Y���ް����Ȃ����ʹװ�̎���MsgErr[ү���޺���]
' �@�\�����@:
'------------------------------------------------------------------------------

    Dim oDynaset   As OraDynaset
    Dim sSQL       As String

    On Error GoTo ErrHandler

    'SQL������
    sSQL = ""
    sSQL = "SELECT * FROM THJMSG WHERE MSGCD = '" & UCase(Trim(strMsgCD)) & "'"

    '�޲ž�Đ���
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
' �@�\���@�@:�@ү���ޏ��擾
' �@�\�@�@�@:�@ү���ޏ����ް��ް�����擾����
' �����@�@�@:�@[I] strMsgCD As String               ''ү���޺���
' �@�@�@�@�@�@ [O] strMsg As String                 ''ү����
' �@�@�@�@�@�@ [O] strMsgLevel As String            ''ү��������
'             [I]  blnNonErrFlg As Boolean         ''�װ�����L���׸�
'                                                      False:�װ�����L��,True:�Ȃ�
' �߂�l�@�@:�@�擾����ү����
' �@�@�@�@�@ �@�Y���ް����Ȃ����ʹװ�̎���MsgErr[ү���޺���]
' �@�\�����@:
'------------------------------------------------------------------------------

    Dim oDynaset   As OraDynaset
    Dim sSQL       As String

    On Error GoTo ErrHandler

    'SQL������
    sSQL = ""
    sSQL = "SELECT MSGNAIYO,MSGLEVEL FROM THJMSG WHERE MSGCD = '" & UCase(Trim(strMsgCD)) & "'"

    '�޲ž�Đ���
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
' �@�\���@�@:�@ү���ގ擾�i�t��������t���j
' �@�\�@�@�@:�@ү���ނ��ް��ް�����擾����
'
' �����@�@�@: [I] strMsgCD As String                ''ү���޺���
'�@�@�@�@�@�@ [I] vntAddMsg As Variant              ''�t��������̔z��@�i�Y����0����J�n�j
'                                                       1���̎��͔z��łȂ��Ă�OK
'�@�@�@�@�@�@ [I] ByVal blnDispFlg As Boolean       ''�װү�����ޯ���o���׸�
'�@�@�@�@�@�@ [I] ByVal blnLogTblFlg As Boolean     ''۸�ð��ُo���׸�
'�@�@�@�@�@�@ [I] ByVal strInfo As String           ''�t�^���
'�@�@�@�@�@�@ [I] ByVal blnTivoliLogFlg As Boolean  ''Tivoli۸ޏo���׸�
' �@�@�@�@ �@ [I] ByVal strICON As String           ''��������  2004/08/24 Add by N.Kigaku
'
' �߂�l�@�@: �擾����ү����
'�@�@�@�@�@�@ �Y���ް����Ȃ����ʹװ�̎��� "MsgErr[ү���޺���]"
' �@�\�����@: DB����擾����ү���ނ�%1�`%n�܂ŕt��������Œu��������
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
    
    
    '' �z��̐��𐔂���
    If IsArray(vntAddMsg) = True Then
        intMsgCount = UBound(vntAddMsg) + 1
    Else
        intMsgCount = 0
    End If
    
    '' SQL���쐬����
    strSQL = "SELECT MSGNAIYO,MSGLEVEL FROM THJMSG WHERE MSGCD = '" & UCase(Trim(strMsgCD)) & "'"

    '' �޲ž�Đ���
    Set oDynaset = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)

    If oDynaset.EOF = False Then
        strMessage = GF_VarToStr(oDynaset![MSGNAIYO])
        strMsgLevel = GF_VarToStr(oDynaset![MSGLEVEL])
        
        '�t��������ɕύX
        If intMsgCount > 0 Then
        
            '�z��
            For i = 1 To intMsgCount
                strTemp = "%" & CStr(i)
                strMessage = Replace(strMessage, strTemp, vntAddMsg(i - 1))
            Next
            
        ElseIf (Len(Trim(vntAddMsg)) > 0) Then
            '�z��ȊO
            strTemp = "%1"
            strMessage = Replace(strMessage, strTemp, vntAddMsg)
        End If
        
    Else
        strMessage = " MsgErr[" & strMsgCD & "] "
        strMsgLevel = ""
    End If

    Set oDynaset = Nothing
    
    '' ү�����ޯ���\��
    If blnDispFlg = True Then
    
'2004/08/24 Update by N.Kigaku
''ү�����ޯ���\���ű��݂�ύX�ł���悤�ɏC���B������strICON��ǉ��B
        strICON = UCase(strICON)
        If (strICON <> "C") And (strICON <> "Q") And (strICON <> "I") And (strICON <> "E") Then
            strICON = "E"
        End If
    
        '' "@@"�����s�ɒu��������
        strCnvMsg = GF_CnvCtrChar(strMessage)
        
        If Forms.Count > 0 Then
            '̫�т����鎞��̫�т�Caption��\������
            intRet = GF_MsgBox(Screen.ActiveForm.Caption, strCnvMsg, "OK", strICON)
        Else
            '���ع�������ق�\������
            intRet = GF_MsgBox(App.Title, strCnvMsg, "OK", strICON)
        End If
    End If

    '' ۸�ð��ُo��
    If blnLogTblFlg = True Then
        intRet = GF_WriteLogData(mstrUserID, mstrPGMCD, strMsgCD, strMsgLevel, strMessage, strInfo, mstrTerminalCD)
    End If
    
    '' TIVOLI�p۸�̧�ُo��
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
' �@�\���@�@:�@�ϊ������iVARIANT -> STRING�j
' �@�\�@�@�@:
' �����@�@�@:�@[I] vVALUE As Variant    ''�ϊ��O��ر���ް�
' �߂�l�@�@:�@�ϊ��㕶����
' �@�\�����@:�@������NULL�̏ꍇ�́A""(����0�̕�����)��Ԃ�
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
' �@�\���@�@:�@�ϊ������iVARIANT -> DOUBLE�j
' �@�\�@�@�@:
' �����@�@�@:�@[I] vVALUE As Variant    ''�ϊ��O��ر���ް�
' �߂�l�@�@:�@�ϊ��㐔�l
' �@�\�����@:�@������NULL�̏ꍇ�́A0��Ԃ�
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
' �@�\���@�@:�@���s�w�������ϊ������i@@ -> vbCrLf�j
' �@�\�@�@�@:
' �����@�@�@:�@[I] sMSG As String       ''�ϊ��O������
' �߂�l�@�@:�@���� - �ϊ��㕶����
' �@�@�@�@�@:�@���s - �ϊ��O������
' �@�\�����@:
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    ''������u������
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
' �@�\���@�@:�@���s�w�������폜�����i@@ -> DELETE�j
' �@�\�@�@�@:
' �����@�@�@:�@[I] sMSG As String       ''�ϊ��O������ / �ϊ��㕶����
' �߂�l�@�@:�@True = ���� / False = ���s
' �@�\�����@:
'------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ''������u������
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
' �@�\���@�@:�@�װ�����
' �@�\�@�@�@:�@�װ���̏������s��
' �����@�@�@:�@[I] strLocation As String           ''�װ�����ꏊ
' �@�@�@�@�@ �@[I] strAdditon As String            ''�t�����
'             [I] intDefSqlCdErr As Integer       ''SQL�������Ұ�����̫�Ĵװ���ޔ��� ��̫��:0
' �@�@�@�@�@ �@[I] strMsgCD As String              ''ү���޺���
' �߂�l�@�@:�@�Ȃ�
' �@�\�����@:
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
    strMsgLevel = "1"   '�v���I�װ��ݒ�

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
    
'2003/08/05 �C��
        'ү���ށAү�������ق��擾
        If Len(Trim(strMsgCD)) > 0 Then
            intRet = GF_GetMsgInfo(strMsgCD, strMsg, strMsgLevel, True)
        End If
    
        '۸ޏo��
        intRet = GF_LogOut(strErrType, strLocation, CStr(lngErrNum), GF_DelChrCode(strErrMsg) & IIf(Len(Trim(strAdditon)) = 0, "", IIf(Len(Trim(strAdditon)) > 0, vbCrLf & Space(4) & strAdditon, "")), 2, strMsgLevel)
        
        '۸�ð��ُo��
        intRet = GF_WriteLogData(mstrUserID, mstrPGMCD, strMsgCD, strMsgLevel, strMsg, strLocation & IIf(Len(Trim(strErrMsg)) > 0, " , ", "") & strErrMsg & IIf(Len(Trim(strAdditon)) > 0, " , " & strAdditon, ""), mstrTerminalCD)

        'ү�����ޯ���\���׸ނ�True�̎���ү�����ޯ����\������
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
' �@�\���@�@:�@�װ�ر
' �@�\�@�@�@:�@�װ��ر����
' �����@�@�@:�@[I] bVB        VB�װ�װ�ر�׸�
'             [I] bParam     ���Ұ��װ�ر�׸�
'             [I] bDataBase  �ް��ް��װ�ر�׸�
'             [I] bSession   ����ݴװ�ر�׸�
'
' �߂�l�@�@:�@�Ȃ�
' �@�\�����@:
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
' �@�\���@�@:�@۸ޏo�́i÷�ĔŁj
' �@�\�@�@�@:�@�װ���̏󋵂�÷��̧�ق֏o�͂���
' �����@�@�@:�@[I] strErrType  As String        ''�װ����(VB or ORACLE)
' �@�@�@�@�@ �@[I] strLocation As String        ''�װ�����ꏊ
'             [I] strErrNum As String          ''�װ�ԍ� or �װ����
' �@�@�@�@�@ �@[I] strErrMSG As String          ''�װү����
'             [I] intTivoliLogKbn As Integer   ''Tivoli۸ޏo�͋敪
'             [I] strErrMsgLvl As String       ''�װү��������
'                                                 0:�ʏ�o��, 1:Tivoli۸ޏo��, 2:����
' �@�\�����@:
' ���l�@�@�@:  ۸�̧�ٖ���LogFile�����è�ɾ�Ă��Ă�������
'------------------------------------------------------------------------------

    Dim intRet     As Integer
    Dim strYmd     As String
    Dim strHms     As String
    Dim strMsg     As String
    Dim FileID     As Integer

    On Error GoTo ErrHandler
    
    '�o��ү���ނ̕ҏW
    strYmd = Format(Date, "yyyy/mm/dd")
    strHms = Format(Time, "hh:mm:ss")

    '۸ޏo��ү���ނ̕ҏW
    strMsg = strYmd & " " & strHms & ", " & gstrUserID & ", " & App.EXEName & vbCr & _
                        strErrType & ", " & strLocation & ", " & "[" & strErrNum & "] " & strErrMsg
    
    If (intTivoliLogKbn = 0) Or (intTivoliLogKbn = 2) Then
    
        If Len(Trim(mstrLogFile)) = 0 Then
            Err.Raise Number:=vbObjectError, Description:="̧�ٖ����ݒ肳��Ă��܂���B"
        End If
    
        '�ިڸ�ؑ�������
        If GF_DirCheck(mstrLogFile) = False Then
            Err.Raise Number:=vbObjectError, Description:="�o�͐悪���݂��܂���B" & "[" & mstrLogFile & "]"
        End If
        
        FileID = FreeFile
        
        Open mstrLogFile For Append As #FileID
        Print #FileID, strMsg
        Close #FileID
        
    End If
    
    'TIVOLI�p۸ޏo��
    If basMainFunc.TVL_LOG_FLG = True Then
        If (intTivoliLogKbn = 1) Or (intTivoliLogKbn = 2) Then
        
            '�ިڸ�ؑ�������
            If GF_DirCheck(basMainFunc.TVL_LOG_DIR) = False Then
                Err.Raise Number:=vbObjectError, Description:="�o�͐悪���݂��܂���B" & "[" & basMainFunc.TVL_LOG_DIR & "]"
            End If
        
            If (strErrMsgLvl = "1") Or (strErrMsgLvl = "2") Then
        
                FileID = FreeFile
            
                If strErrMsgLvl = "1" Then
                    '�װ۸�
                    Open basMainFunc.TVL_LOG_DIR & App.EXEName & "_ERRO.ERR" For Append As #FileID
                ElseIf strErrMsgLvl = "2" Then
                    'ܰ�ݸ�۸�
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
    'ү�����ޯ���\���׸ނ�True�̎���ү�����ޯ����\��
    If mblnMsgDispFlg = True Then
        intRet = GF_MsgBox("ERROR NO. " & Err.Number & " - GF_LogOut", Err.Description, "OK", "E")
    End If
    GF_LogOut = False

End Function

Public Function GF_LogOutDB(strErrType As String, strLocation As String, strMsgID As String) As Integer
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:�@۸ޏo�́iү���޺��ޔŁj
' �@�\�@�@�@:�@ү���޺��ނ��ү���ނ�DB����擾���A÷��̧�ق֏o�͂���
' �����@�@�@:�@[I] strErrType  As String      ''�װ����(VB or ORACLE)
' �@�@�@�@�@ �@[I] strLocation As String      ''�װ�����ꏊ
' �@�@�@�@�@ �@[I] strMsgID    As String      ''�װү���޺���
' �߂�l�@�@:�@�Ȃ�
' ���l     �F
'------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Dim strMsg   As String      ''ү���ޓ��e
    Dim intRet   As Integer
    
    ''ү���ގ擾
    strMsg = GF_GetMsg(strMsgID)
    If strMsg = "" Then
        'ү���ގ擾���s
        strMsg = "ү���ގ擾�s�\"
    End If
        
    ''۸ޏo��
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
' �@�\���@�@:�@��۸��яI��۸ޏo�́i÷�ĔŁj
' �@�\�@�@�@:�@��۸��яI������۸ނ��o�͂���B
' �����@�@�@:  [I] strAddMsg As String   ''�t��ү����
' �@�\�����@:
' ���l�@�@�@:  ۸�̧�ٖ���LogFile�����è�ɾ�Ă��Ă�������
'------------------------------------------------------------------------------

    Dim intRet     As Integer
    Dim strMsg     As String
    Dim FileID     As Integer

    On Error GoTo ErrHandler

    FileID = -1
        
    'TIVOLI�p۸ޏo��
    If basMainFunc.TVL_LOG_FLG = True Then
    
        '۸ޏo��ү���ނ̕ҏW
        strMsg = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss") & ", " & gstrUserID & ", " & App.EXEName & IIf(strMsg = "", "", ", " & strAddMsg)
    
        '�ިڸ�ؑ�������
        If GF_DirCheck(basMainFunc.TVL_LOG_DIR) = False Then
            Err.Raise Number:=vbObjectError, Description:="�o�͐悪���݂��܂���B" & "[" & basMainFunc.TVL_LOG_DIR & "]"
        End If
    
        FileID = FreeFile
    
        '۸ޏo��
        Open basMainFunc.TVL_LOG_DIR & App.EXEName & "_EXE.LOG" For Append As #FileID
        Print #FileID, strMsg
        Close #FileID
                
    End If
    
    GF_ExeLogOut = True
    
    Exit Function

ErrHandler:
    'ү�����ޯ���\���׸ނ�True�̎���ү�����ޯ����\��
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
'' �@�\���@�@: ۸��ް��o��
'' �@�\�@�@�@: ۸ނ����C)۸��ް��ɏ����o��
'' �����@�@�@: [I] strUserID As String        ''հ��ID
''�@�@�@�@�@ : [I] strProgCD As String        ''��۸��Ѻ���
''�@�@�@�@�@ : [I] strMsgCD As String         ''ү���޺���
''�@�@�@�@�@ : [I] strMsgLevel As String      ''ү��������
''�@�@�@�@�@ : [I] strMsg As String           ''ү����
''�@�@�@�@�@ : [I] strNote As String          ''�t�^���
''�@�@�@�@�@ : [I] strTerminalCD As String    ''�[�����ށ@(�ȗ����F��j
'' �߂�l�@�@:
'' ���l�@�@�@:
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
'    '' Null����
'    If strUserID = "" Then strUserID = " "
'    If strProgCD = "" Then strProgCD = " "
'    If strMsgCD = "" Then strMsgCD = " "
'    If strMsgLevel = "" Then strMsgLevel = " "
'    If strMsg = "" Then strMsg = " "
'
'    '' �V���O���N�H�[�e�[�V������u������
'    strMsg = Replace(strMsg, "'", "''")
'    strNote = Replace(strNote, "'", "''")
'
'    '' ���O�e�[�u���ɏ����o��
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
'    '' SQL�����s����
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
'    '۸�̧�ُo��
'    Call GF_LogOut("VB6", "GF_WriteLogData", strErrNum, strErrMsg)
'
'    'ү�����ޯ���\���׸ނ�True�̎���ү�����ޯ����\��
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
' �@�\���@�@: ۸��ް��o��
' �@�\�@�@�@: ۸ނ����C)۸��ް��ɏ����o��
' �����@�@�@: [I] strUserID As String        ''հ��ID
'�@�@�@�@�@ : [I] strProgCD As String        ''��۸��Ѻ���
'�@�@�@�@�@ : [I] strMsgCD As String         ''ү���޺���
'�@�@�@�@�@ : [I] strMsgLevel As String      ''ү��������
'�@�@�@�@�@ : [I] strMsg As String           ''ү����
'�@�@�@�@�@ : [I] strNote As String          ''�t�^���
'�@�@�@�@�@ : [I] strTerminalCD As String    ''�[�����ށ@(�ȗ����F��j
' �߂�l�@�@:
' ���l�@�@�@:
'--------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL              As String
    Dim strErrMsg           As String           '��������ү����
    Dim lngErrNum           As Long             '�װNo
    Dim lclsOraClass        As New clsOraClass  '�ı�ތďo���p
    Dim blnCreOraClass      As Boolean          '�׽��޼ު�č쐬���׸�

    GF_WriteLogData = False

    ''Null�̏ꍇ�A�u�����N1����ݒ�
    If strUserID = "" Then strUserID = " "
    If strProgCD = "" Then strProgCD = " "
    If strMsgCD = "" Then strMsgCD = " "
    If strMsgLevel = "" Then strMsgLevel = " "
    If strMsg = "" Then strMsg = " "

    strErrMsg = ""

    '�ı�ޗp Object �錾
    Set lclsOraClass = New clsOraClass
    Set lclsOraClass.OraDataBase_Strcall = gOraDataBase
    blnCreOraClass = True

    '���ް�װ��ؾ��
    Call lclsOraClass.ErrReset_Strcall

    '�޲��ޕϐ��ǉ�
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_USR_ID", Right(Trim(strUserID), 7)   'հ��ID
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_TER_CD", strTerminalCd       '�[��CD
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_PGM_CD", strProgCD           '��۸���CD
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_MSG_CD", strMsgCD            'ү����CD
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_MSG_LV", strMsgLevel         'ү��������
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_MSG", strMsg                 'ү����
    lclsOraClass.Add_Binds ORAPARM_INPUT, ORATYPE_CHAR, "p_INVEST_INFO", strNote        '�t�^���
    lclsOraClass.Add_Binds ORAPARM_OUTPUT, ORATYPE_VARCHAR2, "p_ErrMsg", strErrMsg      '�װү����

    If (lclsOraClass.ErrCode_Strcall <> 0 Or lclsOraClass.ErrText_Strcall <> "") Then
        Err.Raise Number:=lclsOraClass.ErrCode_Strcall, Description:=lclsOraClass.ErrText_Strcall
    End If

    '�ı���̧ݸ��݂ւ̈������
    strSQL = ""
    strSQL = strSQL & "BEGIN "
    strSQL = strSQL & ":sql_code:=TOS_SE_LOGDATA_OUT"      '�ı���̧ݸ��ݖ�
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

    '���޴װ��ؾ��
    Call lclsOraClass.ErrReset_Strcall

    'SQL�̎��s
    lclsOraClass.ExecSql_Strcall strSQL

    '��������ү���ގ擾
    strErrMsg = GF_VarToStr(gOraParam!p_ErrMsg)

    If (gOraParam!sql_code = -1) Then
        '�X�g�A�h�V�X�e���G���[
        Err.Raise Number:=vbObjectError, Description:=strErrMsg
    End If

    '�޲������Ұ���̑S���
    lclsOraClass.RemoveAll
    Set lclsOraClass = Nothing
    blnCreOraClass = False
    
    '���Ұ��̏�����
    Call GS_ErrorClear(False, True, True, False)

    GF_WriteLogData = True
    Exit Function

ErrHandler:
    ''�װ�����
    
    lngErrNum = Err.Number
    strErrMsg = Err.Description
    
    '÷��۸ޏo��
    Call GF_LogOut("ORACLE", "GF_WriteLogData", CStr(lngErrNum), strErrMsg)
    
    'ү�����ޯ���\���׸ނ�True�̎���ү�����ޯ����\��
    If mblnMsgDispFlg = True Then
        Call GF_MsgBox("ERROR NO. " & lngErrNum & " - GF_WriteLogData", strErrMsg, "OK", "E")
    End If
    
    '�׽��޼ު�Ẳ��
    If blnCreOraClass = True Then
        '�޲������Ұ���̑S���
        lclsOraClass.RemoveAll
        Set lclsOraClass = Nothing
    End If

End Function


'Public Sub GS_AppLog(strLogCode As String, strLogValue As String, strNote As String)
''------------------------------------------------------------------------------
'' @(f)
''
'' �@�\���@�@:�@�ғ�۸ޏo��
'' �@�\�@�@�@:�@�ғ��󋵂��ް��ް��֏o�͂���
'' �����@�@�@:�@sLogCode  As String      ''۸޺���
'' �@�@�@�@�@ �@sLogValue As String      ''۸ޓ��e
'' �@�@�@�@�@ �@sNote     As String      ''���l
'' �߂�l�@�@:�@�Ȃ�
'' �@�\�����@:�@MLTT_028(۸ޏ��ð���)�֏o��
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
'    '�ı�ޗp Object �錾
'    Set oSTORED = New clsOraClass
'    Set oSTORED.OraDataBase_Strcall = gOraDataBase
'
'    '���ް�װ��ؾ��
'    Call oSTORED.ErrReset_Strcall
'
'    '�޲��ޕϐ��ǉ�
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
'        '�޲��ޕϐ��̑S���
'        oSTORED.RemoveAll
'
'        Exit Sub
'    End If
'
'    '�ı����ۼ��ެ�ւ̈������
'    sSQL = "BEGIN "
'    sSQL = sSQL & "MIPC_001 ("          '�ı����ۼ��ެ��
'    sSQL = sSQL & ":p_SHAIN_CD, "       'հ��ID
'    sSQL = sSQL & ":p_SHOZOKU_CD, "     '������ٰ�ߺ���
'    sSQL = sSQL & ":p_GYOMU_CD, "       '�Ɩ�����
'    sSQL = sSQL & ":p_SAGYO_CD, "       '��ƺ���
'    sSQL = sSQL & ":p_LOG_CD, "         '۸޺���
'    sSQL = sSQL & ":p_LOGNAIYO, "       '۸ޓ��e
'    sSQL = sSQL & ":p_BIKO, "           '���l
'    sSQL = sSQL & ":p_UPD_MAN"          '�X�V��
'    sSQL = sSQL & "); "
'    sSQL = sSQL & "END;"
'
'    '���޴װ��ؾ��
'    Call oSTORED.ErrReset_Strcall
'
'    'SQL�̎��s
'    oSTORED.ExecSql_Strcall sSQL
'
'    If (oSTORED.ErrCode_Strcall <> 0 Or oSTORED.ErrText_Strcall <> "") Then
'
'        '�޲��ޕϐ��̑S���
'        oSTORED.RemoveAll
'
'        Exit Sub
'    End If
'
'    '�o�C���h�p�����[�^�̑S���
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
' �@�\���@�@:�@ײ�̨��ޕ����폜(ORACLE�װү���ޗp)
' �@�\�@�@�@:�@ORACLE����̴װү���ނ̍Ō���ɕt������Ă���ײ�̨��ޕ������폜
' �����@�@�@:�@[I] sMSG As String       ''ײ�̨��ޕ����폜�O������
' �߂�l�@�@:�@ײ�̨��ޕ����폜�㕶����
' �@�\�����@:
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
' �@�\���@�@:�@�ިڸ�؍쐬
' �@�\�@�@�@:�@�ިڸ�؂����݂��邩�������A�Ȃ��ꍇ�͍쐬����
' �����@�@�@:�@[I] strPath As String       ''�߽
' �߂�l�@�@:�@True = ���� / False = ���s
' ���l     �F
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
    '�װ������
    Call GS_ErrorHandler("GF_DirCheck")

End Function

'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :  �X�y�[�X��NULL�ϊ�
' �@�\      :  �����񂪋󔒂̂Ƃ���NULL�������Ԃ�
'              �󔒂łȂ��Ƃ���TRIM���āA�V���O���N�H�[�e�[�V�����ň͂񂾕������Ԃ�
' ����     �F [I] strChar   As String    ���肷�镶����
' �߂�l    :  ������
' ���l      :
'------------------------------------------------------------------------------
Public Function GF_ChangeSpaceToNull(strChar As String) As String
    GF_ChangeSpaceToNull = IIf(Len(Trim(strChar)) = 0, "IS NULL", "='" & Trim(strChar) & "'")
End Function

'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :  �X�y�[�X��NULL�ϊ�(���l�p)
' �@�\      :  �����񂪋󔒂̂Ƃ���NULL�������Ԃ�
'              �󔒂łȂ��Ƃ��͂��ƕ������߂�
' ����     �F [I] strChar   As String    ���肷�镶����
'          :  [I] blnEqualFlg As Boolea �C�R�[���t���t���O
' �߂�l    :  ������
' ���l      :
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
' �@�\���@�@: �X�y�[�X��NULL�ϊ�
' �@�\�@�@�@: �����񂪋󔒂̂Ƃ���NULL�������Ԃ�
'            �󔒂łȂ��Ƃ���TRIM���āA�V���O���N�H�[�e�[�V�����ň͂񂾕������Ԃ�
' �����@�@�@: [I] strChar   As String    ���肷�镶����
' �@�@�@�@�@: [I] blnEqualFlg As Boolean �C�R�[���L�薳���t���O�i�ȗ����F�L��j
' �߂�l�@�@:  ������
' ���l�@�@�@:
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
' �@�\���@�@:
' �@�\�@�@�@: ��������w�肵�������ň͂�
' �����@�@�@: [I] strChar   As String       �͂ޕ�����
' �@�@�@�@�@: [I] strEnclose As Boolean     �͂ޕ���
' �@�@�@�@�@: [I] blnEncloseFlg As Boolean  �͂ޕ����񂪋�̎��Ɉ͂����ۂ�
'                                         False :�͂�Ȃ�
'                                         True  :�͂�
' �߂�l�@�@: �͂񂾕�����
' ���l�@�@�@:
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
' �@�\���@�@:�@Ͻ����ү���ގ擾�i�t��������t���j
' �@�\�@�@�@:�@ү���ނ��ް��ް�����擾����
'
' �����@�@�@: [I] strMsgCD  As String               ''ү���޺���
'�@�@�@�@�@�@ [I] lngMaxRow As Long                 ''�z��
'�@�@�@�@�@�@ [I] vntAddMsg As Variant              ''�t��������̔z��@�i�Y����0����J�n�j
'                                                    1���̎��͔z��łȂ��Ă�OK
' �߂�l�@�@: �擾����ү����
'�@�@�@�@�@�@ �Y���ް����Ȃ����ʹװ�̎��� "MsgErr[ү���޺���]"
' �@�\�����@: DB����擾����ү���ނ�%1�`%n�܂ŕt��������Œu��������
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
    
    ''�߂�l�ݒ�
    GF_GetMsg_MasterMente = ""
    
    ''�ϐ�������
    strExistFlg = ""
    
    ''�������Ұ��̔z��̐��𐔂���
    If IsArray(vntAddMsg) = True Then
        intMsgCount = UBound(vntAddMsg) + 1
    Else
        intMsgCount = 0
    End If
    
    ''�ێ��z��ɊY����MSG���ނ����݂��邩����
    lngCounter = 1
    Do Until lngCounter > lngMaxRow
        If Trim(ADD_TYPE_MSG(lngCounter).TYPE_MSG_CD) = Trim(strMsgCD) Then
            strMessage = Trim(ADD_TYPE_MSG(lngCounter).TYPE_MSG_NAIYO)
            strExistFlg = "1"
            Exit Do
        End If
        lngCounter = lngCounter + 1
    Loop
    
    ''�Y��MSG�����݂��Ȃ��ꍇ�͋��ʊ֐����擾����
    If Trim(strExistFlg) = "" Then
        ''MSG�Č���
        If GF_GetMsgInfo(strMsgCD, strMessage, strMsgLevel) = False Then
            GoTo ErrHandler
        End If
        ''�z��Ē�`
        lngMaxRow = lngMaxRow + 1
        ReDim ADD_TYPE_MSG(lngMaxRow)
        ''�z���ް����
        ADD_TYPE_MSG(lngMaxRow).TYPE_MSG_CD = Trim(strMsgCD)
        ADD_TYPE_MSG(lngMaxRow).TYPE_MSG_NAIYO = Trim(strMessage)
    End If
    
    ''MSG���
    If intMsgCount > 0 Then
        '�z��
        For i = 1 To intMsgCount
            strTemp = "%" & CStr(i)
            strMessage = Replace(strMessage, strTemp, vntAddMsg(i - 1))
        Next
        
    ElseIf (Len(Trim(vntAddMsg)) > 0) Then
        '�z��ȊO
        strTemp = "%1"
        strMessage = Replace(strMessage, strTemp, vntAddMsg)
    End If
    
    ''�߂�l�Đݒ�
    GF_GetMsg_MasterMente = strMessage
    
    Exit Function
    
ErrHandler:
    
    Call GS_ErrorHandler("GF_GetMsg_MasterMente")

End Function


