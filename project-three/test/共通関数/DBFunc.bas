Attribute VB_Name = "basDBFunc"
' @(h) DBFunc.bas  ver1.00 ( 2000/08/30 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   �v���W�F�N�g��  : TLF��ۼު��
'   ���W���[����    : basDBFunc
'   �t�@�C����      : DBFunc.bas
'   Version        : 1.00
'   �@�\����       �F DB�ڑ��Ɋւ��鋤�ʊ֐�
'   �쐬��         �F T.Fukutani
'   �쐬��         �F 2000/08/30
'   �C������       �F 2007/10/30 N.Kigaku DB�ؒf���ɵ׸��޲��ޕϐ���j������悤�ɏC���B
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ���錾
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' �p�u���b�N�ϐ��錾
'------------------------------------------------------------------------------
'' ���ʗp
Public gOraSession          As OraSession       ''�Z�b�V������`
Public gOraDataBase         As OraDatabase      ''�f�[�^�x�[�X��`
Public gOraParam            As OraParameters    ''�p�����[�^�I�u�W�F�N�g

'' ���O�o�͗p
Public gWOraSession         As OraSession       ''�Z�b�V������`
Public gWOraDataBase        As OraDatabase      ''�f�[�^�x�[�X��`
'Public gWOraParam           As OraParameters    ''�p�����[�^�I�u�W�F�N�g


Public Function GF_DBOpen(strInstance As String, strUserID As String, strPassWord As String, _
                 Optional blnWOraSessionConnectFlag As Boolean = True) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\�� : DB�ڑ�
' �@�\   : �T�[�o�[���I���N���Ƃ̃Z�b�V�����m��
' ����   : strInstance As String    ''DB��
'          strUserID   As String    ''DB�ڑ�հ�ޖ�
'          strPassWord As String    ''DB�ڑ��߽ܰ��
'          blnWOraSessionConnectFlag As Boolean    ''۸ޏo�͗pOracleDB�ڑ��׸�(TRUE:�ڑ��AFALSE:��ڑ�)
' �߂�l : True = ���� / False = ���s
' ���l   :
'------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    GF_DBOpen = False
   
    ''Oracle�f�[�^�x�[�X�̐ڑ��E�I�[�v��
    Set gOraSession = CreateObject("OracleInProcServer.XOraSession")
    
    Set gOraDataBase = gOraSession.DbOpenDatabase(strInstance, _
                                                    strUserID _
                                                    & "/" & _
                                                    strPassWord, 0&)

    Set gOraParam = gOraDataBase.Parameters
    
    ''�װ���ށA�װү���ގ擾�p���Ұ��ݒ�
    gOraParam.Add "sql_code", 0, ORAPARM_OUTPUT
    gOraParam!sql_code.serverType = ORATYPE_NUMBER
    gOraParam.Add "sql_errm", "", ORAPARM_OUTPUT
    gOraParam!sql_errm.serverType = ORATYPE_VARCHAR2
    
    If blnWOraSessionConnectFlag = True Then
        ''���O�o�͗pOracle�f�[�^�x�[�X�̐ڑ��E�I�[�v��
        Set gWOraSession = CreateObject("OracleInProcServer.XOraSession")
        
        Set gWOraDataBase = gWOraSession.DbOpenDatabase(strInstance, _
                                                       strUserID _
                                                        & "/" & _
                                                        strPassWord, 0&)
    End If
    
    GF_DBOpen = True
    
    Exit Function

'�G���[�������A�I���N�����ŃG���[����Ԃ���ꍇ�͂�����Q�Ƃ���
'�������L���Ŗ����iVB ErrorNumber=440�ASet OraDatabase �����s
'�o���Ȃ��j�ꍇ�͔����o��
ErrHandler:
    Dim intRet     As Integer
    Dim lngErrNum  As Long      ''�װ���ް
    Dim strErrMsg  As String    ''�װү����
    Dim strErrType As String    ''�װ����
    
    strErrType = "ORACLE"
    If Err.Number = 429 Or Err.Number = 440 Then
        lngErrNum = Err.Number
        strErrMsg = "DB�ւ̐ڑ��Ɏ��s���܂����B"
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
' �@�\�� : DB�ؒf
' �@�\   : �T�[�o�[���I���N���Ƃ̃Z�b�V�����I��
' ����   :
' ���l   :
'------------------------------------------------------------------------------

    On Error GoTo ErrHandler

'2007/10/30 Added by N.Kigaku  Start -------
    'Oracle�o�C���h�p�����[�^�폜
    Call GF_RemoveAllBindParameter
'2007/10/30 Add End ------------------------

    ''�n�����������f�[�^�x�[�X�̃N���[�Y�E�ڑ�����
    Set gOraParam = Nothing
    gOraDataBase.Close
    Set gOraDataBase = Nothing
    Set gOraSession = Nothing

    If blnWOraSessionConnectFlag = True Then
        ''���O�o�͗p�n�����������f�[�^�x�[�X�̃N���[�Y�E�ڑ�����
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
'' �@�\�T�v�@:Oracle�޲��ޕϐ��S�폜
''
'' �����@�@�@:
''
'' �߂�l�@�@:
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
