Attribute VB_Name = "basOraFunc"
' @(h) OraFunc.bas  ver1.00 ( 2002/05/07 N.Kigaku )
'------------------------------------------------------------------------------
' @(s)
'�@�v���W�F�N�g���@: TLF��ۼު��
'�@���W���[�����@�@: basOraFunc
'�@�t�@�C�����@�@�@: OraFunc.bas
'�@Version�@�@�@�@: 1.00
'�@�@�\�����@�@�@�@: �I���N���̃f�[�^�x�[�X�Ɋւ��鋤�ʊ֐�
'�@�쐬�ҁ@�@�@�@�@: N.Kigaku
'�@�쐬���@�@�@�@�@: 2002/05/07
'�@���l�@�@�@�@�@�@:
'�@�C�������@�@�@�@: 2006/12/05 N.Kigaku �׸�8.1.7 Nocache�Ή� �������AReadOnly����Nocache�ɕύX
'                  : 2012/06/11 J.Yamaoka SQL������Like�G�X�P�[�v,���e�����u���ǉ�
'�@�@�@�@�@�@�@�@�@: 2015/08/21  NIC ��  LFDB�X�V SESSIONID�C��
'�@�@�@�@�@�@�@�@�@: 2017/02/08 D.Ikeda K545 CS�v���Z�X���P
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ���錾
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' �p�u���b�N�ϐ��錾
'------------------------------------------------------------------------------

'Oracle ALL_COLL_TYPES Fields
Public Type OraAllCollType
    tOwner          As String     '���
    tTypeName       As String     'Type�^����
    tCollType       As String     '�ڸ�������
    tUpperBound     As String     '�z��
    tElemTypeName   As String     '�^
    tLength         As String     '�^�̻���
End Type

''2012/06/11 J.Yamaoka Add
Private Const mstrLikeEscape As String = "\"        ''Like�w�莞�̃G�X�P�[�v����



Public Function GF_GetSYSDATE(ByRef strSysDate As String, Optional ByVal intFormatKbn As Integer = 1) As Boolean
'------------------------------------------------------------------------------
' @(f)
'�@�@�\���@: �c�a�T�[�o�̃V�X�e�����t�擾
'�@�@�\�@�@: �c�a�T�[�o�̃V�X�e�����t���I���N���o�R�Ŏ擾����
'�@�����@�@:�@ strSysDate As String     (out)  �V�X�e�����t
'�@�@�@�@�@:�@ intFormatKbn As Integer  (in)   �����敪
'                0: �Ȃ��i�I���N���̓��t�`���Ɋ�Â��j
'                1: "yyyy/mm/dd hh24:mi:ss"        (�ȗ���)
'                2: "yyyy/mm/dd"
'�@�߂�l�@:�@True = ���� / False = ���s
'�@���l�@�@:
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
       
    '�������OPT���獡��g�p����A�Ԃ��擾����
    strSQL = ""
    strSQL = strSQL & "SELECT " & strFormatSysdate & " FROM DUAL"
'    strSQL = strSQL & "SELECT SYSDATE FROM DUAL"
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oraDyna.EOF Then
        '�V�X�e�����t�̎擾�Ɏ��s���܂����B
        strMsg = GF_GetMsg("WTG027")
        Err.Raise Number:=vbObjectError, Description:=strMsg
    Else
        strSysDate = CStr(oraDyna![SYSDATE])
    End If
    Set oraDyna = Nothing
       
    GF_GetSYSDATE = True
    
    Exit Function
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_GetSYSDATE", strSQL)
End Function

Public Function GF_GetColumLength(ByRef intColumLen As Integer _
                                , ByVal strTBL_NAME As String _
                                , ByVal strCOLUM_NAME As String _
                                , Optional ByVal strOWNER As String = "LFSYS") As Boolean
'------------------------------------------------------------------------------
' @(f)
'�@�@�\���@: �t�B�[���h�T�C�Y�̎擾
'�@�@�\�@�@: ����̃e�[�u���̃t�B�[���h�T�C�Y���擾����
'�@�����@�@:�@ intColumLen As Integer   (out)  �t�B�[���h�T�C�Y
'�@�@�@�@�@:�@ strTBL_NAME As String    (in)   �e�[�u����
'�@�@�@�@�@:�@ strCOLUM_NAME As String  (in)   �t�B�[���h��
'�@�@�@�@�@:�@ strOWNER As String       (in)   ���L�ҁ@�i�ȗ���:LFSYS�j
'�@�߂�l�@:�@True = ���� / False = ���s
'�@���l�@�@:  �I���N���֐���VSIZE�Ɠ���
'               ��jSELECT VSIZE(SYNO) FROM THJMR
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
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    If oraDyna.EOF Then
        strMsg = "�e�[�u���̃t�B�[���h�T�C�Y�̎擾�Ɏ��s���܂����B"
        Err.Raise Number:=vbObjectError, Description:=strMsg
    Else
        intColumLen = CInt(oraDyna![DATA_LENGTH])
    End If
    Set oraDyna = Nothing
       
    GF_GetColumLength = True
    
    Exit Function
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_GetColumLength", strSQL)
End Function


Public Function GF_GetSessionID(ByRef lngSessionID As Long) As Boolean
'------------------------------------------------------------------------------
' @(f)
'�@�@�\���@: �Z�b�V����ID�̎擾
'�@�@�\�@�@: �Z�b�V����ID���擾����
'�@�����@�@:�@ lngSessionID As Long   (out)  �Z�b�V����ID
'�@�߂�l�@:�@True = ���� / False = ���s
'�@���l�@�@:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL            As String
    Dim oraDyna           As OraDynaset
    Dim strMsg            As String
    
    GF_GetSessionID = False
       
    lngSessionID = 0
    '<LFDB�X�V SESSIONID�C��> del Start NIC ��  2015/08/21
    'strSQL = ""
    'strSQL = strSQL & "SELECT USERENV('SESSIONID') SESSIONID FROM DUAL"
    ''�޲ž�Ă̐���
    'Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    'If oraDyna.EOF Then
    '    strMsg = "�Z�b�V����ID�̎擾�Ɏ��s���܂����B"
    '    Err.Raise Number:=vbObjectError, Description:=strMsg
    'Else
    '    lngSessionID = CLng(oraDyna![SESSIONID])
    'End If
    'Set oraDyna = Nothing
    '<LFDB�X�V SESSIONID�C��> del end NIC ��  2015/08/21
    '<LFDB�X�V SESSIONID�C��> ADD Start NIC ��  2015/08/21
    If GF_GetSessionID_Func(lngSessionID, strMsg) = False Then
        strMsg = "�Z�b�V����ID�̎擾�Ɏ��s���܂����B"
        Err.Raise Number:=vbObjectError, Description:=strMsg
        Exit Function
        End If
    '<LFDB�X�V SESSIONID�C��> ADD end NIC ��  2015/08/21
    GF_GetSessionID = True
    
    Exit Function
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_GetSessionID", strSQL)
End Function

Public Sub Init_OraAllCollType(typCollType As OraAllCollType)
'------------------------------------------------------------------------------
' @(f)
'�@�@�\���@: ALL_COLL_TYPES�\���̂̏�����
'�@�@�\�@�@:
'�@�����@�@: typCollType As OraAllCollType   (in/out)  ALL_COLL_TYPES ð��ق�̨����
'�@�߂�l�@: �Ȃ�
'�@���l�@�@:
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
'�@�@�\���@: Type�^(ALL_COLL_TYPES)���擾
'�@�@�\�@�@: Type�^(ALL_COLL_TYPES)�̓��e���擾����
'�@�����@�@: strTypeName As String            (in)  Type�^����
'            typCollType As OraAllCollType   (out)  ALL_COLL_TYPES ð��ق�̨����
'�@�߂�l�@:�@True = ���� / False = ���s
'�@���l�@�@:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim strSQL    As String
    Dim oraDyna   As OraDynaset
    
    GF_GetAllCollType = False
    
    '�\���̂̏�����
    Call Init_OraAllCollType(typCollType)
    
    '�^���̂���̎��͐���I���Ƃ���
    If Len(Trim(strTypeName)) = 0 Then
        GF_GetAllCollType = True
        Exit Function
    End If
    
    'Type�^�̏����擾����
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
            '���
            .tOwner = IIf(IsNull(oraDyna![Owner]), "", RTrim(oraDyna![Owner]))
            'Type�^����
            .tTypeName = IIf(IsNull(oraDyna![TYPE_NAME]), "", RTrim(oraDyna![TYPE_NAME]))
            '�ڸ�������
            .tCollType = IIf(IsNull(oraDyna![COLL_TYPE]), "", RTrim(oraDyna![COLL_TYPE]))
            '�z��
            .tUpperBound = IIf(IsNull(oraDyna![UPPER_BOUND]), "", RTrim(oraDyna![UPPER_BOUND]))
            '�^
            .tElemTypeName = IIf(IsNull(oraDyna![ELEM_TYPE_NAME]), "", RTrim(oraDyna![ELEM_TYPE_NAME]))
            '�^�̻���
            .tLength = IIf(IsNull(oraDyna![length]), "", RTrim(oraDyna![length]))
        End With
    End If
    Set oraDyna = Nothing
    
    GF_GetAllCollType = True
    
    Exit Function
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_GetAllCollType", strSQL)
    
End Function

''2012/06/11 J.Yamaoka Add
Public Function GF_ReplaceSQLLikeEscape(ByVal strCondition As String) As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@: SQL������Like�G�X�P�[�v,���e�����u��
' �@�\�@�@�@:
' �����@�@�@: strCondition As String     ''�u���Ώە�����
' �߂�l�@�@: �u������������
' �@�\�����@:
'------------------------------------------------------------------------------
    strCondition = Replace(strCondition, mstrLikeEscape, String(2, mstrLikeEscape))
    strCondition = Replace(strCondition, "%", mstrLikeEscape & "%", , , vbBinaryCompare)
' 2017/02/08 �� D.Ikeda K545 CS�v���Z�X���P  DEL
'    strCondition = Replace(strCondition, "��", mstrLikeEscape & "��", , , vbBinaryCompare)
' 2017/02/08 �� D.Ikeda K545 CS�v���Z�X���P  DEL
    strCondition = Replace(strCondition, "_", mstrLikeEscape & "_", , , vbBinaryCompare)
' 2017/02/08 �� D.Ikeda K545 CS�v���Z�X���P  DEL
'    strCondition = Replace(strCondition, "�Q", mstrLikeEscape & "�Q", , , vbBinaryCompare)
' 2017/02/08 �� D.Ikeda K545 CS�v���Z�X���P  DEL
    GF_ReplaceSQLLikeEscape = strCondition
End Function
''2012/06/11 J.Yamaoka Add
Public Property Get SQLLikeEscape() As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:Like�G�X�P�[�v����
' �@�\�@�@�@:
' �����@�@�@:�Ȃ�
' �߂�l�@�@:Like�G�X�P�[�v����
' �@�\�����@:
'------------------------------------------------------------------------------
    SQLLikeEscape = mstrLikeEscape
End Property

'<LFDB�X�V SESSIONID�C��> ADD Start NIC ��  2015/08/21
Public Function GF_GetSessionID_Func(plngSessionID As Long, pstrErrMsg As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:�V�[�P���X�擾�֐��Ăяo��
' �@�\�@�@�@:
' �����@�@�@:plngSessionID As Long  �Z�b�V����ID
' �@�@�@�@�@:pstrErrMsg As String       �G���[���b�Z�[�W
' �@�\�����@:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Dim strSQL              As String
    Dim lclsOraClass        As New clsOraClass  ''Oracle�֘A�p�N���X
    Dim strErrMsg           As String
    
    GF_GetSessionID_Func = False
        
    '�X�g�A�h�p�I�u�W�F�N�g�錾
    Set lclsOraClass = New clsOraClass
    Set lclsOraClass.OraDataBase_Strcall = gOraDataBase
    '�T�[�o�[�G���[�̃��Z�b�g
    Call lclsOraClass.ErrReset_Strcall
    '�o�C���h�ϐ��ǉ�
    lclsOraClass.Add_Binds ORAPARM_OUTPUT, ORATYPE_NUMBER, "SESSION", plngSessionID       '�Z�b�V����ID
    lclsOraClass.Add_Binds ORAPARM_OUTPUT, ORATYPE_VARCHAR2, "OUTMSG", pstrErrMsg      '���ʃG���[���b�Z�[�W
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

    '�T�[�o�[�G���[�̃��Z�b�g
    Call lclsOraClass.ErrReset_Strcall
    'SQL���s
    Call lclsOraClass.ExecSql_Strcall(strSQL)

    '�`�F�b�N���ʃ��b�Z�[�W�擾
    strErrMsg = GF_VarToStr(gOraParam!OUTMSG)

    If (gOraParam!sql_code = -1) Then
        '�V�X�e���ُ�
        GF_GetSessionID_Func = False
        '�������ݎ��s
        '���O�o��
        Call GF_GetMsg_Addition("WTK009", , False, True)
        Exit Function
    Else
        plngSessionID = GF_VarToStr(gOraParam!Session)
        GF_GetSessionID_Func = True
    End If

    '�p�����[�^�[�̑S���
    lclsOraClass.RemoveAll

    Set lclsOraClass = Nothing
    
    Exit Function

ErrHandler:
    '�װ�����
    Call GS_ErrorHandler("GF_GetSessionID_Func", strSQL)
    
End Function
'<LFDB�X�V SESSIONID�C��> ADD end NIC ��  2015/08/21


