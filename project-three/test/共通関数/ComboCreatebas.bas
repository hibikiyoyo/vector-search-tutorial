Attribute VB_Name = "basComboCreate"
' @(h)  ComboCreate.BAS           ver 1.0    ( 2000/12/04 T.Fukutani )
'------------------------------------------------------------------------------
' @(s)
'   �v���W�F�N�g��  :   TLF��ۼު��
'   ���W���[����    :   basComboCreate
'   �t�@�C����      :   ComboCreate.BAS
'   �u������������  :   1.00
'   �@�\����        :   �e��R���{�{�b�N�X�̍쐬�֐�
'   �쐬��          :   T.Fukutani
'   �쐬��          :   2000/12/04
'   �C������        :   2001/04/29 N.Kigaku �̔��X�����ޯ���ؽ��ޯ���쐬�֐��Ɉ���1�ǉ�
'�@ �@�@�@�@�@�@�@�@�F  2001/11/26 N.Kigaku GF_CreateEigyoCombo,GF_MatchCombo��ǉ�
'�@ �@�@�@�@�@�@�@�@�F  2001/12/06 N.Kigaku GF_Com_CtlAdditem�C��
'�@ �@�@�@�@�@�@�@�@�F  2001/12/12 N.Kigaku GF_Com_CtlAdditem2,GF_CreateCifCombo2 �ǉ�
'�@ �@�@�@�@�@�@�@�@�F  2002/01/09 N.Kigaku GF_CreateCifCombo2�Ɉ���2�ǉ�,GF_CreateMastCombo�ǉ�
'�@ �@�@�@�@�@�@�@�@�F  2002/01/24 N.Kigaku GF_CreateSyasyuCombo�ǉ�
'�@ �@�@�@�@�@�@�@�@�F  2002/10/10 GF_CreateBunruiCombo��GF_CreateBunruiCombo2�̕���1,2,3���ލ쐬SQL���ɸ�ٰ�߉���ǉ�
'�@ �@�@�@�@�@�@�@�@�F  2005/09/19 N.Kigaku GF_CreateDistCombo�ǉ�
'�@ �@�@�@�@�@�@�@�@�F  2005/10/25 N.Kigaku GF_CreateDistCombo�C��,GF_CreateGroupCombo�ǉ�
'�@ �@�@�@�@�@�@�@�@�F  2005/11/02 N.Kigaku GF_CreateGroupCombo�C��
'�@ �@�@�@�@�@�@�@�@�F  2005/11/04 N.Kigaku GF_CreateDistCombo�C��
'�@ �@�@�@�@�@�@�@�@�F  2005/11/16 N.Kigaku GF_CreateDistCombo�C��
'                   �F  2006/02/10 N.Kigaku GF_CreateGroupCombo �啶���Ō�������悤�ɏC��
'                   �F  2006/12/05 N.Kigaku �׸�8.1.7 Nocache�Ή� �������AReadOnly����Nocache�ɕύX
'                   �F  2016/12/15 M.Tanaka K545 CS�v���Z�X���P GF_CreateGroupList�ǉ�
'                   �F  2018/05/07 M.Kawamura K545 CS�v���Z�X���P
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ���錾
'------------------------------------------------------------------------------
Option Explicit

' 2016/12/15 �� M.Tanaka K545 CS�v���Z�X���P  �ǉ�
'------------------------------------------------------------------------------
' �p�u���b�N�萔�錾
'------------------------------------------------------------------------------
'�@��S���O���[�v���X�g�{�b�N�X�쐬�p�񋓌^�̐錾
Public Enum CGL_KaitoFlg     '�񓚕����t���O
    CGL_InquiryKaitoFlg = 1  '�����񓚕����t���O
    CGL_DeliveryKaitoFlg = 2 '�����[���񓚕����t���O
    CGL_EDFKaitoFlg = 3      '�d���񓚕����t���O
' 2018/05/07 �� M.Kawamura K545 CS�v���Z�X���P
    CGL_DeliEDFKaitoFlg = 4  '�����[���E�d���񓚕����t���O
' 2018/05/07 �� M.Kawamura K545 CS�v���Z�X���P
End Enum
Public Enum CGL_HonkiAttKbn  '�{�@ATT�敪
    CGL_HonkiAttAll = 0      '�{�@ATT�敪�̏��������Ȃ�
    CGL_Att = 1              'ATT
    CGL_Honki = 2            '�{�@
    CGL_Sonota = 3           '���̑�
End Enum
Public Enum CGL_IdArrayKbn   'ID�p�z����e�敪
    CGL_Id = 1               '�O���[�vID�̂�
    CGL_IdAndModelKaito = 2  '�O���[�vID & ',' & �@��}�X�^�񓚕����敪
End Enum
Public Enum CGL_EigyoDispKbn '�c�ƕ\���敪
    CGL_EigyoAll = 0         '�����A�C�O�̏��������Ȃ�
    CGL_Kokunai = 1          '����
    CGL_Kaigai = 2           '�C�O
End Enum
' 2016/12/15 �� M.Tanaka K545 CS�v���Z�X���P  �ǉ�

Public Function GF_Com_CtlAdditem(cntControl As Control, strCombo As String, Optional intIndex As Integer = 0, _
            Optional intOption As Integer = 0, Optional bolCDFlg As Boolean = True) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :   �����ޯ��Ͻ������������ޯ�����쐬����
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               strCombo   As String     �R���{ or ���X�g�{�b�N�X�̖��O
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
'               bolCDFlg   As Boolean    �R�[�h�̕\��(�ȗ���)/��\��    (True:�\���AFalse:��\��)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String    '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    
    GF_Com_CtlAdditem = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    strComboName = Trim(strCombo)
    
    '''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT CDVAL,"
    strSQL = strSQL & "       CDNAME"
    strSQL = strSQL & "   FROM THJCMBXMR"
    strSQL = strSQL & "   WHERE CMBNAME = '" & strComboName & "'"
    strSQL = strSQL & "   ORDER BY SEQNO"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![CDVAL])
            strName = GF_VarToStr(oraDyna![CDNAME])
            
            '�R�[�h�\��/��\��
            If bolCDFlg = True Then
                '�\��
                cntControl.AddItem strCode & "�F" & strName
            Else
                '��\��
                cntControl.AddItem strName
            End If
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_Com_CtlAdditem = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
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
' �@�\��    :   �R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :   �����ޯ��Ͻ������������ޯ�����쐬����
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               strCombo   As String     �R���{ or ���X�g�{�b�N�X�̖��O
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
'               intHyojiKbn as Integer   �\�����e�敪 (�ȗ��� ����:����)
'                  1 = ����:����
'                  2 = ���� ��߰� :����
'                  3 = ����
'               intSpace  As Integer     ���̂ƺ��ނƂ̊Ԋu(�ȗ���0)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String    '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    
    GF_Com_CtlAdditem2 = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    strComboName = Trim(strCombo)
    
    '''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT CDVAL,"
    strSQL = strSQL & "       CDNAME"
    strSQL = strSQL & "   FROM THJCMBXMR"
    strSQL = strSQL & "   WHERE CMBNAME = '" & strComboName & "'"
    strSQL = strSQL & "   ORDER BY SEQNO"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![CDVAL])
            strName = GF_VarToStr(oraDyna![CDNAME])
            
            If intHyojiKbn = 1 Then
            '����:����
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "�F" & strName)
                
            ElseIf intHyojiKbn = 2 Then
            '���� ��߰� :����
                cntControl.AddItem strName & Space(intSpace) & "�F" & strCode
            
            ElseIf intHyojiKbn = 3 Then
            '����
                cntControl.AddItem strName
                
            End If
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_Com_CtlAdditem2 = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_Com_CtlAdditem2", strSQL)

End Function

Public Function GF_Com_CboGetText(cntControl As Control) As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �R���{�{�b�N�X�E���X�g�{�b�N�X�̃e�L�X�g�؂�o��
' �@�\      :   �R���{�{�b�N�X�ɐݒ肳��Ă���e�L�X�g�̕������i�F���E�j�����o��
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
' �߂�l    :   String      �����o����������
' ���l      :   �R���{���I���̏ꍇ��A�󔒃e�L�X�g��I���̏ꍇ�� Null ��ԋp
'------------------------------------------------------------------------------
    GF_Com_CboGetText = ""
    
    If cntControl.ListIndex >= 0 Then
        If (InStr(1, cntControl.Text, "�F") - 1) > 0 Then
            GF_Com_CboGetText = Right(cntControl.Text, Len(cntControl.Text) - InStr(1, cntControl.Text, "�F"))
        End If
    End If
    
End Function

Public Function GF_Com_CboGetCode(cntControl As Control) As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �R���{�{�b�N�X�E���X�g�{�b�N�X�̃e�L�X�g�؂�o��
' �@�\      :   �R���{�{�b�N�X�ɐݒ肳��Ă���e�L�X�g�̕������i�F��荶�j�����o��
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
' �߂�l    :   String      �����o����������
' ���l      :   �R���{���I���̏ꍇ��A�󔒃e�L�X�g��I���̏ꍇ�� Null ��ԋp
'------------------------------------------------------------------------------
    GF_Com_CboGetCode = ""
    
    If cntControl.ListIndex >= 0 Then
        If (InStr(1, cntControl.Text, "�F") - 1) > 0 Then
            GF_Com_CboGetCode = Left(cntControl.Text, InStr(1, cntControl.Text, "�F") - 1)
        End If
    End If
    
End Function

Public Function GF_CreateTantoCombo(cntControl As Control, Optional intPostFlg As Integer, _
        Optional intIndex As Integer = 0, Optional intOption As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �S���҃R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :   �Ј�Ͻ��������ޯ�����쐬����
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               intPostFlg As Integer    �����敪(1�FCS�݌v�A2:�J���l���A3:�c�ƁA4:����)
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String    '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateTantoCombo = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL��
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
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![SYAINCD])
            strName = GF_VarToStr(oraDyna![Name])
            cntControl.AddItem strCode & "�F" & strName
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateTantoCombo = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_CreateTantoCombo", strSQL)

End Function

Public Function GF_CreateCifCombo(cntControl As Control, Optional intIndex As Integer = 0, _
                                  Optional intOption As Integer = 0, Optional bolItemFlg As Boolean = False) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �̔��X�R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :   �̔��XϽ��������ޯ�����쐬����
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
'               bolItemFlg As Boolea     ItemData�ɔ̔��X���ނ�ݒ肷�邩�ۂ�(�ȗ�����)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String    '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateCifCombo = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT CIFNO,"
    strSQL = strSQL & "       CIFNAME"
    strSQL = strSQL & "   FROM THJCIF"
    strSQL = strSQL & "   GROUP BY CIFNO,CIFNAME"
    strSQL = strSQL & "   ORDER BY CIFNO"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![CIFNO])
            strName = GF_VarToStr(oraDyna![CIFNAME])
            cntControl.AddItem strCode & "�F" & strName
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
    ''�װ�����
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
' �@�\��    :   �c�Ə��R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :   �̔��XϽ��������ޯ�����쐬����
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               strCifNO   As String     �̔��X�R�[�h(�ȗ�����)
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
'               bolItemFlg As Boolea     ItemData�ɉc�Ə����ނ�ݒ肷�邩�ۂ�(�ȗ�����)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String    '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateEigyoCombo = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT EIGYONO,"
    strSQL = strSQL & "       EIGYONAME"
    strSQL = strSQL & "   FROM THJCIF"
    If Len(Trim(strCifNO)) > 0 Then
        strSQL = strSQL & "   WHERE CIFNO='" & strCifNO & "'"
    End If
    strSQL = strSQL & "   ORDER BY CIFNO"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![EIGYONO])
            strName = GF_VarToStr(oraDyna![EIGYONAME])
            cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "�F" & strName)
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
    ''�װ�����
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
' �@�\��    :   �c�Ə��R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :   �̔��XϽ��������ޯ�����쐬����
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               strCifNO   As String     �̔��X�R�[�h(�ȗ�����)
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
'               bolItemFlg As Boolea     ItemData�ɉc�Ə����ނ�ݒ肷�邩�ۂ�(�ȗ�����)
'               intHyojiKbn as Integer     �\�����e�敪 (�ȗ��� ����:����)
'                  1 = ����:����
'                  2 = ���� ��߰� :����
'                  3 = ����
'               intSpace  As Integer       ���̂ƺ��ނƂ̊Ԋu(�ȗ���0)
'               blnDispNameFlg As Boolean  ���̂��������ɒǉ����邩�ۂ�(�ȗ����ǉ�)   False:�ǉ����Ȃ��ATrue:�ǉ�
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String    '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateEigyoCombo2 = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT EIGYONO,"
    strSQL = strSQL & "       EIGYONAME"
    strSQL = strSQL & "   FROM THJCIF"
    If Len(Trim(strCifNO)) > 0 Then
        strSQL = strSQL & "   WHERE CIFNO='" & strCifNO & "'"
    End If
    strSQL = strSQL & "   ORDER BY CIFNO"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![EIGYONO])
            strName = GF_VarToStr(oraDyna![EIGYONAME])
            If intHyojiKbn = 1 Then
            '����:����
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "�F" & strName)
                End If
                
            ElseIf intHyojiKbn = 2 Then
            '���� ��߰� :����
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "�F" & strCode)
                End If
                
            ElseIf intHyojiKbn = 3 Then
            '����
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
    ''�װ�����
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
' �@�\��    :   �̔��X�^�c�Ə��R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬�Q
' �@�\      :   �̔��XϽ��������ޯ�����쐬����
' ����      :   cntControl As Control      �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               intCifKbn As Integer       �̔��X�E�c�Ə��쐬�敪
'                  1 = �̔��X
'                  2 = �c�Ə�
'               strCifNO As String         �̔��X����(�̔��X�E�c�Ə��쐬�敪���c�Ə��̎��̂�)
'               intIndex   As Integer      �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer      �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
'               intHyojiKbn as Integer     �\�����e�敪 (�ȗ��� ����:����)
'                  1 = ����:����
'                  2 = ���� ��߰� :����
'                  3 = ����
'               intSpace  As Integer       ���̂ƺ��ނƂ̊Ԋu(�ȗ���0)
'               blnDispNameFlg As Boolean  ���̂��������ɒǉ����邩�ۂ�(�ȗ����ǉ�)   False:�ǉ����Ȃ��ATrue:�ǉ�
'               intShiyuKbn As Integer     �s�A�敪(�ȗ��� ����)
'                  0 = �����ƊC�O
'                  1 = �����̂�
'                  2 = �C�O�̂�
'               intDispCS As Integer       �\���敪(�ȗ���:0)  0:�@��A1:���@��A2:�S��
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim strSQL_Srh   As String      '��������SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String    '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreateCifCombo2 = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL��
    strSQL = ""
    strSQL_Srh = ""
    
    ''�s�A�敪�����쐬
    Select Case intShiyuKbn
        Case 1
            strSQL_Srh = " SHIYUKBN = '1'"
        Case 2
            strSQL_Srh = " SHIYUKBN = '2'"
        Case Else
            strSQL_Srh = ""
    End Select
    
    ''�\���敪�����쐬
    Select Case intDispCS
        Case 0, 1
        
            If Len(strSQL_Srh) > 0 Then
                strSQL_Srh = strSQL_Srh & " AND "
            End If
            If intDispCS = 0 Then
                '�@��
                strSQL_Srh = " CORDER_DISP_CS = '1'"
            ElseIf intDispCS = 1 Then
                '���@��
                strSQL_Srh = " CENV_CARRY_CS = '1'"
            End If
        
        Case Else
    End Select
    
    If intCifKbn = 1 Then
        '�̔��X�����ޯ���ؽ��ޯ���̍쐬
        strSQL = strSQL & "SELECT CIFNO C_NO,"
        strSQL = strSQL & "       CIFNAME C_NAME"
        strSQL = strSQL & "  FROM THJCIF"
         strSQL = strSQL & IIf(Len(strSQL_Srh) = 0, "", " WHERE" & strSQL_Srh)
        strSQL = strSQL & " GROUP BY CIFNO,CIFNAME"
        strSQL = strSQL & " ORDER BY CIFNO"
    ElseIf intCifKbn = 2 Then
        '�c�Ə������ޯ���ؽ��ޯ���̍쐬
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
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![C_NO])
            strName = GF_VarToStr(oraDyna![C_NAME])
            If intHyojiKbn = 1 Then
            '����:����
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "�F" & strName)
                End If
                
            ElseIf intHyojiKbn = 2 Then
            '���� ��߰� :����
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "�F" & strCode)
                End If
'                cntControl.AddItem strName & Space(intSpace) & "�F" & strCode
                
            ElseIf intHyojiKbn = 3 Then
            '����
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
    ''�װ�����
    Call GS_ErrorHandler("GF_CreateCifCombo2", strSQL)

End Function

Public Function GF_CreateBunruiCombo(cntControl As Control, intBunruiFlg As Integer, _
                            Optional strBunrui1 As String, Optional strBunrui2 As String, _
                            Optional intOption As Integer = 1) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:�@ ���޺����ޯ���̍쐬
' �@�\�@�@�@:�@ C-OPT����ð��ق������ޯ�����쐬����B
' �����@�@�@:�@ cntControl     As Control    �ΏۂƂȂ���޺��۰�
' �@�@�@�@�@ �@ intBunruiFlg   As Integer    ���ދ敪(1�F���ނP, 2:���ނQ, 3:���ނR)
' �@�@�@�@�@ �@ strBunrui1�@   AS String     ���޺���1
' �@�@�@�@�@ �@ strBunrui2�@   AS String     ���޺���2
' �@�@�@�@�@ �@ intOption      As Integer    �I�����ڂ̐擪���ٍ��ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ�������)
' �߂�l�@�@:�@ True = ���� / False = ���s
' ���l�@�@�@:
'------------------------------------------------------------------------------
    Dim strSQL       As String      'SQL��
    Dim oDynaset     As OraDynaset  '�޲ž��
    Dim strComboName As String      '���۰ٖ�
    Dim strCode      As String
    Dim strCode1     As String * 4
    Dim strCode2     As String * 4
    Dim strName      As String
    On Error GoTo ErrHandler
    GF_CreateBunruiCombo = False
    
'   ���۰ُ�����
    cntControl.Clear
    cntControl.ListIndex = -1
    
'   SQL��
    strSQL = ""
    Select Case intBunruiFlg
        Case 1
            strSQL = strSQL & " SELECT BUNRUI1     BUNRUI,"
            strSQL = strSQL & "        BUNRUINAME1 BUNRUINAME"
            strSQL = strSQL & "�@ FROM THJBUNRUI1"
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
'   �Y��ں��ނ��Ȃ��ꍇ
    If oDynaset.EOF = True Then
        Exit Function
    End If
'   ���ލ쐬
'   �ْl�ݒ�
    If intOption <> 0 Then
        cntControl.AddItem ""
    End If
    Do Until oDynaset.EOF = True
        strCode = GF_VarToStr(oDynaset![BUNRUI])
        strName = GF_VarToStr(oDynaset![BUNRUINAME])
        cntControl.AddItem Trim(strCode) & "�F" & Trim(strName)
        oDynaset.MoveNext
    Loop
    
    GF_CreateBunruiCombo = True
    Exit Function
    
ErrHandler:
'   �װ�����
    Call GS_ErrorHandler("GF_CreateBunruiCombo", strSQL)

End Function

Public Function GF_CreateBunruiCombo2(cntControl As Control, intBunruiFlg As Integer, _
                            strKatashiki As String, Optional strBunrui1 As String, _
                            Optional strBunrui2 As String, Optional intOption As Integer = 1 _
                            ) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:�@ ���޺����ޯ���̍쐬
' �@�\�@�@�@:�@ �@��^����C-OPTϽ����������A��v���镪�ޖ��̂�C-OPT����ð��ق��
'              �擾���ĺ����ޯ�����쐬����B
' �����@�@�@:�@ cntControl     As Control    �ΏۂƂȂ���޺��۰�
' �@�@�@�@�@ �@ intBunruiFlg   As Integer    ���ދ敪(1�F���ނP, 2:���ނQ, 3:���ނR)]
'              strKatashiki   As String     �@��^��
' �@�@�@�@�@ �@ strBunrui1�@   AS String     ���޺���1
' �@�@�@�@�@ �@ strBunrui2�@   AS String     ���޺���2
' �@�@�@�@�@ �@ intOption      As Integer    �I�����ڂ̐擪���ٍ��ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ�������)
' �߂�l�@�@:�@ True = ���� / False = ���s
' ���l�@�@�@:
'------------------------------------------------------------------------------
    Dim strSQL       As String      'SQL��
    Dim oDynaset     As OraDynaset  '�޲ž��
    Dim strComboName As String      '���۰ٖ�
    Dim strCode      As String
    Dim strCode1     As String * 4
    Dim strCode2     As String * 4
    Dim strName      As String
    On Error GoTo ErrHandler
    GF_CreateBunruiCombo2 = False
    
'   ���۰ُ�����
    cntControl.Clear
    cntControl.ListIndex = -1
    
'   SQL��
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
'   �Y��ں��ނ��Ȃ��ꍇ
    If oDynaset.EOF = True Then
        Exit Function
    End If
'   ���ލ쐬
'   �ْl�ݒ�
    If intOption <> 0 Then
        cntControl.AddItem ""
    End If
    Do Until oDynaset.EOF = True
        strCode = GF_VarToStr(oDynaset![BUNRUI])
        strName = GF_VarToStr(oDynaset![BUNRUINAME])
        cntControl.AddItem Trim(strCode) & "�F" & Trim(strName)
        oDynaset.MoveNext
    Loop
    
    GF_CreateBunruiCombo2 = True
    Exit Function
    
ErrHandler:
'   �װ�����
    Call GS_ErrorHandler("GF_CreateBunruiCombo2", strSQL)

End Function

Public Function GF_CreateBunruiCombo2_2(cntControl As Control, intBunruiFlg As Integer, _
                            strKatashiki As String, Optional strBunrui1 As String, _
                            Optional strBunrui2 As String, Optional intOption As Integer = 1 _
                            ) As Boolean
'------------------------------------------------------------------------------
' @(f)
' �@�\���@�@:�@ ���޺����ޯ���̍쐬
' �@�\�@�@�@:�@ �@��^����C-OPTϽ�2���������A��v���镪�ޖ��̂�C-OPT����ð��ق��
'              �擾���ĺ����ޯ�����쐬����B
' �����@�@�@:�@ cntControl     As Control    �ΏۂƂȂ���޺��۰�
' �@�@�@�@�@ �@ intBunruiFlg   As Integer    ���ދ敪(1�F���ނP, 2:���ނQ, 3:���ނR)]
'              strKatashiki   As String     �@��^��
' �@�@�@�@�@ �@ strBunrui1�@   AS String     ���޺���1
' �@�@�@�@�@ �@ strBunrui2�@   AS String     ���޺���2
' �@�@�@�@�@ �@ intOption      As Integer    �I�����ڂ̐擪���ٍ��ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ�������)
' �߂�l�@�@:�@ True = ���� / False = ���s
' ���l�@�@�@:
'------------------------------------------------------------------------------
    Dim strSQL       As String      'SQL��
    Dim oDynaset     As OraDynaset  '�޲ž��
    Dim strComboName As String      '���۰ٖ�
    Dim strCode      As String
    Dim strCode1     As String * 4
    Dim strCode2     As String * 4
    Dim strName      As String
    On Error GoTo ErrHandler
    GF_CreateBunruiCombo2_2 = False
    
'   ���۰ُ�����
    cntControl.Clear
    cntControl.ListIndex = -1
    
'   SQL��
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
'   �Y��ں��ނ��Ȃ��ꍇ
    If oDynaset.EOF = True Then
        Exit Function
    End If
'   ���ލ쐬
'   �ْl�ݒ�
    If intOption <> 0 Then
        cntControl.AddItem ""
    End If
    Do Until oDynaset.EOF = True
        strCode = GF_VarToStr(oDynaset![BUNRUI])
        strName = GF_VarToStr(oDynaset![BUNRUINAME])
        cntControl.AddItem Trim(strCode) & "�F" & Trim(strName)
        oDynaset.MoveNext
    Loop
    
    GF_CreateBunruiCombo2_2 = True
    Exit Function
    
ErrHandler:
'   �װ�����
    Call GS_ErrorHandler("GF_CreateBunruiCombo2_2", strSQL)

End Function

Public Function GF_MatchCombo(cntControl As Control, strCheck As String _
                            , Optional blnSpaceCheck As Boolean = False _
                            , Optional intLRCheckCode As Integer = 1 _
                            ) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �R���{�{�b�N�X�E���X�g�{�b�N�X�̕\�����e�ݒ�
' �@�\      :   �R���{�{�b�N�X�E���X�g�{�b�N�X�ň�v����e�L�X�g�ɐݒ肷��
' ����      :   cntControl As Control   �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               strCheck As String      �����Ώۃf�[�^
'               blnSpaceCheck As Boolean �󔒎��ɓ���`�F�b�N���s�����ǂ���
'                                         True = ��������A False = �������Ȃ�
'               intLRCheckCode As Integer ���E�ǂ���̺��ނ����o����
'                                         1 = �� �A2 = �E
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    Dim intloop As Integer
    Dim strCode As String
    
    On Error GoTo ErrHandler
    
    GF_MatchCombo = False
    
    '�R���g���[��������
    cntControl.ListIndex = -1
    
    For intloop = 0 To cntControl.ListCount - 1
        If (InStr(1, cntControl.List(intloop), "�F") - 1) > 0 Then
            If intLRCheckCode = 1 Then
                '�F�̍����̃R�[�h�����o��
                strCode = Left(cntControl.List(intloop), InStr(1, cntControl.List(intloop), "�F", vbTextCompare) - 1)
            Else
                '�F�̉E���̃R�[�h�����o��
                strCode = Mid(cntControl.List(intloop), InStrRev(cntControl.List(intloop), "�F", -1, vbTextCompare) + 1)
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
    ''�װ�����
    Call GS_ErrorHandler("GF_MatchCombo")

End Function


Public Function GF_SetCifCombo(cntControl As Control, strCheck As String) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �̔��X�R���{�{�b�N�X�E���X�g�{�b�N�X�̕\�����e�ݒ�
' �@�\      :   �̔��X�R���{�{�b�N�X�E���X�g�{�b�N�X�̕\�����e��ݒ肷��
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               strCheck As String      �����Ώۃf�[�^
' �߂�l    :   True = ���� / False = ���s
' ���l      :   ItemData��ϯ�ݸޑΏۂƂ��邽�߁A���炩����ItemData���ް������Ă���
'               ���l�^�̂�(������s��)
'------------------------------------------------------------------------------
    Dim intloop As Integer
    
    On Error GoTo ErrHandler
    
    GF_SetCifCombo = False
    
    '�R���g���[��������
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
    ''�װ�����
    Call GS_ErrorHandler("GF_SetCifCombo")

End Function

Public Function GF_CreateGrpCombo(cntControl As Control, _
        Optional intIndex As Integer = 0, Optional intOption As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   GRP�R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :   C-OPTϽ��������ޯ�����쐬����
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    
    GF_CreateGrpCombo = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT GRP"
    strSQL = strSQL & " FROM THJCOPTMR"
    strSQL = strSQL & " GROUP BY GRP"
    strSQL = strSQL & " ORDER BY GRP"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If (oraDyna.EOF = True) Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If (intOption <> 0) Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
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
    ''�װ�����
    Call GS_ErrorHandler("GF_CreateGrpCombo", strSQL)
    
End Function

Public Function GF_CreateMastCombo(cntControl As Control, _
        Optional intIndex As Integer = 0, Optional intOption As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   Mast�R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    
    GF_CreateMastCombo = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT MASTTYPE"
    strSQL = strSQL & " FROM THJMSTTAIOMR"
    strSQL = strSQL & " GROUP BY MASTTYPE"
    strSQL = strSQL & " ORDER BY MASTTYPE"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If (oraDyna.EOF = True) Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If (intOption <> 0) Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
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
    ''�װ�����
    Call GS_ErrorHandler("GF_CreateMastCombo", strSQL)
    
End Function

Public Function GF_CreateSyasyuCombo(cntControl As Control, _
        Optional intIndex As Integer = 0, Optional intOption As Integer = 0) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �Ԏ�R�[�h�R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    
    GF_CreateSyasyuCombo = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL��
    strSQL = ""
    strSQL = strSQL & " SELECT SERIESCD "
    strSQL = strSQL & "   FROM THJHNKTYPEMR"
    strSQL = strSQL & "  WHERE SHIYUKBN = ' '"
    strSQL = strSQL & "  GROUP BY SERIESCD"
    strSQL = strSQL & "  ORDER BY SERIESCD "
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If (oraDyna.EOF = True) Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If (intOption <> 0) Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
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
    ''�װ�����
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
' �@�\��    :   ���}�X�^�[�R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :   ���}�X�^�[�������ޯ�����쐬����
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
'               bolItemFlg As Boolea     ItemData�Ɍ����ނ�ݒ肷�邩�ۂ�(�ȗ�����)
'               intHyojiKbn as Integer     �\�����e�敪 (�ȗ��� ����:����)
'                  1 = ����:����
'                  2 = ���� ��߰� :����
'                  3 = ����
'               intSpace  As Integer       ���̂ƺ��ނƂ̊Ԋu(�ȗ���0)
'               blnDispNameFlg As Boolean  ���̂��������ɒǉ����邩�ۂ�(�ȗ����ǉ�)   False:�ǉ����Ȃ��ATrue:�ǉ�
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String    '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    
    GF_CreatePRFCombo = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    
    '''SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT CPREFECTURE_CD,"
    strSQL = strSQL & "       VCPREFECTURE_NAME"
    strSQL = strSQL & "  FROM M23_PREFECTURE"
    strSQL = strSQL & " ORDER BY NLIST_SEQ"
    
    '�޲ž�Ă̐���
    Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![CPREFECTURE_CD])
            strName = GF_VarToStr(oraDyna![VCPREFECTURE_NAME])
            If intHyojiKbn = 1 Then
            '����:����
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "�F" & strName)
                End If
                
            ElseIf intHyojiKbn = 2 Then
            '���� ��߰� :����
                If (blnDispNameFlg = True) Or (Len(Trim(strName)) > 0) Then
                    cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "�F" & strCode)
                End If
                
            ElseIf intHyojiKbn = 3 Then
            '����
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
    ''�װ�����
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
' �@�\��    :   �d����R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :   �����Ͻ��������ޯ�����쐬����(��ٰ�ߺ���(M30_GROUP)�w��\)
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               vntGroupCD As Variant    �O���[�v�R�[�h(M30_GROUP�ɑΉ�)�i�Y����0����J�n�j
'                                        1���̎��͔z��łȂ��Ă�OK
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
'               intHyojiKbn as Integer     �\�����e�敪 (�ȗ��� ����:����)
'                  1 = ����:����
'                  2 = ���� ��߰� :����
'                  3 = ����
'               intSpace  As Integer       ���̂ƺ��ނƂ̊Ԋu(�ȗ���0)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String      '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    Dim intMsgCount As Integer
    Dim i           As Integer
    Dim strWHERE    As String
        
    GF_CreateDistCombo = False
    
    
    '' �z��̐��𐔂���
    If IsArray(vntGroupCD) = True Then
        intMsgCount = UBound(vntGroupCD) + 1
    Else
        intMsgCount = 0
    End If

    strWHERE = ""
    
    '���������쐬
    If intMsgCount > 0 Then
        '�z��
        For i = 0 To intMsgCount - 1
            If Trim(strWHERE) = "" Then
                strWHERE = strWHERE & "IN ( '" & vntGroupCD(i) & "'"
            Else
                strWHERE = strWHERE & ",'" & vntGroupCD(i) & "'"
            End If
        Next
        strWHERE = strWHERE & ")"
    
    ElseIf (Len(Trim(vntGroupCD)) > 0) Then
        '�z��ȊO
        strWHERE = " = '" & vntGroupCD & "'"
    End If
    
    '�R���g���[��������
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
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![CDIST_CD])
            strName = GF_VarToStr(oraDyna![VCDIST_NAME])
            
            If intHyojiKbn = 1 Then
            '����:����
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "�F" & strName)
                
            ElseIf intHyojiKbn = 2 Then
            '���� ��߰� :����
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "�F" & strCode)
                
            ElseIf intHyojiKbn = 3 Then
            '����
                cntControl.AddItem strName
                
            End If
            
            oraDyna.MoveNext
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateDistCombo = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
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
' �@�\��    :   �O���[�v�R�[�h�R���{�{�b�N�X�E���X�g�{�b�N�X�̍쐬
' �@�\      :�@�@հ��ID�ɑΉ������ٰ�ߺ��ނ��ٰ�߽�(M30_GROUP)�������ޯ�����쐬����
'�@�@�@�@�@�@�@�@հ��ID�̌�����admin("1")�̏ꍇ�A�S��ٰ�ߑΏہB
' ����      :   cntControl As Control    �ΏۂƂȂ�R���{�R���g���[���y�у��X�g�R���g���[��
'               intIndex   As Integer    �f�t�H���g�\���C���f�b�N�X(�ȗ���0)
'               intOption  As Integer    �I�����ڂ̐擪�Ƀk�����ڂ�ǉ�����ꍇ�ɂP���w�肷��(�ȗ����Ȃ�)
'               intHyojiKbn as Integer     �\�����e�敪 (�ȗ��� ����:����)
'                  1 = ����:����
'                  2 = ���� ��߰� :����
'                  3 = ����
'               intSpace  As Integer       ���̂ƺ��ނƂ̊Ԋu(�ȗ���0)
'               strUserID As String        հ��ID
'
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim oraDyna      As OraDynaset  '�޲ž��
    Dim strComboName As String      '�R���g���[����
    Dim strCode      As String
    Dim strName      As String
    Dim strWHERE     As String
    Const strAdminFlg As String = "1"
    
    GF_CreateGroupCombo = False
    
    If Trim(strUserID) <> "" Then
    
    ''հ�ްϽ�����������A�Ǘ��׸ނ��擾
        strSQL = ""
        strSQL = strSQL & " SELECT CADMIN_FLG"
        strSQL = strSQL & " FROM M29_USER "
        strSQL = strSQL & " WHERE UPPER(CUSER_ID) = UPPER('" & Trim(strUserID) & "')"
        
        Set oraDyna = gOraDataBase.CreateDynaset(strSQL, ORADYN_NOCACHE)
       
        If oraDyna.EOF = True Then
            ''�Y���ް��Ȃ�
            Exit Function
        Else
            If GF_VarToStr(oraDyna![CADMIN_FLG]) = strAdminFlg Then
                ''�Ǘ��҂̏ꍇ�A�S��ٰ�ߑΏ�
                strWHERE = ""
            Else
                ''�Ǘ��҈ȊO�Aհ�ް��ٰ��Ͻ��o�^��ٰ�߂��Ώ�
                strWHERE = ""
                strWHERE = strWHERE & "SELECT CGROUP_CD FROM M31_GROUP_USER"
                strWHERE = strWHERE & " WHERE UPPER(CUSER_ID) = UPPER('" & Trim(strUserID) & "')"
            End If
        
        End If
    
    End If
    
    
    '�R���g���[��������
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
    
    ''�ް���������
    If oraDyna.EOF = True Then
        '''�Y���f�[�^�Ȃ�
        'ү���ޕ\��
'        intRet = GF_MsgBoxDB(Me.Caption, "WTG001", "OK", "E")
        Exit Function
    Else
        '�k�����ڐݒ肠��H
        If intOption <> 0 Then
            '�k�����ڒǉ�
            cntControl.AddItem ""
        End If
        
        '���ڐݒ�
        Do
            strCode = GF_VarToStr(oraDyna![CGROUP_CD])
            strName = GF_VarToStr(oraDyna![VCGROUP_NAME])
        
            If intHyojiKbn = 1 Then
            '����:����
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strCode & "�F" & strName)
                
            ElseIf intHyojiKbn = 2 Then
            '���� ��߰� :����
                cntControl.AddItem IIf(Len(Trim(strCode)) = 0, " ", strName & Space(intSpace) & "�F" & strCode)
                
            ElseIf intHyojiKbn = 3 Then
            '����
                cntControl.AddItem strName
                
            End If
            
            oraDyna.MoveNext
        
        Loop Until oraDyna.EOF = True
        
        cntControl.ListIndex = intIndex
    End If
    
    GF_CreateGroupCombo = True
    
    Exit Function
    
ErrHandler:
    ''�װ�����
    Call GS_ErrorHandler("GF_CreateGroupCombo", strSQL)

End Function

' 2016/12/15 �� M.Tanaka K545 CS�v���Z�X���P  �ǉ�
Public Function GF_CreateGroupList(ByRef cntControl As Control, ByRef strGroupIdArray() As String, _
        ByVal intKaitoFlg As CGL_KaitoFlg, ByVal intHonkiAttKbn As CGL_HonkiAttKbn, _
        ByVal intIdArrayKbn As CGL_IdArrayKbn, ByVal intEigyoDispKbn As CGL_EigyoDispKbn) As Boolean
'------------------------------------------------------------------------------
' @(f)
'
' �@�\��    :   �@��S���O���[�v���X�g�{�b�N�X�̍쐬
' �@�\      :
' ����      :   cntControl        As Control   �ΏۂƂȂ郊�X�g�R���g���[��
'               srtGroupIDArray() As String    �O���[�vID�p�z��  (�O���[�vID�A�@��}�X�^�񓚕����敪�i�[�p�z��)
'               strKaitoFlg       As Integer   �񓚕����t���O���(1:�����񓚕����t���O�A2:�����[���񓚕����t���O�A3:�d���񓚕����t���O�A4:�����[���E�d���񓚕����t���O)
'               strHonkiAttKbn    As Integer   �{�@ATT�敪       (0:�{�@ATT�敪�̏��������Ȃ��A1:ATT�A2:�{�@�A3:���̑�)
'               intIdArrayKbn     As Integer   ID�p�z����e�敪  (1:�O���[�vID�̂݁A2:�O���[�vID & ',' & �@��}�X�^�񓚕����敪)
'               intEigyoDispKbn   As Integer   �c�ƕ\���敪      (0:�����A�C�O�̏��������Ȃ��A1:�����A2:�C�O)
' �߂�l    :   True = ���� / False = ���s
' ���l      :
'------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Dim strSQL       As String      'SQL��
    Dim strSQL_WHERE As String      'SQL��WHERE��
    Dim oraDynaCount As OraDynaset  '�޲ž��(����)
    Dim oraDynaData  As OraDynaset  '�޲ž��(�f�[�^)
    Dim strKaigaiEigyoId  As String '�C�O�c�ƃO���[�vID
    Dim strKokunaiEigyoId As String '�����c�ƃO���[�vID
    Dim intCnt       As Integer     '����
    Dim intIndex     As Integer     '�C���f�b�N�X
    
    GF_CreateGroupList = False
    
    '�R���g���[��������
    cntControl.Clear
    cntControl.ListIndex = -1
    strKaigaiEigyoId = ""
    strKokunaiEigyoId = ""
    intCnt = 0
    intIndex = 0
    ReDim strGroupIdArray(0)
  
    'INI�t�@�C���}�X�^��荑�c�A�C�c�̃O���[�vID���擾
    '�c�ƕ\���敪
    Select Case intEigyoDispKbn
        Case CGL_EigyoAll '�����A�C�O�̏��������Ȃ��ꍇ
            '�������Ȃ�
        Case CGL_Kokunai  '�����̏ꍇ
            '�C�c�̃O���[�vID���擾
            strKaigaiEigyoId = LF_GetIniTable("KAITOEG_KAIGAI_GROUPID", 1)
            '�擾�ł��Ȃ������ꍇ
            If Len(strKaigaiEigyoId) = 0 Then
                '�G���[���b�Z�[�W�\��
                Call GF_GetMsg_Addition("WTK785", "�C�c�O���[�vID", True)
                Exit Function
            End If
        Case CGL_Kaigai   '�C�O�̏ꍇ
            '���c�̃O���[�vID���擾
            strKokunaiEigyoId = LF_GetIniTable("KAITOEG_KOKUNAI_GROUPID", 1)
            '�擾�ł��Ȃ������ꍇ
            If Len(strKokunaiEigyoId) = 0 Then
                '�G���[���b�Z�[�W�\��
                Call GF_GetMsg_Addition("WTK785", "���c�O���[�vID", True)
                Exit Function
            End If
    End Select
    
    '�������擾����
    'SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT COUNT(NGROUP_ID) Cnt"
    strSQL = strSQL & " FROM TCS_GROUP"
    
    'SQL��WHERE��
    strSQL_WHERE = ""
    strSQL_WHERE = strSQL_WHERE & " WHERE 1 = 1"
    '�񓚕����t���O
    Select Case intKaitoFlg
        Case CGL_InquiryKaitoFlg  '�����񓚕����t���O�̏ꍇ
            strSQL_WHERE = strSQL_WHERE & " AND CINQUIRY_KAITO_FLG = '1'"
        Case CGL_DeliveryKaitoFlg '�����[���񓚕����t���O�̏ꍇ
            strSQL_WHERE = strSQL_WHERE & " AND CDELIVERY_KAITO_FLG = '1'"
        Case CGL_EDFKaitoFlg      '�d���񓚕����t���O�̏ꍇ
            strSQL_WHERE = strSQL_WHERE & " AND CEDF_KAITO_FLG = '1'"
' 2018/05/07 �� M.Kawamura K545 CS�v���Z�X���P
        Case CGL_DeliEDFKaitoFlg  '�����[���E�d���񓚕����t���O�̏ꍇ
            strSQL_WHERE = strSQL_WHERE & " AND ( CDELIVERY_KAITO_FLG = '1'"
            strSQL_WHERE = strSQL_WHERE & "  OR   CEDF_KAITO_FLG = '1' )"
' 2018/05/07 �� M.Kawamura K545 CS�v���Z�X���P
    End Select
    '�{�@ATT�敪
    Select Case intHonkiAttKbn
        Case CGL_HonkiAttAll    '�{�@ATT�敪�̏��������Ȃ�
            '�����Ȃ�
        Case CGL_Att            'ATT�̏ꍇ
            strSQL_WHERE = strSQL_WHERE & " AND HONKIATTKBN = '1'"
        Case CGL_Honki          '�{�@�̏ꍇ
            strSQL_WHERE = strSQL_WHERE & " AND HONKIATTKBN = '2'"
        Case CGL_Sonota         '���̑��̏ꍇ
            strSQL_WHERE = strSQL_WHERE & " AND HONKIATTKBN IS NULL"
    End Select
    '�c�ƕ\���敪
    Select Case intEigyoDispKbn
        Case CGL_EigyoAll '�����A�C�O�̏��������Ȃ�
            '�����Ȃ�
        Case CGL_Kokunai  '�����̏ꍇ
            strSQL_WHERE = strSQL_WHERE & " AND NGROUP_ID <> '" & strKaigaiEigyoId & "'"
        Case CGL_Kaigai   '�C�O�̏ꍇ
            strSQL_WHERE = strSQL_WHERE & " AND NGROUP_ID <> '" & strKokunaiEigyoId & "'"
    End Select
    
    'SQL����WHERE��ǉ�
    strSQL = strSQL & strSQL_WHERE
    
    '�޲ž�Ă̐���
    Set oraDynaCount = gOraDataBase.CreateDynaset(strSQL, ORADYN_READONLY)
    
     '�ް���������
    If oraDynaCount.EOF = False Then
        '1���ȏ�̏ꍇ
        If 1 <= GF_VarToNum(oraDynaCount![Cnt]) Then
            '�����Z�b�g
            intCnt = GF_VarToNum(oraDynaCount![Cnt])
        '�Y���f�[�^�Ȃ��̏ꍇ
        Else
            '�G���[�ɂ��Ȃ�
            GF_CreateGroupList = True
            Exit Function
        End If
    End If
    
    Set oraDynaCount = Nothing
    
    'SQL��
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "  NGROUP_ID,"
    strSQL = strSQL & "  CGROUP,"
    strSQL = strSQL & "  CMODEL_KAITO"
    strSQL = strSQL & " FROM TCS_GROUP"
    strSQL = strSQL & strSQL_WHERE        'WHERE��ǉ�
    strSQL = strSQL & " ORDER BY NDISPNO"
    
    '�޲ž�Ă̐���
    Set oraDynaData = gOraDataBase.CreateDynaset(strSQL, ORADYN_READONLY)

    '�z��̗v�f����`
    ReDim strGroupIdArray(intCnt - 1)
    
    '���ڐݒ�
    Do
        cntControl.AddItem GF_VarToStr(oraDynaData![CGROUP])
        'ID�p�z����e�敪
        Select Case intIdArrayKbn
            Case CGL_Id              '�O���[�vID�݂̂̏ꍇ
                strGroupIdArray(intIndex) = GF_VarToStr(oraDynaData![NGROUP_ID])
            Case CGL_IdAndModelKaito '�O���[�vID�Ƌ@��}�X�^�񓚕����敪)�̏ꍇ
                '�J���}�Ȃ��Ŋi�[
                strGroupIdArray(intIndex) = GF_VarToStr(oraDynaData![NGROUP_ID]) & "," & GF_VarToStr(oraDynaData![CMODEL_KAITO])
        End Select
        
        intIndex = intIndex + 1
        oraDynaData.MoveNext
    Loop Until (oraDynaData.EOF = True)
    
    Set oraDynaData = Nothing
    
    '�Ō�ɂ�ListIndex������
    cntControl.ListIndex = -1
    
    GF_CreateGroupList = True
    
    Exit Function
    
ErrHandler:
    '�װ�����
    Call GS_ErrorHandler("GF_CreateGroupList", strSQL)
    
End Function

' 2016/12/15 �� M.Tanaka K545 CS�v���Z�X���P  �ǉ�

' 2016/12/15 �� M.Tanaka K545 CS�v���Z�X���P  �ǉ�
Private Function LF_GetIniTable(ByVal strKeyCd As String, ByVal intNumber As Integer) As String
'------------------------------------------------------------------------------
' @(f)
'
' �@�\���@�@:�@�ݒ�l�̎擾
' �@�\�@�@�@:�@INI�t�@�C���}�X�^����ݒ�l���擾����
' �����@�@�@:�@strKeyCd As String       �L�[�R�[�h
' �@�@�@�@�@�@ intNumber As Integer     ����
' �߂�l�@�@:�@�擾�����l
'
' �@�\�����@:
'------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim oraDyna     As OraDynaset
    Dim sSQL        As String

    LF_GetIniTable = ""

    'SQL������
    sSQL = ""
    sSQL = sSQL & "SELECT VCSET_CD FROM M68_INI_TABLE"
    sSQL = sSQL & " WHERE VCKEY_CD = '" & strKeyCd & "'"
    sSQL = sSQL & "   AND NNUMBER = " & intNumber

    '�޲ž�Đ���
    Set oraDyna = gOraDataBase.CreateDynaset(sSQL, ORADYN_READONLY Or ORADYN_NOCACHE)
    
    If oraDyna.EOF = False Then
        LF_GetIniTable = GF_VarToStr(oraDyna![VCSET_CD])
    End If
    
    Set oraDyna = Nothing

    Exit Function

ErrHandler:
    Call GS_ErrorHandler("LF_GetIniTable", sSQL)
End Function
' 2016/12/15 �� M.Tanaka K545 CS�v���Z�X���P  �ǉ�

