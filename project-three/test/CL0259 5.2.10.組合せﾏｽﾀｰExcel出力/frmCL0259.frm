VERSION 5.00
Begin VB.Form frmCL0259 
   BorderStyle     =   1  '???(????)
   Caption         =   "?g?????}?X?^?[Excel?o??"
   ClientHeight    =   4665
   ClientLeft      =   2775
   ClientTop       =   2910
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "?l?r ?S?V?b?N"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCL0259.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6767.535
   ScaleMode       =   0  '???
   ScaleWidth      =   7410
   Begin VB.ComboBox cmbSetCs 
      Height          =   300
      ItemData        =   "frmCL0259.frx":0442
      Left            =   1380
      List            =   "frmCL0259.frx":0444
      Style           =   2  '???????? ??
      TabIndex        =   7
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox txtKeyOpt2 
      Height          =   270
      IMEMode         =   3  '????
      Left            =   1380
      MaxLength       =   8
      TabIndex        =   6
      Text            =   "12345678"
      Top             =   3240
      Width           =   840
   End
   Begin VB.TextBox txtKeyOpt1 
      Height          =   270
      IMEMode         =   3  '????
      Left            =   1380
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "12345678"
      Top             =   2880
      Width           =   840
   End
   Begin VB.TextBox txtAtt 
      Height          =   270
      IMEMode         =   3  '????
      Left            =   1380
      MaxLength       =   11
      TabIndex        =   4
      Text            =   "12345678901"
      Top             =   2520
      Width           =   1110
   End
   Begin VB.TextBox txtModel 
      Height          =   270
      IMEMode         =   3  '????
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "12345678901234567890"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox cmbExportCs 
      Height          =   300
      Left            =   1380
      Style           =   2  '???????? ??
      TabIndex        =   0
      Top             =   1080
      Width           =   1270
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "?????"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "?o??"
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox cmbModelTypeCs 
      Height          =   300
      Left            =   1380
      Style           =   2  '???????? ??
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
   Begin VB.Frame framDATE 
      Caption         =   "?K?p??"
      Height          =   1515
      Left            =   3780
      TabIndex        =   14
      Top             =   1080
      Width           =   3255
      Begin VB.TextBox txtDate 
         Height          =   270
         IMEMode         =   3  '????
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1080
         Width           =   1035
      End
      Begin VB.OptionButton optApply 
         Caption         =   "??V???"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optApply 
         Caption         =   "?S??"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   1395
      End
      Begin VB.OptionButton optApply 
         Caption         =   "???t?w??"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   1395
      End
   End
   Begin VB.TextBox txtModelTypeCd 
      Height          =   270
      IMEMode         =   3  '????
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "1234"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblNowDate 
      BackStyle       =   0  '????
      Caption         =   "2000/11/28"
      BeginProperty Font 
         Name            =   "?l?r ?S?V?b?N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   28
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblDispID 
      BackStyle       =   0  '????
      Caption         =   "???ID"
      BeginProperty Font 
         Name            =   "?l?r ?S?V?b?N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblOpt2 
      Caption         =   "KEY-OPT2"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   3274
      Width           =   1005
   End
   Begin VB.Label lblCon3 
      Caption         =   "(?O????v)"
      Height          =   255
      Left            =   2280
      TabIndex        =   25
      Top             =   3289
      Width           =   1215
   End
   Begin VB.Label lblPack 
      Caption         =   "?g??????"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   3649
      Width           =   1005
   End
   Begin VB.Label lblOpt1 
      Caption         =   "KEY-OPT1"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   2914
      Width           =   1005
   End
   Begin VB.Label lblCon2 
      Caption         =   "(?O????v)"
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   2929
      Width           =   1215
   End
   Begin VB.Label lblCon1 
      Caption         =   "(?O????v)"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   2554
      Width           =   1215
   End
   Begin VB.Label lblAtt 
      Caption         =   "ATT?R?[?h"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   2554
      Width           =   1005
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '????????
      AutoSize        =   -1  'True
      Caption         =   "?g?????}?X?^?[Excel?o??"
      BeginProperty Font 
         Name            =   "?l?r ?S?V?b?N"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1365
      TabIndex        =   19
      Top             =   480
      Width           =   4725
   End
   Begin VB.Label lblExport_CS 
      Caption         =   "?s?A??"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   1114
      Width           =   1005
   End
   Begin VB.Label lblModel_Type_Cd 
      Caption         =   "????"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   1474
      Width           =   1005
   End
   Begin VB.Label lblModel 
      Caption         =   "?@??"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2194
      Width           =   1005
   End
   Begin VB.Label lblModel_Type_Cs 
      Caption         =   "???R?[?h"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   1834
      Width           =   1005
   End
End
Attribute VB_Name = "frmCL0259"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' @(s)
'   ?v???W?F?N?g??  : TLF??????
'   ???W???[????    : frmCL0259
'   ?t?@?C????      : frmCL0259.frm
'   Version         : 1.0.0.1
'   ?@?\????       ?F  ?g???????Excel?o??
'   ????         ?F J.Hamaji
'   ????         ?F 2004/10/20
'   ?C??????       ?F 2004/10/28 THS T.Y (?????????Excel?o?????)
'   ?@?@?@?@       ?F 2005/07/07 THS J.Yamaoka (?????????t??????Ver???X)
'   ?@?@?@?@       ?F 2006/04/13 THS Sugawara Ver1.0.0.1  6????????f?[?^?????V?[?g??o????????s?????C??
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ?????
'------------------------------------------------------------------------------
Option Explicit
'------------------------------------------------------------------------------
' ???W???[???????
'------------------------------------------------------------------------------
Private mstrUserID              As String       '?[?????????i?[(??PG?p)
Private mbolLoadFlag            As Boolean      '????????????
Private mlngKensu               As Long         'Excel?o?????

'------------------------------------------------------------------------------
' ???W???[??????
'------------------------------------------------------------------------------

Public Property Get LoadFlag() As Boolean
'------------------------------------------------------------------------------
' ?????[?h?t???O
'------------------------------------------------------------------------------
    LoadFlag = mbolLoadFlag
End Property

Private Sub cmdClose_Click()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:????I??????
' ?@?\?@?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------
'?f?[?^?x?[?X???f
Call GS_DBClose
End

End Sub
Private Sub cmdOutput_Click()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?o??{?^??????
' ?@?\?@?@?@:?o??{?^??????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------
On Error GoTo ErrHandler

    Dim blnRet As Boolean

    '??d????h?~(????????????v????????????)
    If (Screen.MousePointer = vbHourglass) Then
        '???????
    Else
        '?????????????v????
        Screen.MousePointer = vbHourglass
        Me.Enabled = False
        
        'Excel?o????????J?n????
         blnRet = LF_Load_Process
            
        '?}?E?X?|?C???^???
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        
        'Excel?o?????????????
        If blnRet = True Then
            '?s?A????t?H?[?J?X?Z?b?g????
            cmbExportCs.SetFocus
        End If
 
    End If
    
    Exit Sub
    
ErrHandler:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Call GS_ErrorHandler("cmdOutput_Click")
    
End Sub

Private Sub Form_Load()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?t?H?[?????[?h
' ?@?\?@?@?@:
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------
    Dim bolRet  As Boolean
    Dim strDate     As String
    
    '???????????(?????l???)
    mbolLoadFlag = False
    
    '???[?U?[ID?èÔ
    mstrUserID = gstrUserID
    
    '???????\??
    If (GF_GetSYSDATE(strDate, 1) = True) Then
        lblDispID.Caption = App.EXEName
        lblNowDate.Caption = Format$(CDate(strDate), "YYYY/MM/DD")
    End If

    '???\????u???
    Call GS_CenteringForm(Me, 0)
    
    '???????èÔ
    Call GF_FormInit(Me)
        
    '?????u?~?v???????
    Call GS_DelControlBox(Me)
       
    '?R???g???[?????????
    Call LS_InitControl
            
    '????????????
    mbolLoadFlag = True
    
End Sub

Private Sub LS_InitControl()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?R???g???[????????
' ?@?\?@?@?@:?e?R???g???[????????????s??
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------
    
    txtModelTypeCd.Text = ""                ''???R?[?h
    txtModel.Text = ""                      ''?@??
    txtAtt.Text = ""                        ''ATT?R?[?h
    txtKeyOpt1.Text = ""                    ''KEY-OPT1
    txtKeyOpt2.Text = ""                    ''KEY-OPT2
    txtDate.Text = ""                       ''?K?p??
    optApply(0).Value = True

    '???t?w????
    Call LS_ApplyCtrl
    
    '?R???{?{?b?N?X???
    ''?s?A???R???{??
    If GF_Com_CtlAdditem2(cmbExportCs, "TGLOBAL_EXPORT2", , 1, 2, 100) = False Then
        Call GF_MsgBoxDB(Me.Caption, "WTK009", "O", "C")
        Exit Sub
    End If
    ''?????R???{??
    If GF_Com_CtlAdditem2(cmbModelTypeCs, "GLOBAL_cboSyasyuKubun", , 1, 2, 100) = False Then
        Call GF_MsgBoxDB(Me.Caption, "WTK009", "O", "C")
        Exit Sub
    End If
    ''?g???????R???{??
    If GF_Com_CtlAdditem2(cmbSetCs, strKumiKbn, , 1, 1) = False Then
        Call GF_MsgBoxDB(Me.Caption, "WTK009", "O", "C")
        Exit Sub
    End If
    
End Sub

Private Sub LS_ApplyCtrl()
''--------------------------------------------------------------------------------
'' @(f)
'' ?@?\??   : ???t?w????
'' ?@?\     :
'' ????     :
'' ???l   :
'' ?@?\???? :
''--------------------------------------------------------------------------------

    Call LS_EnabledEx(txtDate, optApply(2).Value)

End Sub

Private Sub LS_EnabledEx(ByVal ctrObject As Control, ByVal blnFlag As Boolean)
'--------------------------------------------------------------------------------
' @(f)
' ?@?\??   : ?R???g???[????Enable????
' ?@?\     :
' ????     : ByVal ctrObject As Control    ?R???g???[???I?u?W?F?N?g
'          : ByVal blnFlag   As Boolean    ????t???O
' ???l   :
' ?@?\???? :
'--------------------------------------------------------------------------------

    '' ?R???g???[???????
    ctrObject.Enabled = blnFlag
        
    '' ?w?i?F???X
    If blnFlag Then

        ctrObject.BackColor = cNoProtectColor
    Else

        ctrObject.BackColor = cProtectColor
    End If
        
End Sub

Private Sub optApply_Click(Index As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?K?p???I?v?V?????{?^??????
' ?@?\?@?@?@:?K?p???I?v?V?????{?^??????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------
    
    '???t?w????
    Call LS_ApplyCtrl
    
    '' ???t?w???O?????t?N???A
    If Index <> 2 Then
        txtDate.Text = ""
    End If
    
End Sub

Private Sub cmbExportCs_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?s?A???@?L?[?????
' ?@?\?@?@?@:?s?A???@?L?[?????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler

    Select Case KeyAscii
        Case vbKeyReturn
            GS_Com_NextCntl cmbExportCs
    End Select
    
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("cmbExportCs_KeyPress")

End Sub
Private Sub cmbModelTypeCs_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?????@?L?[?????
' ?@?\?@?@?@:?????@?L?[?????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler

    Select Case KeyAscii
        Case vbKeyReturn
            GS_Com_NextCntl cmbModelTypeCs
    End Select
    
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("cmbModelTypeCs_KeyPress")

End Sub

Private Sub txtModelTypeCd_GotFocus()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:???R?[?h?@?t?H?[?J?X?èÔ
' ?@?\?@?@?@:???R?[?h?@?t?H?[?J?X?èÔ
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler
    
    Call GS_TextSelect(txtModelTypeCd)
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtModelTypeCd_GotFocus")

End Sub

Private Sub txtModelTypeCd_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:???R?[?h?@?L?[?????
' ?@?\?@?@?@:???R?[?h?@?L?[?????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler

    Select Case KeyAscii
        Case vbKeyReturn
            GS_Com_NextCntl txtModelTypeCd
    End Select
    If GF_Com_KeyPress(13, KeyAscii) = 0 Then Exit Sub
    
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtModelTypeCd_KeyPress")

End Sub
Private Sub txtModel_GotFocus()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?@??@?t?H?[?J?X?èÔ
' ?@?\?@?@?@:?@??@?t?H?[?J?X?èÔ
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler
    
    Call GS_TextSelect(txtModel)
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtModel_GotFocus")

End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?@??@?L?[?????
' ?@?\?@?@?@:?@??@?L?[?????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler

    Select Case KeyAscii
        Case vbKeyReturn
            GS_Com_NextCntl txtModel
    End Select
    If GF_Com_KeyPress(13, KeyAscii) = 0 Then Exit Sub
    
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtModel_KeyPress")

End Sub

Private Sub txtAtt_GotFocus()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:ATT?R?[?h?@?t?H?[?J?X?èÔ
' ?@?\?@?@?@:ATT?R?[?h?@?t?H?[?J?X?èÔ
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler
    
    Call GS_TextSelect(txtAtt)
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtAtt_GotFocus")

End Sub

Private Sub txtAtt_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:ATT?R?[?h?@?L?[?????
' ?@?\?@?@?@:ATT?R?[?h?@?L?[?????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler

    Select Case KeyAscii
        Case vbKeyReturn
            GS_Com_NextCntl txtAtt
    End Select
    If GF_Com_KeyPress(13, KeyAscii) = 0 Then Exit Sub

    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtAtt_KeyPress")

End Sub
Private Sub txtKeyOpt1_GotFocus()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:KEY-OPT1?@?t?H?[?J?X?èÔ
' ?@?\?@?@?@:KEY-OPT1?@?t?H?[?J?X?èÔ
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler
    
    Call GS_TextSelect(txtKeyOpt1)
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtKeyOpt1_GotFocus")

End Sub

Private Sub txtKeyOpt1_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:KEY-OPT1  ?L?[?????
' ?@?\?@?@?@:KEY-OPT1?@?L?[?????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler

    Select Case KeyAscii
        Case vbKeyReturn
            GS_Com_NextCntl txtKeyOpt1
    End Select
    If GF_Com_KeyPress(13, KeyAscii) = 0 Then Exit Sub

    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtKeyOpt1_KeyPress")

End Sub
Private Sub txtKeyOpt2_GotFocus()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:KEY-OPT2?@?t?H?[?J?X?èÔ
' ?@?\?@?@?@:KEY-OPT2?@?t?H?[?J?X?èÔ
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler
    
    Call GS_TextSelect(txtKeyOpt2)
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtKeyOpt2_GotFocus")

End Sub

Private Sub txtKeyOpt2_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:KEY-OPT2  ?L?[?????
' ?@?\?@?@?@:KEY-OPT2?@?L?[?????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler

    Select Case KeyAscii
        Case vbKeyReturn
            GS_Com_NextCntl txtKeyOpt2
    End Select
    If GF_Com_KeyPress(13, KeyAscii) = 0 Then Exit Sub

    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("txtKeyOpt2_KeyPress")

End Sub
Private Sub cmbSetCs_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?????@?L?[?????
' ?@?\?@?@?@:?????@?L?[?????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

On Error GoTo ErrHandler

    Select Case KeyAscii
        Case vbKeyReturn
            GS_Com_NextCntl cmbSetCs
    End Select
    
    Exit Sub
ErrHandler:
    Call GS_ErrorHandler("cmbSetCs_KeyPress")

End Sub
Private Sub txtDate_GotFocus()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?K?p???w??@?t?H?[?J?X?èÔ
' ?@?\?@?@?@:?K?p???w??@?t?H?[?J?X?èÔ
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

    txtDate.Text = Replace(txtDate.Text, "/", "")
    txtDate.MaxLength = 8

End Sub

Private Sub txtDate_LostFocus()
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?K?p???w??@?t?H?[?J?X????
' ?@?\?@?@?@:?K?p???w??@?t?H?[?J?X????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------

    txtDate.MaxLength = 10
    If Len(txtDate.Text) = 8 Then
        txtDate.Text = Format(txtDate.Text, "@@@@/@@/@@")
    End If
    
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------
' @(f)
' ?@?\???@?@:?K?p???w??@?L?[????
' ?@?\?@?@?@:?K?p???w??@?L?[????
' ?????@?@?@:
' ???l?@?@:
' ?@?\?????@:
'------------------------------------------------------------------------------
    'Enter?L?[??t?H?[?J?X???
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call GS_Com_NextCntl(txtDate)
        Exit Sub
    End If

    '??????O????s???
    If GF_Com_KeyPress(1, KeyAscii) = 0 Then Exit Sub

End Sub


Private Function LF_DirectoryCheck() As Boolean
'--------------------------------------------------------------------------------
' @(f)
' ?@?\??    : Excel?t?@?C???o????f?B???N?g???`?F?b?N(?T?[?o/?N???C?A???g)
' ?@?\      :
' ????      :
' ???l    : TRUE?F???? FALSE:?G???[ Boolean
' ?@?\?????@:
'--------------------------------------------------------------------------------
On Error GoTo ErrHandler
    Dim blnPresenceFlg      As Boolean          '?t?@?C???L???t???O
    Dim strPath(0)          As String            '0?F?p?X??
    Dim intMsg              As String
    
    Dim strMsg(0)           As String           ''ADD 2005/07/07
    Dim strMsgbox           As String           ''ADD 2005/07/07
    
    LF_DirectoryCheck = False
    
    '?N???C?A???g????p?X???i?[
    strPath(0) = gstrClientPath
    
    ''DEL 2004/10/28 THS T.Y (?????????Excel?o?????) START>>>>>
'''''    '?T?[?o???o????f?B???N?g???`?F?b?N
'''''    If GF_FSO_GetFileInfo(FSO_FolderExists, blnPresenceFlg, gstrServerPath) = False Then Exit Function
'''''
'''''    If blnPresenceFlg = False Then
'''''        '?T?[?o???o???f?B???N?g???`?F?b?N???s
'''''        '???O?o??
'''''        Call GF_GetMsg_Addition("WTK401", , False, True)
'''''        'MSG?\??
'''''        Call GF_MsgBoxDB(Me.Caption, "WTK401", "OK", "C")
'''''        Exit Function
'''''    End If
    ''<<<<<END
    
    '?N???C?A???g???o????f?B???N?g???`?F?b?N
    If GF_FSO_GetFileInfo(FSO_FolderExists, blnPresenceFlg, gstrClientPath) = False Then Exit Function
    
    If blnPresenceFlg = False Then
        '?N???C?A???g???o???f?B???N?g???`?F?b?N???s
        '???O?o???MSG?\??
        Call GF_GetMsg_Addition("WTK402", strPath, True, True, , False, "C")
        
        Exit Function
    End If
    
    '?N???C?A???g??f?B???N?g????????t?@?C???????`?F?b?N
    If GF_FSO_GetFileInfo(FSO_FileExists, blnPresenceFlg, gstrClientPath & gstrFileName) = False Then Exit Function
    
    If blnPresenceFlg = True Then
        '?N???C?A???g???o???f?B???N?g????????t?@?C???L??
        '???O?o???MSG?\??
'' 2005/07/07 ADD J.Y >>>>> ???????m?FMSG?p?X?????
        strMsg(0) = gstrClientPath & gstrFileName
        strMsgbox = GF_GetMsg_Addition("QCL074", strMsg(0), False, False)
        intMsg = GF_MsgBox(Me.Caption, strMsgbox, "OC", "Q")
''        intMsg = GF_MsgBoxDB(Me.Caption, "QTK008", "OC", "Q")
        'MSGOX??L?????Z???{?^???????????‡€???I??
        If intMsg = 2 Then Exit Function
        
    End If
    
    LF_DirectoryCheck = True
    
    Exit Function

ErrHandler:

    Call GS_ErrorHandler("LF_DirectoryCheck")

End Function

Private Function LF_Load_Process() As Boolean
'--------------------------------------------------------------------------------
' @(f)
' ?@?\??    : Excel?o??????J?n????
' ?@?\      :
' ????      :
' ???l    : TRUE?F???? FALSE:?G???[ Boolean
' ?@?\?????@:
'--------------------------------------------------------------------------------
On Error GoTo ErrHandler
    
    LF_Load_Process = False
    
    '???N??????x???????o???????
        ''UPD 2004/10/28 THS T.Y (?????????Excel?o?????) START>>>>>
        If gstrClientPath = "" Then
'''''        If gstrServerPath = "" Then
'''''            'MSG?\??(EXCEL?o???t?H???_?èÔ?????(?T?[?o))
'''''            Call GF_MsgBoxDB(frmCL0259.Caption, "WTK398", "OK", "C")
'''''            Exit Function
'''''        ElseIf gstrClientPath = "" Then
        ''<<<<<END
            'MSG?\??(EXCEL?o???t?H???_?èÔ?????(?N???C?A???g))
            Call GF_MsgBoxDB(frmCL0259.Caption, "WTK399", "OK", "C")
            Exit Function
        ElseIf gstrFileName = "" Then
            'MSG?\??(EXCEL?t?@?C?????èÔ?????)
            Call GF_MsgBoxDB(frmCL0259.Caption, "WTK400", "OK", "C")
            Exit Function
        End If
        
        '???????`?F?b?N
        If LF_InputCheck() = False Then Exit Function
        
        'Excel?t?@?C???o????f?B???N?g???`?F?b?N
        If LF_DirectoryCheck() = False Then Exit Function
        
        'Excel?t?@?C???o??
        If LF_OutPutStart = False Then Exit Function
        
    LF_Load_Process = True
    
    Exit Function

ErrHandler:
Call GS_ErrorHandler("LS_Load_Process")

End Function

Private Function LF_InputCheck() As Boolean
'--------------------------------------------------------------------------------
' @(f)
' ?@?\??    : ???????`?F?b?N
' ?@?\      :
' ????      :
' ???l    : TRUE?F???? FALSE:???????G???[   Boolean
' ?@?\?????@:
'--------------------------------------------------------------------------------
On Error GoTo ErrHandler
    
    Dim strSQL                  As String
    Dim clsOracle               As OraDynaset
    Dim strExportCs             As String
    Dim strModelTypeCs          As String
    
    LF_InputCheck = False
    
    ''????????I??????????G???[
    If cmbExportCs.Text = "" And cmbModelTypeCs.Text = "" And txtModelTypeCd.Text = "" _
        And txtModel.Text = "" And txtAtt.Text = "" And txtKeyOpt1.Text = "" And txtKeyOpt2.Text = "" _
        And cmbSetCs.Text = "" Then
            'MSG?\???i???t??????j
            Call GF_MsgBoxDB(Me.Caption, "WCL001", "OK", "E")
            '?t?H?[?J?X???
            Call GS_Com_TxtGotFocus(cmbExportCs)
            Exit Function
    End If
    
    '???t?w??`?F?b?N
    If optApply(2).Value = True Then
    
        If Trim(txtDate.Text) = "" Then
            'MSG?\???i???t??????j
            Call GF_MsgBoxDB(Me.Caption, "WTK471", "OK", "E")
            '?t?H?[?J?X???
            Call GS_Com_TxtGotFocus(txtDate)
            Exit Function
            
        ElseIf Len(Replace(txtDate.Text, "/", "")) <> 8 Then
            'MSG?\???i???t????8????O?j
            Call GF_MsgBoxDB(Me.Caption, "WTK472", "OK", "E")
            '?t?H?[?J?X???
            Call GS_Com_TxtGotFocus(txtDate)
            Exit Function
            
        ElseIf IsDate(txtDate.Text) = False Then
            'MSG?\???i???t?s???G???[?j
            Call GF_MsgBoxDB(Me.Caption, "WTK198", "OK", "E")
            '?t?H?[?J?X???
            Call GS_Com_TxtGotFocus(txtDate)
            Exit Function
            
        End If
    
    End If
    
    ''?@????`?F?b?N
    ''?s?A???I??
    Select Case cmbExportCs.ListIndex
        Case 0
            strExportCs = "B"
        Case 1
            strExportCs = " "
        Case 2
            strExportCs = "A"
        Case Else
            strExportCs = "B"
    End Select

''    strExportCs = GF_Com_CboGetText(cmbExportCs)
    If cmbModelTypeCs.Text <> "" Then
        strModelTypeCs = GF_Com_CboGetText(cmbModelTypeCs)
    Else
        strModelTypeCs = ""
    End If
        
    strSQL = LF_ChkModelSql(strExportCs, strModelTypeCs, txtModelTypeCd.Text, txtModel.Text)
    
    Set clsOracle = gOraDataBase.CreateDynaset(strSQL, ORADYN_READONLY)
    
    '?f?[?^?L??????
    If (clsOracle.EOF = True) Then
        '?f?[?^???????
        'MSG?\??
        Call GF_MsgBoxDB(Me.Caption, "ICL004", "OK", "I")
        Exit Function
    End If
    
    LF_InputCheck = True
    
    Exit Function
    
ErrHandler:
    Call GS_ErrorHandler("LF_InputCheck")
End Function

Private Function LF_OutPutStart() As Boolean
'--------------------------------------------------------------------------------
' @(f)
' ?@?\??    : Excel?o?????
' ?@?\      :
' ????      :
' ???l    : TRUE?F???? FALSE:?G???[ Boolean
' ?@?\?????@:
'--------------------------------------------------------------------------------
On Error GoTo ErrHandler
    Dim blnPresenceFlg          As Boolean
    Dim strSQL                  As String
    Dim clsOracle               As OraDynaset
    Dim strlogHuka(0 To 2)      As String       '0?F(????)?g?????}?X?^?[?@1?F?????@2?F?t?@?C????
    Dim strFuyoInf              As String       '?t?^???
    Dim strHuka(0 To 1)         As String       '0?F?????@1?F?t?@?C????
    Dim strfuyo                 As String       '?s?A??
    
    LF_OutPutStart = False
    
    '????
    strSQL = LF_GetSelectSql()
    Set clsOracle = gOraDataBase.CreateDynaset(strSQL, ORADYN_READONLY)
    
    '?f?[?^?L??????
    If (clsOracle.EOF = True) Then
        '?f?[?^???????
        'MSG?\??
        Call GF_MsgBoxDB(Me.Caption, "ICL004", "OK", "I")
        Exit Function
    End If
    
    ''DEL 2004/10/28 THS T.Y (?????????Excel?o?????) START>>>>>
'''''    '?T?[?o??Excel?t?@?C???????`?F?b?N
'''''    If GF_FSO_GetFileInfo(FSO_FileExists, blnPresenceFlg, gstrServerPath & gstrFileName) = False Then Exit Function
'''''    'Excel?t?@?C?L??????
'''''    If blnPresenceFlg = True Then
'''''
'''''        '?T?[?o??Excel?t?@?C??????
'''''        '?t?@?C????????
'''''        Kill gstrServerPath & gstrFileName
'''''        DoEvents
'''''
'''''    End If
    ''<<<<<END
    
    'Excel???????????
    If LF_CreateCombExcel(clsOracle) = False Then
        '??????????s
        '???O?o??
        Call GF_GetMsg_Addition("WCL006", , False, True)
        'MSG?\??
        Call GF_MsgBoxDB(frmCL0259.Caption, "WCL006", "OK", "C")
        Exit Function
    End If
    
    ''DEL 2004/10/28 THS T.Y (?????????Excel?o?????) START>>>>>
'''''    'Excel?t?@?C???R?s?[????
'''''    If GF_FSO_CopyFile(gstrServerPath & gstrFileName, gstrClientPath & gstrFileName) = False Then
'''''        'Excel?t?@?C???R?s?[???????s
'''''        strFuyoInf = "?R?s?[??:[" & gstrServerPath & gstrFileName & "],?R?s?[??:[" & gstrClientPath & gstrFileName & "]"
'''''        '???O?o??
'''''        Call GF_GetMsg_Addition("WTK403", , False, True, strFuyoInf)
'''''        'MSG?\??
'''''        Call GF_MsgBoxDB(frmCL0259.Caption, "WTK403", "OK", "C")
'''''        '?t?@?C????????
'''''        Kill gstrServerPath & gstrFileName
'''''        DoEvents
'''''        Exit Function
'''''    End If
'''''
'''''    '?t?@?C????????
'''''    Kill gstrServerPath & gstrFileName
'''''    DoEvents
    ''<<<<<END
        
    '???O?p?t???????z??i?[
    strlogHuka(0) = "?g?????}?X?^"
    strlogHuka(1) = mlngKensu
    strlogHuka(2) = gstrFileName
    '???b?Z?[?W?p?t???????z??i?[
    strHuka(0) = mlngKensu
    strHuka(1) = gstrFileName
    
    '?t?^???i?[
    '?s?A?R???{????
    Select Case cmbExportCs.ListIndex
        Case 0
            strfuyo = ""
        Case 1
            strfuyo = "????"
        Case 2
            strfuyo = "?C?O"
        Case 3
            strfuyo = "?????E?C?O"
    End Select
    
    strFuyoInf = "?s?A??[" & strfuyo & "],"
    strFuyoInf = strFuyoInf & "????[" & GF_Com_CboGetText(cmbModelTypeCs) & "],"
    strFuyoInf = strFuyoInf & "???R?[?h[" & txtModelTypeCd.Text & "],"
    strFuyoInf = strFuyoInf & "?@??[" & txtModel.Text & "],"
    strFuyoInf = strFuyoInf & "ATT?R?[?h[" & txtAtt.Text & "],"
    strFuyoInf = strFuyoInf & "KEY-OPT1[" & txtKeyOpt1.Text & "],"
    strFuyoInf = strFuyoInf & "KEY-OPT2[" & txtKeyOpt2.Text & "],"
    strFuyoInf = strFuyoInf & "?g??????[" & cmbSetCs.Text & "],"
    
''    strFuyoInf = strFuyoInf & IIf(Trim(strFuyoInf) = "", "", ",")
    strFuyoInf = strFuyoInf & "?K?p???F"
    If optApply(0).Value = True Then
        '??V???
        strFuyoInf = strFuyoInf & "??V???"
    ElseIf optApply(1).Value = True Then
        '?S??
        strFuyoInf = strFuyoInf & "?S??"
    ElseIf optApply(2).Value = True Then
        '???t?w??
        strFuyoInf = strFuyoInf & "???t?w??[" & txtDate.Text & "]"
    End If
    
    'Excel?t?@?C???o?????????I??
    '???O?o??
    Call GF_GetMsg_Addition("ITK256", strlogHuka, False, True, strFuyoInf)
    'MSG?\??
    Call GF_GetMsg_Addition("ICL075", strHuka, True, False, , False, "I")
    
    LF_OutPutStart = True
    
    Exit Function
    
ErrHandler:
Call GS_ErrorHandler("LF_OutPutStart")
End Function

Private Function LF_ChkModelSql(ByVal strShiyuKbn As String, ByVal strModelTypeCs As String, _
                                ByVal strModelTypeCd As String, ByVal strModel As String) As String
'--------------------------------------------------------------------------------
' @(f)
' ?@?\??    : ?@??????A?@????????????????f?[?^?èÔ?pSQL
' ?@?\      :
' ????      :   [in]   strShiyuKbn       String  ???????s?A??(????,?C?O,????C?O)
'               [in]   strModelTypeCs    String  ??????????
'               [in]   strModelTypeCd    String  ?????????R?[?h
'               [in]   strModel     ?@?@ String  ???????@??
' ???l    : ????SQL
' ?@?\?????@: ?èÔ?pSQL?????
'--------------------------------------------------------------------------------
    Dim strSQL          As String
    Dim strCondition    As String   ''???o????
    Dim strSubCondition As String   ''???
    
    strCondition = ""
   
    strSQL = ""
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "  HONKITYPE"             '?@??
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  THJHNKTYPEMR"          '?@??}?X?^?[
        
    
    '???R?[?h????
    If Trim(strModelTypeCd) <> "" Then
        strCondition = " WHERE "
        strCondition = strCondition & " SERIESCD = '" & GF_ChangeQuateSing(Trim(strModelTypeCd)) & "'"
        strSubCondition = " AND "
    Else
        ''????????:'WHERE',????:'AND'
        If Trim(strCondition) <> "" Then
            strSubCondition = " AND "
        Else
            strSubCondition = " WHERE "
        End If
    End If
    
    '????????
    If Trim(strModelTypeCs) <> "" Then
        strCondition = strCondition & strSubCondition
        strCondition = strCondition & " SYASYUKBN = '" & GF_ChangeQuateSing(strModelTypeCs) & "'"
        strSubCondition = " AND "
    Else
        ''????????:'WHERE',????:'AND'
        If Trim(strCondition) <> "" Then
            strSubCondition = " AND "
        Else
            strSubCondition = " WHERE "
        End If
    End If
    
    '?@?????
    If Trim(strModel) <> "" Then
        strCondition = strCondition & strSubCondition
        strCondition = strCondition & " HONKITYPE = '" & GF_ChangeQuateSing(Trim(strModel)) & "'"
        strSubCondition = " AND "
    Else
        ''????????:'WHERE',????:'AND'
        If Trim(strCondition) <> "" Then
            strSubCondition = " AND "
        Else
            strSubCondition = " WHERE "
        End If
    End If
    
    '?s?A????
    If strShiyuKbn <> "B" Then
        strCondition = strCondition & strSubCondition
        strCondition = strCondition & "  SHIYUKBN = '" & strShiyuKbn & "'"
    Else
        ''???????
    End If
    
    strSQL = strSQL & strCondition
    
    strSQL = strSQL & "  ) UNION ALL "
    strSQL = strSQL & "  ( SELECT "
    strSQL = strSQL & "  HONKITYPE"             '?@??
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  M34_HNK_HISTORY"       '?@??}?X?^?[????????
    strSQL = strSQL & strCondition & ")"
    
    LF_ChkModelSql = strSQL
    
End Function

Private Function LF_GetSelectSql() As String
'--------------------------------------------------------------------------------
' @(f)
' ?@?\??    : ?g?????}?X?^????f?[?^?èÔ?pSQL
' ?@?\      :
' ???l    : ????SQL
' ?@?\?????@: ?èÔ?pSQL?????
'--------------------------------------------------------------------------------
    Dim strSQL          As String
    Dim strApply        As String
    Dim strCondition    As String       '???????i?[??????
    Dim strSubCondition As String       '???????????????
    Dim strExportCs     As String
    Dim strModelTypeCs  As String
    
    strCondition = ""
    strSubCondition = ""
   
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "  K1.SHIYUKBN"              '?s?A??
    strSQL = strSQL & " ,K1.HONKITYPE"             '?@??
    strSQL = strSQL & " ,K1.ATTCD"                 'ATT?R?[?h
    strSQL = strSQL & " ,K1.KEYKBN "               'KEY??
    strSQL = strSQL & " ,K1.KEYOPT1 "              'KEY-OPT1
    strSQL = strSQL & " ,K1.KEYOPT2 "              'KEY-OPT2
    strSQL = strSQL & " ,K1.KUMINO "               '?g??????
    strSQL = strSQL & " ,K1.TKYMDS "               '?K?p?J?n??
    strSQL = strSQL & " ,K1.TKYMDE "               '?K?p?I????
    strSQL = strSQL & " ,K1.KUMIKBN "              '?g??????
    strSQL = strSQL & " ,K1.KUMI01OPT "            '?g????OPT 01
    strSQL = strSQL & " ,K1.KUMI02OPT "            '?g????OPT 02
    strSQL = strSQL & " ,K1.KUMI03OPT "            '?g????OPT 03
    strSQL = strSQL & " ,K1.KUMI04OPT "            '?g????OPT 04
    strSQL = strSQL & " ,K1.KUMI05OPT "            '?g????OPT 05
    strSQL = strSQL & " ,K1.KUMI06OPT "            '?g????OPT 06
    strSQL = strSQL & " ,K1.KUMI07OPT "            '?g????OPT 07
    strSQL = strSQL & " ,K1.KUMI08OPT "            '?g????OPT 08
    strSQL = strSQL & " ,K1.KUMI09OPT "            '?g????OPT 09
    strSQL = strSQL & " ,K1.KUMI10OPT "            '?g????OPT 10
    strSQL = strSQL & " ,K1.CAGENCY_NOCHECK_FLG "  '???X?`?F?b?N?s?v?t???O
    strSQL = strSQL & " ,K1.ATTFLG "               'ATT???????t???O
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  THJKUMIMR K1"             '?g?????}?X?^?[
        
    strSQL = strSQL & " , (SELECT K3.SHIYUKBN,K3.HONKITYPE,K3.ATTCD,K3.KEYKBN,K3.KEYOPT1,"
    strSQL = strSQL & "           K3.KEYOPT2,K3.CPROCTRL_ALARM_FLG,K3.KUMINO,"
    ''?K?p?????????????
    If optApply(0).Value = True Then
        strSQL = strSQL & " MAX(K3.TKYMDS) TKYMDS "
    Else
        strSQL = strSQL & " TKYMDS"
    End If
    strSQL = strSQL & " FROM THJKUMIMR K3 "
    strSQL = strSQL & " ,(SELECT * FROM THJHNKTYPEMR"
    strSQL = strSQL & "    UNION ALL "
    strSQL = strSQL & "   SELECT * FROM M34_HNK_HISTORY ) HNK "
    strSQL = strSQL & " , (SELECT CDVAL FROM THJCMBXMR WHERE "
    
    strSQL = strSQL & " CMBNAME = '" & strKumiKbn & "'"
    
    ''?g??????
    If Trim(cmbSetCs.Text) <> "" Then
        strSQL = strSQL & " AND CDVAL = '" & GF_Com_CboGetCode(cmbSetCs) & "'"
    Else
    End If
    
    strSQL = strSQL & " ) CMB "
    
    ''?s?A???I??
    Select Case cmbExportCs.ListIndex
        Case 0
            strExportCs = "B"
        Case 1
            strExportCs = " "
        Case 2
            strExportCs = "A"
        Case Else
            strExportCs = "B"
    End Select
    
    If strExportCs <> "B" Then
        strCondition = " WHERE "
        strSQL = strSQL & strCondition
        strSQL = strSQL & " HNK.SHIYUKBN = '" & strExportCs & "'"
        strSubCondition = " AND "
    Else
        ''????????:'WHERE',????:'AND'
        If Trim(strCondition) <> "" Then
            strSubCondition = " AND "
        Else
            strSubCondition = " WHERE "
        End If
    End If
    
    ''???R?[?h
    If Trim(txtModelTypeCd.Text) <> "" Then
        strCondition = strSubCondition
        strSQL = strSQL & strSubCondition
        strSQL = strSQL & " HNK.SERIESCD = '" & GF_ChangeQuateSing(Trim(txtModelTypeCd)) & "'"
        strSubCondition = " AND "
    Else
        ''????????:'WHERE',????:'AND'
        If Trim(strCondition) <> "" Then
            strSubCondition = " AND "
        Else
            strSubCondition = " WHERE "
        End If
    End If

    ''????
    strModelTypeCs = GF_Com_CboGetText(cmbModelTypeCs)
    If Trim(strModelTypeCs) <> "" Then
        strCondition = strSubCondition
        strSQL = strSQL & strSubCondition
        strSQL = strSQL & " HNK.SYASYUKBN = '" & GF_ChangeQuateSing(Trim(strModelTypeCs)) & "'"
        strSubCondition = " AND "
    Else
        ''????????:'WHERE',????:'AND'
        If Trim(strCondition) <> "" Then
            strSubCondition = " AND "
        Else
            strSubCondition = " WHERE "
        End If
    End If
    
    '?@??
    If Trim(txtModel.Text) <> "" Then
        strCondition = strSubCondition
        strSQL = strSQL & strSubCondition
        strSQL = strSQL & " HNK.HONKITYPE = '" & GF_ChangeQuateSing(Trim(txtModel.Text)) & "'"
    Else
        ''????????:'WHERE',????:'AND'
        If Trim(strCondition) <> "" Then
            strSubCondition = " AND "
        Else
            strSubCondition = " WHERE "
        End If
    End If
    
    ''????????:'WHERE',????:'AND'
    If Trim(strCondition) <> "" Then
        strSubCondition = " AND "
    Else
        strSubCondition = " WHERE "
    End If
    
    strSQL = strSQL & strSubCondition
    strSQL = strSQL & " K3.SHIYUKBN = HNK.SHIYUKBN "
    strSQL = strSQL & " AND K3.HONKITYPE = HNK.HONKITYPE "
    ''ATT?R?[?h
    If Trim(txtAtt.Text) <> "" Then
        strSQL = strSQL & " AND K3.ATTCD LIKE '" & GF_ChangeQuateSing(Trim(txtAtt.Text)) & "%'"
    End If
    ''KEY-OPT1
    If Trim(txtKeyOpt1.Text) <> "" Then
        strSQL = strSQL & " AND K3.KEYOPT1 LIKE '" & GF_ChangeQuateSing(Trim(txtKeyOpt1.Text)) & "%'"
    End If
    ''KEY-OPT2
    If Trim(txtKeyOpt2.Text) <> "" Then
        strSQL = strSQL & " AND K3.KEYOPT2 LIKE '" & GF_ChangeQuateSing(Trim(txtKeyOpt2.Text)) & "%'"
    End If
    ''????A???[???t???O
    strSQL = strSQL & " AND K3.CPROCTRL_ALARM_FLG ='" & strProAlarm & "'" '????A???[???t???O
    
    ''?g??????
    strSQL = strSQL & " AND K3.KUMIKBN = CMB.CDVAL "
    
    '?K?p??????
    If optApply(0).Value = True Then
        '??V???
        ''?s?A,?@??,ATT?R?[?h,KEY??,KEY-OPT1,KEY-OPT2,?g??????,?g??????,?K?p?J?n??
        strSQL = strSQL & " GROUP BY K3.SHIYUKBN,K3.HONKITYPE,K3.ATTCD,K3.KEYKBN,"
        strSQL = strSQL & " K3.KEYOPT1,K3.KEYOPT2,K3.CPROCTRL_ALARM_FLG,K3.KUMINO, "
        strSQL = strSQL & " K3.TKYMDS "
        
    ElseIf optApply(1).Value = True Then
        '?S??
    ElseIf optApply(2).Value = True Then
        '???t?w??
        strApply = Replace(txtDate, "/", "")
        strSQL = strSQL & " AND ("
        strSQL = strSQL & "   (K3.TKYMDS <= '" & strApply & "'"
        strSQL = strSQL & "   AND '" & strApply & "' <= K3.TKYMDE )"
        strSQL = strSQL & " OR "
        strSQL = strSQL & "   (K3.TKYMDS <= '" & strApply & "'"
        strSQL = strSQL & "   AND K3.TKYMDE = '" & strEndDAte & "') )"
        
    End If
    
    strSQL = strSQL & " ) K2 "
    
    strSQL = strSQL & " WHERE K1.SHIYUKBN = K2.SHIYUKBN "
    strSQL = strSQL & " AND   K1.HONKITYPE = K2.HONKITYPE "
    strSQL = strSQL & " AND   K1.ATTCD = K2.ATTCD "
    strSQL = strSQL & " AND   K1.KEYKBN = K2.KEYKBN "
    strSQL = strSQL & " AND   K1.KEYOPT1 = K2.KEYOPT1 "
    strSQL = strSQL & " AND   K1.KEYOPT2 = K2.KEYOPT2 "
    strSQL = strSQL & " AND   K1.CPROCTRL_ALARM_FLG = K2.CPROCTRL_ALARM_FLG "
    strSQL = strSQL & " AND   K1.KUMINO = K2.KUMINO "
    strSQL = strSQL & " AND   K1.TKYMDS = K2.TKYMDS "
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & " K1.SHIYUKBN "        '?s?A??(????)
    strSQL = strSQL & " ,K1.HONKITYPE "      '?@??(????)
    strSQL = strSQL & " ,K1.ATTCD  "         'ATT?R?[?h(????)
    strSQL = strSQL & " ,K1.KEYKBN  "        'KEY??(????)
    strSQL = strSQL & " ,K1.KEYOPT1  "       'KEY-OPT1(????)
    strSQL = strSQL & " ,K1.KEYOPT2  "       'KEY-OPT2(????)
    strSQL = strSQL & " ,K1.KUMINO  "        '?g??????(????)
    strSQL = strSQL & " ,K1.TKYMDS  "        '?K?p?J?n??(????)
    
    LF_GetSelectSql = strSQL
    
End Function

Private Function LF_CreateCombExcel(ByVal PclsOra As OraDynaset) As Boolean
'--------------------------------------------------------------------------------
' @(f)
' ?@?\??    : Excel???????????
' ?@?\      :
' ????      :[in]    PclsOra    OraDynaset  ?I???N???I?u?W?F?N?g
' ???l    : TRUE?F???? FALSE:?G???[ Boolean
' ?@?\?????@:
'--------------------------------------------------------------------------------
On Error GoTo ErrHandler
    Dim objXLS                  As Xls
    Dim strXLFile               As String           '?p?X???????O
    Dim lngCount                As Long             '?s???J?E???g
    Dim intsheet                As Integer          '?V?[?g??
    Const lngCellMaxLen         As Long = 60000     '???s??
    
    LF_CreateCombExcel = False
    
    '???????èÔ
    ''UPD 2004/10/28 THS T.Y (?????????Excel?o?????) START>>>>>
    strXLFile = gstrClientPath & gstrFileName
'''''   strXLFile = gstrServerPath & gstrFileName
    ''<<<<<END
    
    'Excel?N??
    Set objXLS = New Xls

    '???????
    objXLS.CreateBook strXLFile, , 1
    
    '?V?[?g?????????
    intsheet = 0

    '?P?V?[?g?????????
    objXLS.SheetNo = intsheet
    
    '?o???????????
    mlngKensu = 0
    '?s?J?E???g???????
    lngCount = 0
    
    'Excel?t?@?C??????????????
    Do Until PclsOra.EOF = True
    
        With objXLS
        
            If lngCount = 0 Then
                '?w?b?_??
                .Pos(0, 0).Str = "?X?V??"
                .Pos(1, 0).Str = "?s?A??"
                .Pos(2, 0).Str = "?@??"
                .Pos(3, 0).Str = "?`?s?s?R?[?h"
                .Pos(4, 0).Str = "KEY??"
                .Pos(5, 0).Str = "KEY-OPT1"
                .Pos(6, 0).Str = "KEY-OPT2"
                .Pos(7, 0).Str = "?g??????"
                .Pos(8, 0).Str = "?K?p?J?n??"
                .Pos(9, 0).Str = "?K?p?I????"
                .Pos(10, 0).Str = "?g??????"
                .Pos(11, 0).Str = "?g????OPT 01"
                .Pos(12, 0).Str = "?g????OPT 02"
                .Pos(13, 0).Str = "?g????OPT 03"
                .Pos(14, 0).Str = "?g????OPT 04"
                .Pos(15, 0).Str = "?g????OPT 05"
                .Pos(16, 0).Str = "?g????OPT 06"
                .Pos(17, 0).Str = "?g????OPT 07"
                .Pos(18, 0).Str = "?g????OPT 08"
                .Pos(19, 0).Str = "?g????OPT 09"
                .Pos(20, 0).Str = "?g????OPT 10"
                .Pos(21, 0).Str = "???X?`?F?b?N?s?v?t???O"
                .Pos(22, 0).Str = "ATT???????t???O"
    
            End If
            
            '2?s???~?f?[?^?o??
            lngCount = lngCount + 1
            .Pos(0, lngCount).Str = ""                                         '?X?V??
            .Pos(1, lngCount).Str = Trim(GF_VarToStr(PclsOra![SHIYUKBN]))      '?s?A??
            .Pos(2, lngCount).Str = Trim(GF_VarToStr(PclsOra![HONKITYPE]))     '?@??
            .Pos(3, lngCount).Str = Trim(GF_VarToStr(PclsOra![ATTCD]))         '?`?s?s?R?[?h
            .Pos(4, lngCount).Str = Trim(GF_VarToStr(PclsOra![KEYKBN]))        'KEY??
            .Pos(5, lngCount).Str = Trim(GF_VarToStr(PclsOra![KEYOPT1]))       'KEY-OPT1
            .Pos(6, lngCount).Str = Trim(GF_VarToStr(PclsOra![KEYOPT2]))       'KEY-OPT2
            .Pos(7, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMINO]))        '?g??????
            .Pos(8, lngCount).Str = Trim(GF_VarToStr(PclsOra![TKYMDS]))        '?K?p?J?n??
            .Pos(9, lngCount).Str = Trim(GF_VarToStr(PclsOra![TKYMDE]))        '?K?p?I????
            .Pos(10, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMIKBN]))      '?g??????
            .Pos(11, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI01OPT]))    '?g????OPT 01
            .Pos(12, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI02OPT]))    '?g????OPT 02
            .Pos(13, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI03OPT]))    '?g????OPT 03
            .Pos(14, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI04OPT]))    '?g????OPT 04
            .Pos(15, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI05OPT]))    '?g????OPT 05
            .Pos(16, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI06OPT]))    '?g????OPT 06
            .Pos(17, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI07OPT]))    '?g????OPT 07
            .Pos(18, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI08OPT]))    '?g????OPT 08
            .Pos(19, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI09OPT]))    '?g????OPT 09
            .Pos(20, lngCount).Str = Trim(GF_VarToStr(PclsOra![KUMI10OPT]))    '?g????OPT 10
            .Pos(21, lngCount).Str = Trim(GF_VarToStr(PclsOra![CAGENCY_NOCHECK_FLG]))    '???X?`?F?b?N?s?v?t???O
            .Pos(22, lngCount).Str = Trim(GF_VarToStr(PclsOra![ATTFLG]))    'ATT???????t???O
            mlngKensu = mlngKensu + 1
           
            PclsOra.MoveNext
       End With
   
        '???????????60,000??(?w?b?_???????)????
        If lngCount > lngCellMaxLen Then
'2006/04/13 THS Sugawara 6????????f?[?^?????V?[?g??o????????s?????C?? Add >>>>>
            '?V?[?g?????C???N???????g
            intsheet = intsheet + 1
'<<<<<
            
            '?s???????
            lngCount = 0
            '?V?[?g????
            objXLS.AddSheet 1
            '?V?[?g??I??
'2006/04/13 THS Sugawara 6????????f?[?^?????V?[?g??o????????s?????C?? Update >>>>>
'            objXLS.SheetNo = intsheet + 1
            objXLS.SheetNo = intsheet
'<<<<<
        End If
        
    Loop
    
    '?G?N?Z???N???[?Y
    objXLS.CloseBook
    
    DoEvents
    
    '?Z??????????
    Call LS_ExcelFormat(strXLFile)

    LF_CreateCombExcel = True
    
    Exit Function
    
ErrHandler:

Call GS_ErrorHandler("LF_CreateCombExcel")

Err.Clear
End Function

Private Sub LS_ExcelFormat(ByVal strFName As String)
'--------------------------------------------------------------------------------
' @(f)
' ?@?\??    : Excel???????
' ?@?\?T?v  : Excel???????
' ????      :  strFName As String          ?t?@?C???p?X
'
' ???l   :
'--------------------------------------------------------------------------------

On Error GoTo ErrHandler
    Dim objExcel As excel.Application
    Dim objBook As excel.Workbook
    Dim objSheet As excel.Worksheet
    Dim intCnt As Integer               '?V?[?g??J?E???g
    
    Const lngMaxRow As Long = 65536     '?s????
    Const intRow As Integer = 23        '??
    
    Set objBook = Nothing
    '?G?N?Z???N??
    Set objExcel = New excel.Application
    '?u?b?N?????
    Set objBook = objExcel.Workbooks.Open(strFName)
    '?V?[?g????????????????
    For intCnt = 1 To objBook.Worksheets.Count
        '?V?[?g??I??
        Set objSheet = objBook.Worksheets(intCnt)
        '????????
        objSheet.Range(objSheet.Cells(1, 1), objSheet.Cells(lngMaxRow, intRow)).NumberFormat = "@"
        '?t?H???g???
        objSheet.Range(objSheet.Cells(1, 1), objSheet.Cells(lngMaxRow, intRow)).Font.Name = "MS ????"
        '?????T?C?Y???
        objSheet.Range(objSheet.Cells(1, 1), objSheet.Cells(lngMaxRow, intRow)).Font.Size = 9
        '???l???
        objSheet.Cells().HorizontalAlignment = xlHAlignLeft
        '????K??
        objSheet.Columns("A:W").AutoFit
        '?Y?[?????(100%)
        objExcel.ActiveWindow.Zoom = 100
    Next

    '?u?b?N?N???[?Y
    objBook.Close True
    '?G?N?Z???I??
    objExcel.Quit
    
    Set objSheet = Nothing
    Set objBook = Nothing
    Set objExcel = Nothing

    Exit Sub
    
ErrHandler:

    objBook.Close True
    objExcel.Quit

    Set objSheet = Nothing
    Set objBook = Nothing
    Set objExcel = Nothing
    Call GS_ErrorHandler("LS_ExcelFormat")

Err.Clear
End Sub

