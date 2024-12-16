VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOraClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) OraClass.cls  ver1.00 ( 2000/09/15 T.Ooka )
'------------------------------------------------------------------------------
' @(s)
'   プロジェクト名 : TLFﾌﾟﾛｼﾞｪｸﾄ
'   モジュール名　 : clsOraClass
'   ファイル名　　 : OraClass.cls
'   バージョン　　 : 1.00
'   機能説明　　　 : ｵﾗｸﾙ関連ｸﾗｽﾓｼﾞｭｰﾙ
'   作成者　　　　 : T.OOKA
'   作成日　　　　 : 2000/09/15
'   修正履歴　　　 : 2000/09/26  [Method]SetRollbackSegment ﾛｰﾙﾊﾞｯｸｾｸﾞﾒﾝﾄ変更
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 環境宣言
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------
'内部変数宣言
'------------------------------------------------------------
'Move to Bindpara
Private m_OraDataBase_Bind      As Object
Private m_ParaName_Bind         As String
Private m_IOType_Bind           As Integer
Private m_Value_Bind            As Variant
Private m_PType_Bind            As Integer
Private m_MinSize_Bind          As Integer

'Move to Bindparas
Private ColBindParas            As New Collection
Private m_OraDataBase_Binds     As Object

'Move to StoredCall から移行
Private m_OraDataBase_Strcall   As Object
Private m_ErrText_Strcall       As String
Private m_ErrCode_Strcall       As Integer

'============================================================
'
'  =  ORACLE共通関連クラスモジュール - 1  =
'
'============================================================
'============================================================
'バインド変数関連プロパティ
'============================================================
'------------------------------------------------------------
'データベース名のプロパティ
'------------------------------------------------------------
Public Property Set OraDataBase_Bind(P As Object)
    Set m_OraDataBase_Bind = P
End Property

Public Property Get OarDataBase_Bind() As Object
    Set OraDataBase_Bind = m_OraDataBase_Bind
End Property

'------------------------------------------------------------
'変数名のプロパティ
'------------------------------------------------------------
Public Property Let ParaName_Bind(P As String)
    m_ParaName_Bind = P
End Property

Public Property Get ParaName_Bind() As String
    ParaName_Bind = m_ParaName_Bind
End Property

'------------------------------------------------------------
'変数の入力・出力タイプのプロパティ
'------------------------------------------------------------
Public Property Let IOType_Bind(P As Integer)
    m_IOType_Bind = P
End Property

Public Property Get IOType_Bind() As Integer
    IOType_Bind = m_IOType_Bind
End Property

'------------------------------------------------------------
'変数値のプロパティ
'------------------------------------------------------------
Public Property Let Value_Bind(P As Variant)
    m_Value_Bind = P
End Property

Public Property Get Value_Bind() As Variant
    Value_Bind = m_Value_Bind
End Property

'------------------------------------------------------------
'変数のデータ型のプロパティ
'------------------------------------------------------------
Public Property Let PType_Bind(P As Integer)
    m_PType_Bind = P
End Property

Public Property Get PType_Bind() As Integer
    PType_Bind = m_PType_Bind
End Property

'------------------------------------------------------------
'変数のデータ型のプロパティ
'------------------------------------------------------------
Public Property Let MinSize_Bind(P As Integer)
    m_MinSize_Bind = P
End Property

Public Property Get MinSize_Bind() As Integer
    MinSize_Bind = m_MinSize_Bind
End Property

'============================================================
'メソッド 関連
'============================================================
'------------------------------------------------------------
'バインド変数の登録
'------------------------------------------------------------
Public Function RegBindPara_Bind() As Integer

    m_OraDataBase_Bind.Parameters.Add m_ParaName_Bind, m_Value_Bind, m_IOType_Bind
    m_OraDataBase_Bind.Parameters(ParaName_Bind).serverType = m_PType_Bind
    If m_MinSize_Bind > -1 Then
        m_OraDataBase_Bind.Parameters(ParaName_Bind).MinimumSize = m_MinSize_Bind
    End If
    RegBindPara_Bind = 0
End Function

'------------------------------------------------------------
'バインド変数の削除
'------------------------------------------------------------
Public Function UnRegBindPara_Bind() As Integer
    
    m_OraDataBase_Bind.Parameters.Remove m_ParaName_Bind
    UnRegBindPara_Bind = 0
End Function

'------------------------------------------------------------
'バインド変数の値を取得
'------------------------------------------------------------
Public Function GetValue_Bind() As Variant
    GetValue_Bind = m_OraDataBase_Bind.Parameters(m_ParaName_Bind).Value
End Function

'============================================================
'
'  =  ORACLE共通関連クラスモジュール - 2  =
'
'============================================================
Public Property Set OraDataBase_Binds(P As Object)
    Set m_OraDataBase_Binds = P
End Property

Public Property Get OraDataBase_Binds() As Object
    Set OraDataBase_Binds = m_OraDataBase_Binds
End Property

'------------------------------------------------------------
'Count プロパティの獲得関数
'------------------------------------------------------------
Property Get Count_Binds() As Integer
    Count_Binds = ColBindParas.Count_Binds
End Property

'------------------------------------------------------------
'Item 取得メソッド
'------------------------------------------------------------
Public Function Item_Binds(ByVal Index As Variant) As clsOraClass
    Set Item_Binds = ColBindParas.Item(Index)
End Function

'------------------------------------------------------------
'バインド変数Add メソッド
'------------------------------------------------------------
Public Function Add_Binds(ByVal IOType As Integer, _
                          ByVal PType As Integer, _
                          ByVal ParaName As Variant, _
                          Optional ByVal Value As Variant, _
                          Optional ByVal MinSize As Integer = -1) As clsOraClass

    Dim Key As Variant
    Dim ObjBindPara As New clsOraClass

    ObjBindPara.ParaName_Bind = ParaName
    ObjBindPara.PType_Bind = PType
    ObjBindPara.IOType_Bind = IOType
    ObjBindPara.MinSize_Bind = MinSize

    '値が渡ってきているか
    If IsMissing(Value) Then
        Select Case PType
        Case ORATYPE_VARCHAR2, ORATYPE_STRING, ORATYPE_VARCHAR
            ObjBindPara.Value_Bind = ""
        Case ORATYPE_CHAR, ORATYPE_CHARZ
            ObjBindPara.Value_Bind = ""
        Case ORATYPE_NUMBER, ORATYPE_SINT, ORATYPE_UINT
            ObjBindPara.Value_Bind = 0
        Case ORATYPE_FLOAT
            ObjBindPara.Value_Bind = 0
        Case ORATYPE_DATE
            ObjBindPara.Value_Bind = 0
        Case Else
            ObjBindPara.Value_Bind = 0
        End Select
    Else
        ObjBindPara.Value_Bind = Value
    End If

    ' 新しいオブジェクトをプライベートCollectionオブジェクトに追加
    Key = ParaName
    ColBindParas.Add ObjBindPara, Key

    'データベースオブジェクトを設定する
    Set ColBindParas(Key).OraDataBase_Bind = OraDataBase_Binds

    'データベースのパラメータBIND処理を行う
    ColBindParas(Key).RegBindPara_Bind

    ' 新しいオブジェクトに対する参照を返す
    Set Add_Binds = ObjBindPara

End Function

'------------------------------------------------------------
'バインド変数 Remove メソッド
'------------------------------------------------------------
Public Sub Remove(Index As Variant)

    ColBindParas(Index).UnRegBindPara
    ColBindParas.Remove Index
End Sub

'------------------------------------------------------------
'バインド変数 Remove All メソッド
'------------------------------------------------------------
Public Sub RemoveAll()
    Dim ii As Integer
    Dim Cnt As Integer
    Dim PName() As String

    Cnt = ColBindParas.Count
    ReDim PName(Cnt)

    For ii = 1 To Cnt
        PName(ii) = ColBindParas(ii).ParaName_Bind
    Next
        
    For ii = 1 To Cnt
        ColBindParas(PName(ii)).UnRegBindPara_Bind
        ColBindParas.Remove PName(ii)
    Next

End Sub

'============================================================
'
'  =  ORACLE共通関連クラスモジュール - 3  =
'
'============================================================
'============================================================
'プロパティ取得／設定関数
'============================================================
Public Property Set OraDataBase_Strcall(P As Object)
    Set OraDataBase_Binds = P
    Set m_OraDataBase_Strcall = P
End Property

Public Property Get OraDataBase_Strcall() As Object
    Set OraDataBase_Strcall = m_OraDataBase_Strcall
End Property

Public Property Get ErrText_Strcall() As String
    ErrText_Strcall = m_OraDataBase_Strcall.LastServerErrText
End Property

Public Property Get ErrCode_Strcall() As Integer
    ErrCode_Strcall = m_OraDataBase_Strcall.LastServerErr
End Property

'============================================================
'メソッド宣言
'============================================================
'------------------------------------------------------------
'SQLエラー情報の初期化メソッド
'------------------------------------------------------------
Public Sub ErrReset_Strcall()
    OraDataBase_Strcall.LastServerErrReset
End Sub

'------------------------------------------------------------
'SQL文の実行メソッド
'------------------------------------------------------------
Public Sub ExecSql_Strcall(Sql As String)
'    Dim Ora_SqlStmt         As OraSqlStmt
'    Set Ora_SqlStmt = m_OraDataBase_Strcall.CreateSql(Sql, &H1&)
    m_OraDataBase_Strcall.dbExecuteSQL (Sql)
'    OraDataBase_Strcall
End Sub

'------------------------------------------------------------
'ロールバックセグメントの指定メソッド
'------------------------------------------------------------
Public Sub SetRollbackSegment()
    On Error Resume Next
    m_OraDataBase_Strcall.BeginTrance
    m_OraDataBase_Strcall.dbExecuteSQL ("set transaction use rollback segment LRBS")
    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
    m_MinSize_Bind = -1
End Sub
