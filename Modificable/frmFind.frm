VERSION 5.00
Object = "{B97E3E11-CC61-11D3-95C0-00C0F0161F05}#162.0#0"; "ALGControls.ocx"
Begin VB.Form frmFind 
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin ALGControls.QueryDB QueryDB1 
      Height          =   2040
      Left            =   450
      TabIndex        =   0
      Top             =   270
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   3598
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ErrorLog                 As ErrType

Private mvarControlData          As DataShare.udtControlData
Private mvarError                As String
Private mvarTableField           As String
Private mvarTitulo               As String                'titulo de la grilla
Private mvarTituloColumnas       As String                'titulo de c/columna
Private mvarAnchoColumnas        As String                'ancho de c/columna
Private mvarSQL                  As String                'query
Private mvarWhere                As String                'clausula where para filtrar la query del diccionario
Private mvarFormatoColumnas      As String                'formato de c/columna

Private Sub Form_Load()
   DoEvents
End Sub

Private Sub Form_Resize()

   If mvarMDIForm.WindowState = vbMinimized Then Exit Sub
   
'   With QueryDB1
'     .Left = 100
'     .Top = 100
'     .Width = Me.ScaleWidth - 200
'     .Height = Me.ScaleHeight - 200
'   End With
'
   CenterForm Me
   
End Sub

Private Sub QueryDB1_Hide()
   Me.Hide
End Sub

Private Sub QueryDB1_ItemSelected()
  Me.Hide
End Sub

Public Property Let ControlData(ByRef vData As DataShare.udtControlData)
    mvarControlData = vData
    QueryDB1.ControlData = mvarControlData
    
   With ErrorLog
      .Form = Me.Name
      .Empresa = mvarControlData.Empresa
   End With
    
End Property

Public Property Get ControlData() As DataShare.udtControlData
    ControlData = mvarControlData
End Property

Public Property Let TableField(vData As String)
   mvarTableField = vData
   QueryDB1.TableField = mvarTableField
End Property
Public Property Get TableField() As String
   TableField = mvarTableField
End Property

Public Property Get Titulo() As String
  Titulo = mvarTitulo
End Property
Public Property Let Titulo(ByVal vData As String)
  mvarTitulo = vData
End Property

Public Property Get TituloColumnas() As String
  TituloColumnas = mvarTituloColumnas
End Property

Public Property Let TituloColumnas(ByVal vData As String)
  mvarTituloColumnas = vData
End Property

Public Property Get FormatoColumnas() As String
  FormatoColumnas = mvarFormatoColumnas
End Property

Public Property Let FormatoColumnas(ByVal vData As String)
  mvarFormatoColumnas = vData
End Property
   
Public Property Get AnchoColumnas() As String
  AnchoColumnas = mvarAnchoColumnas
End Property

Public Property Let AnchoColumnas(ByVal vData As String)
  mvarAnchoColumnas = vData
End Property
   
Public Property Get SQL() As String
  SQL = mvarSQL
End Property

Public Property Let SQL(ByVal vData As String)
  mvarSQL = vData
End Property

Public Property Get Where() As String
  Where = mvarWhere
End Property

Public Property Let Where(ByVal vData As String)
  mvarWhere = vData
End Property

Public Property Get Error() As String
  Error = mvarError
End Property

'***********************************************************
'La propiedad tag se pasa con el valor "%%SubQuery%%" cuando
'se quiere evitar el sComposer
'ya que este no permite ejecutar querys con SubQuerys
'***********************************************************
Public Sub LoadData()
   
   On Error GoTo GestErr
   
   mvarError = NullString
   
   With QueryDB1
      .TableField = mvarTableField
      .Titulo = mvarTitulo
      .TituloColumnas = mvarTituloColumnas
      .SQL = mvarSQL
      .Where = mvarWhere
      .FormatoColumnas = mvarFormatoColumnas
      .AnchoColumnas = mvarAnchoColumnas
      .Tag = Me.Tag

      .Refresh
   End With

   Exit Sub

GestErr:
   
   mvarError = "Error"
   
   LoadError ErrorLog, "[frmFind] ShowData"
   ShowErrMsg ErrorLog
   
   Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmFind = Nothing

End Sub

