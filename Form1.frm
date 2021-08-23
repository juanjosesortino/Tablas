VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador de Registros"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionador de Empresa"
      Height          =   765
      Index           =   0
      Left            =   30
      TabIndex        =   19
      Top             =   0
      Width           =   10500
      Begin MSDataListLib.DataCombo dtcEmpresas 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6735
      Left            =   30
      TabIndex        =   6
      Top             =   780
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   529
      TabCaption(0)   =   "Registros"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtcClave5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dtcClave4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dtcClave3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtcClave2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtcClave1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "StatusBar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Buscador"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListViewFind"
      Tab(1).Control(1)=   "Frame1(1)"
      Tab(1).ControlCount=   2
      Begin VB.OptionButton Option5 
         Height          =   195
         Left            =   -71190
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   810
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Height          =   195
         Left            =   -72210
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   810
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buscar"
         Height          =   885
         Index           =   1
         Left            =   -74940
         TabIndex        =   7
         Top             =   360
         Width           =   10260
         Begin VB.TextBox txtTexto 
            Height          =   405
            Left            =   120
            TabIndex        =   9
            Top             =   300
            Width           =   8415
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   525
            Left            =   8760
            TabIndex        =   8
            Top             =   240
            Width           =   1305
         End
      End
      Begin MSComctlLib.ListView ListViewTablas 
         Height          =   6150
         Left            =   -74940
         TabIndex        =   12
         Top             =   750
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   10848
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TABLA"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "MB      "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "EXTENTS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "INITIAL_EXT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "NEXT_EXT"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "MAX_EXT"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewConexion 
         Height          =   6600
         Left            =   -74970
         TabIndex        =   13
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "OSuser"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Username"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Machine"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Program"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Serial"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewTablespaces 
         Height          =   6600
         Left            =   -74970
         TabIndex        =   14
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Tablespace"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "MB Tamaño"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "MB Usados"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "MB Libres"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fichero de datos"
            Object.Width           =   5468
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewSQL 
         Height          =   6600
         Left            =   -74970
         TabIndex        =   15
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Programa"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Fecha/Hora"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Usuario"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SQL"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   4260
         Left            =   150
         TabIndex        =   5
         Top             =   2580
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7514
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Esquema"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Empresa"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Importacion"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Version"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewFind 
         Height          =   5340
         Left            =   -74940
         TabIndex        =   16
         Top             =   1320
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   9419
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clave2"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clave3"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Clave4"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Clave5"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Valor"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.StatusBar StatusBar 
         Height          =   375
         Left            =   270
         TabIndex        =   18
         Top             =   7320
         Width           =   10170
         _ExtentX        =   17939
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   4
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3704
               MinWidth        =   3704
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
               MinWidth        =   3528
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
               MinWidth        =   5292
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   8819
               MinWidth        =   8819
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcClave1 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   480
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcClave2 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   900
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcClave3 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   1290
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcClave4 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   1680
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcClave5 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   2100
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71310
         TabIndex        =   17
         Top             =   420
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 Option Explicit

Private SQL               As String
Private itmX              As ListItem
Private itmT              As ListItem
Private sngCoordenadaX    As Single
Private sngCoordenadaY    As Single
Private iorden            As Integer
Private iordenTablas      As Integer
Private rstGlobal         As ADODB.Recordset
Private cnnGlobal         As ADODB.Connection
Private fs                As Object
Private fNumber1          As Integer

Private Sub cmdBuscar_Click()
Dim rst1    As ADODB.Recordset

   On Error GoTo GestErr

   StatusBar.Panels(1).Text = "Buscando..."
   If dtcEmpresas.BoundText = "ALL" Then
      MsgBox "Debe seleccionar una empresa"
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
            
   Set cnnGlobal = New ADODB.Connection
   cnnGlobal.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=" & dtcEmpresas.BoundText & ";Data Source=BASE"
   cnnGlobal.Open
   
   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
   
   Set rst1 = New ADODB.Recordset
   rst1.CursorLocation = adUseClient
   rst1.LockType = adLockReadOnly
   rst1.CursorType = adOpenStatic
   
   SQL = " SELECT TAB_CLAVE1, TAB_CLAVE2, TAB_CLAVE3, TAB_CLAVE4, TAB_CLAVE5, TAB_VALOR "
   SQL = SQL & "  FROM TABLAS "
   SQL = SQL & " WHERE UPPER (TAB_CLAVE1) LIKE '%" & UCase(txtTexto.Text) & "%' "
   SQL = SQL & "    OR UPPER (TAB_CLAVE2) LIKE '%" & UCase(txtTexto.Text) & "%' "
   SQL = SQL & "    OR UPPER (TAB_CLAVE3) LIKE '%" & UCase(txtTexto.Text) & "%' "
   SQL = SQL & "    OR UPPER (TAB_CLAVE4) LIKE '%" & UCase(txtTexto.Text) & "%' "
   SQL = SQL & "    OR UPPER (TAB_CLAVE5) LIKE '%" & UCase(txtTexto.Text) & "%' "
   SQL = SQL & "    OR UPPER (TAB_VALOR)  LIKE '%" & UCase(txtTexto.Text) & "%' "
   rstGlobal.Open SQL, cnnGlobal
      
   If rstGlobal.RecordCount > 0 Then
      rstGlobal.MoveFirst
      ListViewFind.ListItems.Clear
      Do While Not rstGlobal.EOF
         Set itmX = ListViewFind.ListItems.Add
         
         itmX.SubItems(1) = IIf(IsNull(rstGlobal("TAB_CLAVE1").Value), "", rstGlobal("TAB_CLAVE1").Value)
         itmX.SubItems(2) = IIf(IsNull(rstGlobal("TAB_CLAVE2").Value), "", rstGlobal("TAB_CLAVE2").Value)
         itmX.SubItems(3) = IIf(IsNull(rstGlobal("TAB_CLAVE3").Value), "", rstGlobal("TAB_CLAVE3").Value)
         itmX.SubItems(4) = IIf(IsNull(rstGlobal("TAB_CLAVE4").Value), "", rstGlobal("TAB_CLAVE4").Value)
         itmX.SubItems(5) = IIf(IsNull(rstGlobal("TAB_CLAVE5").Value), "", rstGlobal("TAB_CLAVE5").Value)
         itmX.SubItems(6) = IIf(IsNull(rstGlobal("TAB_VALOR").Value), "", rstGlobal("TAB_VALOR").Value)
         rstGlobal.MoveNext
      Loop
   Else
      ListViewFind.ListItems.Clear
   End If
   
   Screen.MousePointer = vbNormal
   StatusBar.Panels(1).Text = ""
      
   Exit Sub

GestErr:
   Screen.MousePointer = vbNormal
   MsgBox "[SetListviewEmpresas]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub dtcEmpresas_Click(Area As Integer)
   SetListviewEmpresas dtcEmpresas.BoundText
   ObtenerValor
End Sub

Private Sub Form_Load()
   Inicio
End Sub

Private Sub Inicio()
Dim rst1    As ADODB.Recordset

   On Error GoTo GestErr

   StatusBar.Panels(1).Text = "Inicializando..."
   Screen.MousePointer = vbHourglass
            
   SetComboEmpresas
   SetClave1
   SetListviewEmpresas
   
   Screen.MousePointer = vbNormal
   StatusBar.Panels(1).Text = ""
      
   Exit Sub
GestErr:
   Screen.MousePointer = vbNormal
   MsgBox "[Inicio]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub SetListviewEmpresas(Optional ByVal strEmpresa As String)
Dim rst1    As ADODB.Recordset

   On Error GoTo GestErr

   StatusBar.Panels(1).Text = "Inicializando..."
   Screen.MousePointer = vbHourglass
            
   Set cnnGlobal = New ADODB.Connection
   cnnGlobal.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
   cnnGlobal.Open
   
   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
   
   Set rst1 = New ADODB.Recordset
   rst1.CursorLocation = adUseClient
   rst1.LockType = adLockReadOnly
   rst1.CursorType = adOpenStatic
   
   SQL = " SELECT   DBA_USERS.USERNAME, DBA_USERS.CREATED, EMPRESAS.EMP_DESCRIPCION "
   SQL = SQL & "    FROM DBA_USERS, "
   SQL = SQL & "         EMPRESAS "
   SQL = SQL & "   WHERE USERNAME LIKE 'SYSADMIN_%' "
   SQL = SQL & "     AND SUBSTR (DBA_USERS.USERNAME, 10, 3) = EMPRESAS.EMP_CODIGO_EMPRESA(+) "
   If Len(strEmpresa) > 0 And strEmpresa <> "ALL" Then
      SQL = SQL & " AND USERNAME LIKE '" & strEmpresa & "' "
   End If
   SQL = SQL & "ORDER BY USERNAME desc"
   
   rstGlobal.Open SQL, cnnGlobal
      
   rstGlobal.MoveFirst
   ListView.ListItems.Clear
   Do While Not rstGlobal.EOF
      Set itmX = ListView.ListItems.Add
      
      itmX.SubItems(1) = IIf(IsNull(rstGlobal("USERNAME").Value), "", rstGlobal("USERNAME").Value)
      itmX.SubItems(2) = IIf(IsNull(rstGlobal("EMP_DESCRIPCION").Value), "", rstGlobal("EMP_DESCRIPCION").Value)
      itmX.SubItems(3) = IIf(IsNull(rstGlobal("CREATED").Value), "", rstGlobal("CREATED").Value)

      
      SQL = "SELECT NRO_VERSION FROM " & rstGlobal("USERNAME").Value & ".VERSION_PRODUCTO"
      Set rst1 = New ADODB.Recordset
      rst1.CursorLocation = adUseClient
      rst1.LockType = adLockReadOnly
      rst1.CursorType = adOpenStatic
      
      On Error Resume Next
      rst1.Open SQL, cnnGlobal
      If rst1.RecordCount > 0 Then
         rst1.MoveFirst
         itmX.SubItems(4) = IIf(IsNull(rst1("NRO_VERSION").Value), "", rst1("NRO_VERSION").Value)
      End If
      If Not rst1 Is Nothing Then
         If rst1.State <> adStateClosed Then rst1.Close
      End If
      Set rst1 = Nothing
      Err.Clear
      On Error GoTo GestErr

      rstGlobal.MoveNext
   Loop
   
   Screen.MousePointer = vbNormal
   StatusBar.Panels(1).Text = ""
      
   Exit Sub

GestErr:
   Screen.MousePointer = vbNormal
   MsgBox "[SetListviewEmpresas]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub SetComboEmpresas()
Dim Rst  As ADODB.Recordset

   Set cnnGlobal = New ADODB.Connection
   cnnGlobal.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
   cnnGlobal.Open
   
   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
   
   SQL = " SELECT   'ALL' USERNAME, 'Todas las Empresas' EMP_DESCRIPCION "
   SQL = SQL & "    FROM DUAL "
   SQL = SQL & "UNION ALL "
'   SQL = SQL & "SELECT   DBA_USERS.USERNAME, 'SYSADMIN' "
'   SQL = SQL & "    FROM DBA_USERS, EMPRESAS "
'   SQL = SQL & "   WHERE USERNAME LIKE 'SYSADMIN' "
'   SQL = SQL & "     AND SUBSTR (DBA_USERS.USERNAME, 10, 3) = EMPRESAS.EMP_CODIGO_EMPRESA(+) "
'   SQL = SQL & "UNION ALL "
   SQL = SQL & "SELECT   DBA_USERS.USERNAME, SUBSTR (DBA_USERS.USERNAME, 10, 3) || ' ' || EMPRESAS.EMP_DESCRIPCION "
   SQL = SQL & "    FROM DBA_USERS, EMPRESAS "
   SQL = SQL & "   WHERE USERNAME LIKE 'SYSADMIN%' "
   SQL = SQL & "     AND SUBSTR (DBA_USERS.USERNAME, 10, 3) = EMPRESAS.EMP_CODIGO_EMPRESA(+) "
   SQL = SQL & "     AND EMPRESAS.EMP_DESCRIPCION IS NOT NULL "
   SQL = SQL & "ORDER BY 1 "

   rstGlobal.Open SQL, cnnGlobal
   
   Do While Not rstGlobal.EOF
      Set dtcEmpresas.RowSource = rstGlobal
      dtcEmpresas.BoundColumn = "USERNAME"
      dtcEmpresas.ListField = "EMP_DESCRIPCION"
      
      rstGlobal.MoveNext
   Loop
   dtcEmpresas.BoundText = "ALL"
End Sub
Private Sub SetClave1()
Dim Rst  As ADODB.Recordset

   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
   
   SQL = " SELECT DISTINCT TAB_CLAVE1"
   SQL = SQL & "           FROM TABLAS "
   SQL = SQL & "       ORDER BY TAB_CLAVE1"

   rstGlobal.Open SQL, cnnGlobal
   
   Do While Not rstGlobal.EOF
      Set dtcClave1.RowSource = rstGlobal
      dtcClave1.BoundColumn = "TAB_CLAVE1"
      dtcClave1.ListField = "TAB_CLAVE1"

      rstGlobal.MoveNext
   Loop
   rstGlobal.MoveFirst
   dtcClave1.BoundText = rstGlobal("TAB_CLAVE1").Value
   ObtenerValor
End Sub
Private Sub SetClave2()
Dim Rst  As ADODB.Recordset

   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
   
   SQL = " SELECT DISTINCT TAB_CLAVE2 "
   SQL = SQL & "           FROM TABLAS "
   SQL = SQL & "          WHERE TAB_CLAVE1 = '" & dtcClave1.BoundText & "'"
   SQL = SQL & "       ORDER BY TAB_CLAVE2 "


   rstGlobal.Open SQL, cnnGlobal
   
   If rstGlobal.RecordCount = 1 Then
      Set dtcClave2.RowSource = Nothing
      dtcClave2.BoundText = ""
      Set dtcClave3.RowSource = Nothing
      dtcClave3.BoundText = ""
      Set dtcClave4.RowSource = Nothing
      dtcClave4.BoundText = ""
      Set dtcClave5.RowSource = Nothing
      dtcClave5.BoundText = ""
      ObtenerValor
      Exit Sub
   End If
   
   Do While Not rstGlobal.EOF
      Set dtcClave2.RowSource = rstGlobal
      dtcClave2.BoundColumn = "TAB_CLAVE2"
      dtcClave2.ListField = "TAB_CLAVE2"

      rstGlobal.MoveNext
   Loop
   rstGlobal.MoveFirst
   dtcClave2.BoundText = rstGlobal("TAB_CLAVE2").Value
   ObtenerValor
End Sub
Private Sub SetClave3()
Dim Rst  As ADODB.Recordset

   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
   
   SQL = " SELECT DISTINCT TAB_CLAVE3 "
   SQL = SQL & "           FROM TABLAS "
   SQL = SQL & "          WHERE TAB_CLAVE1 = '" & dtcClave1.BoundText & "'"
   SQL = SQL & "          AND   TAB_CLAVE2 = '" & dtcClave2.BoundText & "'"
   SQL = SQL & "       ORDER BY TAB_CLAVE3 "


   rstGlobal.Open SQL, cnnGlobal
   
   If rstGlobal.RecordCount = 1 Or dtcClave2.BoundText = "" Then
      Set dtcClave3.RowSource = Nothing
      dtcClave3.BoundText = ""
      Set dtcClave4.RowSource = Nothing
      dtcClave4.BoundText = ""
      Set dtcClave5.RowSource = Nothing
      dtcClave5.BoundText = ""
      ObtenerValor
      Exit Sub
   End If
   
   Do While Not rstGlobal.EOF
      Set dtcClave3.RowSource = rstGlobal
      dtcClave3.BoundColumn = "TAB_CLAVE3"
      dtcClave3.ListField = "TAB_CLAVE3"

      rstGlobal.MoveNext
   Loop
   rstGlobal.MoveFirst
   dtcClave3.BoundText = rstGlobal("TAB_CLAVE3").Value
   ObtenerValor
End Sub
Private Sub SetClave4()
Dim Rst  As ADODB.Recordset

   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
   
   SQL = " SELECT DISTINCT TAB_CLAVE4 "
   SQL = SQL & "           FROM TABLAS "
   SQL = SQL & "          WHERE TAB_CLAVE1 = '" & dtcClave1.BoundText & "'"
   SQL = SQL & "          AND   TAB_CLAVE2 = '" & dtcClave2.BoundText & "'"
   SQL = SQL & "          AND   TAB_CLAVE3 = '" & dtcClave3.BoundText & "'"
   SQL = SQL & "       ORDER BY TAB_CLAVE4 "


   rstGlobal.Open SQL, cnnGlobal
   
   If rstGlobal.RecordCount = 1 Or dtcClave3.BoundText = "" Then
      Set dtcClave4.RowSource = Nothing
      dtcClave4.BoundText = ""
      Set dtcClave5.RowSource = Nothing
      dtcClave5.BoundText = ""
      ObtenerValor
      Exit Sub
   End If
   
   Do While Not rstGlobal.EOF
      Set dtcClave4.RowSource = rstGlobal
      dtcClave4.BoundColumn = "TAB_CLAVE4"
      dtcClave4.ListField = "TAB_CLAVE4"

      rstGlobal.MoveNext
   Loop
   rstGlobal.MoveFirst
   dtcClave4.BoundText = rstGlobal("TAB_CLAVE4").Value
   ObtenerValor
End Sub
Private Sub SetClave5()
Dim Rst  As ADODB.Recordset

   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
   
   SQL = " SELECT DISTINCT TAB_CLAVE5 "
   SQL = SQL & "           FROM TABLAS "
   SQL = SQL & "          WHERE TAB_CLAVE1 = '" & dtcClave1.BoundText & "'"
   SQL = SQL & "          AND   TAB_CLAVE2 = '" & dtcClave2.BoundText & "'"
   SQL = SQL & "          AND   TAB_CLAVE3 = '" & dtcClave3.BoundText & "'"
   SQL = SQL & "          AND   TAB_CLAVE4 = '" & dtcClave4.BoundText & "'"
   SQL = SQL & "       ORDER BY TAB_CLAVE5 "


   rstGlobal.Open SQL, cnnGlobal
   
   If rstGlobal.RecordCount = 1 Or dtcClave4.BoundText = "" Then
      Set dtcClave5.RowSource = Nothing
      dtcClave5.BoundText = ""
      ObtenerValor
      Exit Sub
   End If
   
   Do While Not rstGlobal.EOF
      Set dtcClave5.RowSource = rstGlobal
      dtcClave5.BoundColumn = "TAB_CLAVE5"
      dtcClave5.ListField = "TAB_CLAVE5"

      rstGlobal.MoveNext
   Loop
   rstGlobal.MoveFirst
   dtcClave5.BoundText = rstGlobal("TAB_CLAVE5").Value
   ObtenerValor
End Sub
Private Sub dtcClave1_Change()
   SetClave2
End Sub
Private Sub dtcClave2_Change()
   SetClave3
End Sub
Private Sub dtcClave3_Change()
   SetClave4
End Sub
Private Sub dtcClave4_Change()
   SetClave5
End Sub
Private Sub ObtenerValor()
Dim rstGlobal  As ADODB.Recordset
   
   On Error Resume Next
   
   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
      
   For Each itmX In ListView.ListItems
      SQL = " SELECT TAB_EMPRESA, TAB_CLAVE1, TAB_CLAVE2, TAB_CLAVE3, TAB_CLAVE4, TAB_CLAVE5, TAB_VALOR, TAB_COMENTARIO, TAB_TABLA_RESERVADA, TAB_VALOR_PREDETERMINADO, "
      SQL = SQL & "       TAB_VALORES_POSIBLES "
      SQL = SQL & "  FROM " & itmX.SubItems(1) & ".TABLAS "
      If Len(dtcClave1.BoundText) > 0 Then
         SQL = SQL & "          WHERE TAB_CLAVE1 = '" & dtcClave1.BoundText & "'"
      End If
      If Len(dtcClave2.BoundText) > 0 Then
         SQL = SQL & "          AND   TAB_CLAVE2 = '" & dtcClave2.BoundText & "'"
      End If
      If Len(dtcClave3.BoundText) > 0 Then
         SQL = SQL & "          AND   TAB_CLAVE3 = '" & dtcClave3.BoundText & "'"
      End If
      If Len(dtcClave4.BoundText) > 0 Then
         SQL = SQL & "          AND   TAB_CLAVE4 = '" & dtcClave4.BoundText & "'"
      End If
      If Len(dtcClave5.BoundText) > 0 Then
         SQL = SQL & "          AND   TAB_CLAVE5 = '" & dtcClave5.BoundText & "'"
      End If
      
      rstGlobal.Open SQL, cnnGlobal

      If rstGlobal.RecordCount > 0 Then
         itmX.SubItems(5) = IIf(IsNull(rstGlobal("TAB_VALOR").Value), "", rstGlobal("TAB_VALOR").Value)
      Else
         itmX.SubItems(5) = ""
      End If
      rstGlobal.Close
   Next itmX
         
End Sub

Private Sub ListViewFind_DblClick()
Dim itmX              As ListItem

   On Error GoTo GestErr
      
   Set itmX = ListViewFind.HitTest(sngCoordenadaX, sngCoordenadaY)
   
   If Not itmX Is Nothing Then
      
      SSTab.Tab = 0
      
      If Len(itmX.SubItems(1)) > 0 Then
         dtcClave1.BoundText = itmX.SubItems(1)
      End If
      If Len(itmX.SubItems(2)) > 0 Then
         dtcClave2.BoundText = itmX.SubItems(2)
      End If
      If Len(itmX.SubItems(3)) > 0 Then
         dtcClave3.BoundText = itmX.SubItems(3)
      End If
      If Len(itmX.SubItems(4)) > 0 Then
         dtcClave4.BoundText = itmX.SubItems(4)
      End If
      If Len(itmX.SubItems(5)) > 0 Then
         dtcClave5.BoundText = itmX.SubItems(5)
      End If
       
   End If
   
   Exit Sub

GestErr:
   Screen.MousePointer = vbNormal
   MsgBox "[ListViewFind_DblClick]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub ListViewFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   sngCoordenadaX = X
   sngCoordenadaY = Y
End Sub
