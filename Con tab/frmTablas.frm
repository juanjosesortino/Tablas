VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTablas 
   Caption         =   "Registro del Sistema"
   ClientHeight    =   8655
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10980
   Icon            =   "frmTablas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8385
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18865
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   8415
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   529
      TabCaption(0)   =   "Registros"
      TabPicture(0)   =   "frmTablas.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MDIExtend1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgSplitter"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cdg1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tvwClaves"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvwParametros"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imlToolbarIcons"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lvwImportExport"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "picTitles"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancelar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdImportExport"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "picBasura"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Picture1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "picSplitter"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Buscador"
      TabPicture(1)   =   "frmTablas.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "ListViewFind"
      Tab(1).ControlCount=   2
      Begin VB.PictureBox picSplitter 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillColor       =   &H00808080&
         Height          =   4800
         Left            =   2130
         ScaleHeight     =   2090.126
         ScaleMode       =   0  'User
         ScaleWidth      =   780
         TabIndex        =   24
         Top             =   1290
         Visible         =   0   'False
         Width           =   72
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   690
         ScaleHeight     =   180
         ScaleWidth      =   4980
         TabIndex        =   23
         Top             =   5550
         Visible         =   0   'False
         Width           =   4980
      End
      Begin VB.PictureBox picBasura 
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   5625
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   600
         ScaleWidth      =   510
         TabIndex        =   19
         Top             =   3585
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CommandButton cmdImportExport 
         Caption         =   "&Exporta"
         Height          =   360
         Left            =   5760
         TabIndex        =   18
         Top             =   2775
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   5760
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.PictureBox picTitles 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   180
         ScaleHeight     =   300
         ScaleWidth      =   7845
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1245
         Width           =   7845
         Begin VB.Label lblTitle 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Parámetros"
            Height          =   270
            Index           =   1
            Left            =   2085
            TabIndex        =   16
            Tag             =   " Vista Lista:"
            Top             =   0
            Width           =   3210
         End
         Begin VB.Label lblTitle 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Claves"
            Height          =   270
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Tag             =   " Vista Árbol:"
            Top             =   0
            Width           =   2010
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buscar"
         Height          =   885
         Index           =   1
         Left            =   -74940
         TabIndex        =   5
         Top             =   360
         Width           =   10260
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   525
            Left            =   8760
            TabIndex        =   7
            Top             =   240
            Width           =   1305
         End
         Begin VB.TextBox txtTexto 
            Height          =   405
            Left            =   120
            TabIndex        =   6
            Top             =   300
            Width           =   8415
         End
      End
      Begin VB.OptionButton Option4 
         Height          =   195
         Left            =   -72210
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   810
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Height          =   195
         Left            =   -71190
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   810
         Width           =   255
      End
      Begin VB.ComboBox cmbUsuarios 
         Height          =   315
         Left            =   -74940
         TabIndex        =   2
         Top             =   390
         Width           =   3435
      End
      Begin MSComctlLib.ListView ListViewTablas 
         Height          =   6150
         Left            =   -74940
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
      Begin MSComctlLib.ListView ListViewFind 
         Height          =   6390
         Left            =   -74940
         TabIndex        =   12
         Top             =   1320
         Width           =   12465
         _ExtentX        =   21987
         _ExtentY        =   11271
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Member"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "VarType"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Valor"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lvwImportExport 
         Height          =   1500
         Left            =   6525
         TabIndex        =   20
         Top             =   4125
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2646
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Claves exportadas"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Comentario"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Uso"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList imlToolbarIcons 
         Left            =   7530
         Top             =   1590
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":047A
               Key             =   "LibroCerrado"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":08CC
               Key             =   "LibroAbierto"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":0D1E
               Key             =   "Export"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":1132
               Key             =   "Import"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":1244
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":1696
               Key             =   "Clave"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":1AE8
               Key             =   "Parametro"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":1F3A
               Key             =   "Root"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":238C
               Key             =   "Refresh"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":29A0
               Key             =   "Print"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":2AB4
               Key             =   "Links"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTablas.frx":2C14
               Key             =   "AnalisisRelaciones"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwParametros 
         Height          =   4395
         Left            =   2160
         TabIndex        =   21
         Top             =   1695
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   7752
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Dato"
            Object.Width           =   4128
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Comentario"
            Object.Width           =   4498
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Uso"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwClaves 
         Height          =   4395
         Left            =   180
         TabIndex        =   22
         Top             =   1680
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   7752
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlToolbarIcons"
         Appearance      =   1
         OLEDropMode     =   1
      End
      Begin MSComDlg.CommonDialog cdg1 
         Left            =   8160
         Top             =   3030
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgSplitter 
         Height          =   4785
         Left            =   4455
         MousePointer    =   9  'Size W E
         Top             =   1080
         Width           =   150
      End
      Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
         Left            =   8190
         Top             =   2340
         _cx             =   847
         _cy             =   847
         PassiveMode     =   0   'False
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
         TabIndex        =   13
         Top             =   420
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ABMAlta            As Integer = 1
Private Const ABMBaja            As Integer = 2
Private Const ABMModificacion    As Integer = 3

'constantes
Private Const sglSplitLimit   As Integer = 1200
Private KeyROOT               As String                     'clave para tarea root
Private Title                 As String
Private frm                   As frmDialog

Private Type udtRecord
   Clave1                     As String
   Clave2                     As String
   Clave3                     As String
   Clave4                     As String
   Clave5                     As String
   Valor                      As String
   COMENTARIO                 As String
   RESERVADA                  As String
End Type

Private FDRecord              As udtRecord
Private mvarMenuKey           As String                     'clave del menu

Private strSQL                As String                     'sql para la seleccion de los elementos de la tabla
Private WithEvents rstTablas  As ADODB.Recordset            'recordset de tablas o parametros internos
Attribute rstTablas.VB_VarHelpID = -1
Private rstPrint              As ADODB.Recordset            'recordset para la impresion
Attribute rstPrint.VB_VarHelpID = -1

Private objTabla              As Object

'Private ReportObject          As clsReport                  'instancia de la clase report
Private ReportFile            As CRAXDRT.Report             'es el archivo cryXXXXX

Private mvarControlData       As DataShare.udtControlData   'información de control
Private ErrorLog              As ErrType                     'información del error generado

'variables de trabajo
Private bMoving               As Boolean                    'movimiento del spliter
Private strFindText           As String                     'string de busqueda
Private nods                  As New Collection             'colleccion de nodos para la busqueda
Private itms                  As New Collection             'coleccion de items para la busqueda
Private LastItmFound          As ListItem                   'ultimo itema encontrado
Private bEndOfFind            As Boolean                    'indicador de fine de busqueda
Private iLevel                As Integer                    'indica el nivel del nodo seleccionado (0,1,...n) 0=raiz
Private IsVisible             As Boolean
Private strCurrentOperation   As String                     '"Importar/Exportar/Imprimir/Relacion"


'variables usadas para operaciones de D&D
Private SourceObject          As Object                     'referencia al objecto que inicia el D&D
Private SourceNode            As Node                       'referencia al nodo que esta siendo dragged
Private SourceItem            As ListItem                   'referencia al item que esta siendo dragged
Private ShiftState            As Integer                    'estado del Shift durante operacion de D&D
Private pBar                  As New clsProgressBar
Private bEndOfAnalisis        As Boolean                    'sirve para interrumpir el analisis de las relaciones

Private Sub cmdCancelar_Click()
Dim ix As Integer

   bEndOfAnalisis = True
   
   tvwClaves.Visible = True
   lvwParametros.Visible = True
   
   lvwImportExport.Visible = False
   cmdImportExport.Visible = False
   cmdCancelar.Visible = False
   picBasura.Visible = False
     
   SizeControls imgSplitter.Left
  
End Sub


Private Sub Form_Load()
Dim retval As Long

   On Error GoTo GestErr
   
   Me.Icon = LoadPicture(Icons & "Forms.ico")
   DoEvents
   
   ' En el evento Load definir procedimientos que sean independientes de la Empresa y/o MenuKey

   Set pBar.Canvas = Picture1

'   Set ReportObject = New clsReport
   
   'este DoEvents es necesario para procesar eventos del usuario
   'durante la fase de Load (ej, El usuario descargo el form)
   DoEvents
   
   InitForm
   
   Exit Sub
   
GestErr:
'   LoadError ErrorLog, "Form_Load"
'   ShowErrMsg ErrorLog
End Sub

Private Sub Form_Resize()
   
   If Me.MDIExtend1.WindowState = vbMaximized Or Me.MDIExtend1.WindowState = vbMinimized Then Exit Sub
   
'   Me.Move 0, 0, mvarMDIForm.ScaleWidth, mvarMDIForm.ScaleHeight
   
   SizeControls imgSplitter.Left
   
End Sub

Private Sub Form_Unload(Cancel As Integer)


      
      tvwClaves.Nodes.Clear
      lvwParametros.ListItems.Clear

      ReleaseObjects

   
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With imgSplitter
      picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
   End With
   picSplitter.Visible = True
   bMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sglPos As Single
  
   If bMoving Then
      sglPos = X + imgSplitter.Left
      If sglPos < sglSplitLimit Then
         picSplitter.Left = sglSplitLimit
      ElseIf sglPos > SSTab.Width - sglSplitLimit Then
         picSplitter.Left = SSTab.Width - sglSplitLimit
      Else
         picSplitter.Left = sglPos
      End If
   End If
  
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SizeControls picSplitter.Left
   picSplitter.Visible = False
   bMoving = False
End Sub

Sub SizeControls(X As Single)
   On Error Resume Next
   
   If lvwImportExport.Visible Then Exit Sub
   If Me.MDIExtend1.WindowState = vbMinimized Then Exit Sub
   
   'establecer el ancho
   If X < 1500 Then X = 1500
   If X > (SSTab.Width - 1500) Then X = SSTab.Width - 1500
   tvwClaves.Width = X
   imgSplitter.Left = X
   lvwParametros.Left = X + 400
   lvwParametros.Width = SSTab.Width - (tvwClaves.Width + 140)
   lblTitle(0).Width = tvwClaves.Width
   lblTitle(1).Left = lvwParametros.Left + 20
   lblTitle(1).Width = lvwParametros.Width - 40
   
   picTitles.Width = ScaleWidth
   picTitles.Top = 500 'IIf(tbr1.Visible, tbr1.Top + tbr1.Height, 0)
   tvwClaves.Top = picTitles.Top + picTitles.Height
   
   lvwParametros.Top = tvwClaves.Top
   
   'establecer el alto
   tvwClaves.Height = 7500 'Me.ScaleHeight - (picTitles.Top + picTitles.Height)
   
   lvwParametros.Height = tvwClaves.Height
   imgSplitter.Top = tvwClaves.Top
   imgSplitter.Height = tvwClaves.Height
   
End Sub

Private Sub lvwImportExport_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyDelete Then TirarBasura

End Sub

Private Sub lvwImportExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ' verifico si hemos iniciado una operacion de drag-drop
   If Button <> vbLeftButton Then Exit Sub
   ' seteo el nodo que esta siendo dragged, me voy si no hay ninguno
   Set SourceItem = lvwImportExport.HitTest(X, Y)
   If SourceItem Is Nothing Then Exit Sub
   
   ShiftState = Shift
   
   ' salvo valores para mas adelante
   Set SourceObject = lvwImportExport
   ' inicio operacion de drag
   lvwImportExport.OLEDrag

End Sub

Private Sub lvwImportExport_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim rstExport     As New ADODB.Recordset
Dim itmx          As ListItem
Dim strKey        As String
Dim strFullPath   As String
Dim aKeys()

   On Error GoTo GestErr
   
   ' chequeo que el destino es el mismo lvwImportExport
   If SourceObject Is lvwImportExport Then
      MsgBox "No es posible efectuar esta operación", vbExclamation
      Set lvwImportExport.DropHighlight = Nothing
      Exit Sub
   End If
   
   Select Case ActiveControl.Name
      Case "tvwClaves"
         strKey = GetNodeKey(tvwClaves.SelectedItem)
         objTabla.GetKeys strKey, aKeys
         Set tvwClaves.DropHighlight = Nothing
      Case "lvwParametros"
         strKey = GetParameterKey(lvwParametros.SelectedItem)
         objTabla.GetKeys strKey, aKeys
   End Select
         
   Select Case UBound(aKeys)
      Case 1
         If aKeys(1) = KeyROOT Then
         Else
            rstTablas.Filter = FDRecord.Clave1 & " = '" & aKeys(1) & "'"
         End If
      Case 2
         rstTablas.Filter = FDRecord.Clave1 & " = '" & aKeys(1) & "' AND  " & FDRecord.Clave2 & " = '" & aKeys(2) & "'"
      Case 3
         rstTablas.Filter = FDRecord.Clave1 & " = '" & aKeys(1) & "' AND  " & FDRecord.Clave2 & " = '" & aKeys(2) & "' AND " & FDRecord.Clave3 & " = '" & aKeys(3) & "'"
      Case 4
         rstTablas.Filter = FDRecord.Clave1 & " = '" & aKeys(1) & "' AND  TAB_CLAVE2 = '" & aKeys(2) & "' AND " & FDRecord.Clave3 & " = '" & aKeys(3) & "' AND " & FDRecord.Clave4 & " = '" & aKeys(4) & "'"
      Case 5
         rstTablas.Filter = FDRecord.Clave1 & " = '" & aKeys(1) & "' AND  TAB_CLAVE2 = '" & aKeys(2) & "' AND " & FDRecord.Clave3 & " = '" & aKeys(3) & "' AND " & FDRecord.Clave4 & " = '" & aKeys(4) & "' AND " & FDRecord.Clave5 & " = '" & aKeys(5) & "'"
   End Select
   
   If rstTablas.RecordCount > 0 Then rstTablas.MoveFirst

   Do While Not rstTablas.EOF
   
      Select Case NumberOfKeys(rstTablas(FDRecord.Clave1), rstTablas(FDRecord.Clave2), rstTablas(FDRecord.Clave3), rstTablas(FDRecord.Clave4), rstTablas(FDRecord.Clave5))
         Case 1
             strKey = rstTablas(FDRecord.Clave1)
         Case 2
             strKey = rstTablas(FDRecord.Clave1) & "\" & rstTablas(FDRecord.Clave2)
         Case 3
             strKey = rstTablas(FDRecord.Clave1) & "\" & rstTablas(FDRecord.Clave2) & "\" & rstTablas(FDRecord.Clave3)
         Case 4
             strKey = rstTablas(FDRecord.Clave1) & "\" & rstTablas(FDRecord.Clave2) & "\" & rstTablas(FDRecord.Clave3) & "\" & rstTablas(FDRecord.Clave4)
         Case 5
             strKey = rstTablas(FDRecord.Clave1) & "\" & rstTablas(FDRecord.Clave2) & "\" & rstTablas(FDRecord.Clave3) & "\" & rstTablas(FDRecord.Clave4) & "\" & rstTablas(FDRecord.Clave5)
      End Select
      
      strFullPath = strKey
      Set itmx = lvwImportExport.FindItem(strFullPath, lvwText)
      If TypeName(itmx) = "Nothing" Then
         
         If strCurrentOperation = "Relacion" Then
            If strKey = Trim(Mid(lvwImportExport.ColumnHeaders(1).Text, InStr(lvwImportExport.ColumnHeaders(1).Text, ":") + 1)) Then
               MsgBox "No es posible relacionar un parámetro con si mismo", vbOKOnly, App.ProductName
               Exit Sub
            End If
         End If
         Set itmx = lvwImportExport.ListItems.Add()
         itmx.Text = strFullPath
         If strCurrentOperation = "Relacion" Then
            itmx.SubItems(1) = "Directa"
         Else
            itmx.SubItems(1) = IIf(IsNull(rstTablas(FDRecord.Valor)), NullString, rstTablas(FDRecord.Valor))
            itmx.SubItems(2) = IIf(IsNull(rstTablas(FDRecord.COMENTARIO)), NullString, rstTablas(FDRecord.COMENTARIO))
               itmx.SubItems(3) = IIf(IsNull(rstTablas(FDRecord.RESERVADA)), No, rstTablas(FDRecord.RESERVADA))
               If itmx.SubItems(3) = No Then
                  itmx.SubItems(3) = NullString
               Else
                  itmx.SubItems(3) = "Reservado"
               End If
         End If
         itmx.key = "K" & CStr(strKey)
      Else
         MsgBox "La clave " & strKey & " ya existe en la lista", vbOKOnly, App.ProductName
      End If
      
      rstTablas.MoveNext
   Loop
         
   rstTablas.Filter = adFilterNone
   
   Exit Sub
   
GestErr:
'   LoadError ErrorLog, "lvwImportExport_OLEDragDrop"
'   ShowErrMsg ErrorLog

End Sub

Private Sub lvwImportExport_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)

   ' paso la propiedad Key de el nodo que esta siendo dragged
   ' (este valor no es usado, actualmente paso cualquier cosa)
   Data.SetData SourceItem.key
   If ShiftState And vbCtrlMask Then
       AllowedEffects = vbDropEffectCopy
   Else
       AllowedEffects = vbDropEffectMove
   End If
   
End Sub

Private Sub lvwParametros_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim strOldKey  As String
Dim strNewKey  As String
Dim strKey     As String
Dim itmx As ListItem
Dim aKeys()

   On Error GoTo GestErr

   If NewString = NullString Then
      MsgBox "El nombre del parámetro debe contener almenos 1 carácter", vbOKOnly, App.ProductName
      Cancel = True
      Exit Sub
   End If
   
   If (MsgBox("Confirma el nuevo nombre del parámetro ?", vbQuestion + vbYesNo, App.ProductName) = vbNo) Then
      Cancel = True
      Exit Sub
   End If
   
   NewString = EliminaAcentos(NewString)
   
   Set itmx = lvwParametros.FindItem(NewString)
   If TypeName(itmx) = "Nothing" Then
   Else
      Cancel = True
      MsgBox "El parámetro ya existe.", vbExclamation, Title
      Exit Sub
   End If
   
   ' obtengo la clave del parametro
   strKey = GetParameterKey(lvwParametros.SelectedItem)
   
   objTabla.GetKeys strKey, aKeys
   
   'clave del item seleccionado antes del cambio
   strOldKey = strKey
   
   'obtiene un nueva clave
   strNewKey = RenameKey(strOldKey, iLevel + 1, NewString)
   
   'actualizo la base
   If TypeName(objTabla) = "clsParametrosInternos" Then
      objTabla.AlterKey strOldKey, strNewKey
   Else
      objTabla.AlterKey "", strOldKey, strNewKey
   End If
   
   objTabla.GetKeys strOldKey, aKeys
   
   'actualizo el recordset
   Select Case UBound(aKeys)
      Case 1
      Case 2
        rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "')"
        rstTablas(FDRecord.Clave2) = NewString
      
      Case 3
         rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "')"
         rstTablas(FDRecord.Clave3) = NewString
      
      Case 4
         rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND  (" & FDRecord.Clave4 & " = '" & aKeys(4) & "')"
         rstTablas(FDRecord.Clave4) = NewString
      
      Case 5
         rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND  (" & FDRecord.Clave4 & " = '" & aKeys(4) & "') AND  (" & FDRecord.Clave5 & " = '" & aKeys(5) & "')"
         rstTablas(FDRecord.Clave5) = NewString
   
   End Select
   rstTablas.Update
   
   'actualizo el listview
   Set itmx = lvwParametros.SelectedItem
   itmx.key = "K" & CStr(NewString)
   
   rstTablas.Filter = adFilterNone
  
   Exit Sub
   
GestErr:
'   LoadError ErrorLog, "lvwParametros_AfterLabelEdit"
'   ShowErrMsg ErrorLog
   
   Cancel = True
   rstTablas.Filter = adFilterNone
End Sub

Private Sub lvwParametros_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
   
   ' paso la propiedad Key de el nodo que esta siendo dragged
   ' (este valor no es usado, actualmente paso cualquier cosa)
   Data.SetData SourceItem.key
   If ShiftState And vbCtrlMask Then
       AllowedEffects = vbDropEffectCopy
   Else
       AllowedEffects = vbDropEffectMove
   End If

End Sub


Private Sub LoadTablas()
Dim sql As String

   Set rstTablas = Fetch("ALG", strSQL)
   
   ConstruyeTreeView
      
End Sub

Private Sub picBasura_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If strCurrentOperation = "Relacion" And lvwImportExport.SelectedItem.SubItems(1) = "Indirecta" Then
      MsgBox "No es posible eliminar una relación indirecta", vbOKOnly, App.ProductName
      Exit Sub
   End If

   TirarBasura

End Sub

Private Sub tvwClaves_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim aKeys()
'Dim strSQL     As String
Dim strOldKey  As String
Dim strNewKey  As String
Dim strKey     As String
'Dim ix         As Integer
'Dim iPos       As Integer

   On Error GoTo GestErr
   
   If NewString = NullString Then
      MsgBox "El nombre de la clave debe contener almenos 1 carácter", vbOKOnly, App.ProductName
      Cancel = True
      Exit Sub
   End If
   
   
   If (MsgBox("Confirma el nuevo nombre de la clave ?", vbQuestion + vbYesNo, App.ProductName) = vbNo) Then
      Cancel = True
      Exit Sub
   End If
   
   NewString = EliminaAcentos(NewString)
   
   ' obtengo la clave del record decifrando la clave del nodo y la metto en aKeys
   strKey = GetNodeKey(tvwClaves.SelectedItem)
   objTabla.GetKeys strKey, aKeys
   
   'clave del item seleccionado antes del cambio
   strOldKey = strKey
   
   'obtiene un nueva clave
   strNewKey = RenameKey(strOldKey, iLevel, NewString)

   'intento el cambio de la clave
   On Error Resume Next
   tvwClaves.SelectedItem.key = CStr(strNewKey)
   
   If Err.Number <> 0 Then
      MsgBox "La clave ya existe", vbExclamation, Title
      Cancel = True
      Exit Sub
   End If
   
   On Error GoTo GestErr
   
   'actualizo la base
   If TypeName(objTabla) = "clsParametrosInternos" Then
      objTabla.AlterKey strOldKey, strNewKey
   Else
      objTabla.AlterKey "", strOldKey, strNewKey
   End If
   
   
   'filtro los records interesados
   Select Case UBound(aKeys)
      Case 1
         rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "')"
      Case 2
         rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "')"
      Case 3
         rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "')"
      Case 4
         rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND  (" & FDRecord.Clave4 & " = '" & aKeys(4) & "')"
      Case 5
         rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND  (" & FDRecord.Clave4 & " = '" & aKeys(4) & "') AND  (" & FDRecord.Clave5 & " = '" & aKeys(5) & "')"
   End Select
   
   'actualizo el recordset
   Do While Not rstTablas.EOF
   
      Select Case UBound(aKeys)
         Case 1
            rstTablas(FDRecord.Clave1) = NewString
         Case 2
            rstTablas(FDRecord.Clave2) = NewString
         Case 3
            rstTablas(FDRecord.Clave3) = NewString
         Case 4
            rstTablas(FDRecord.Clave4) = NewString
         Case 5
            rstTablas(FDRecord.Clave5) = NewString
      End Select
      rstTablas.Update
      rstTablas.MoveNext
   Loop
   
   rstTablas.Filter = adFilterNone
   
   ConstruyeTreeView
   
   Exit Sub
   
GestErr:
   Cancel = True
'   LoadError ErrorLog, "tvwClaves_AfterLabelEdit"
'   ShowErrMsg ErrorLog
End Sub


Private Sub tvwClaves_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyDelete Then
'      mnuEliminar_Click
   End If

End Sub

Private Sub tvwClaves_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If strCurrentOperation = "Relacion" Then Exit Sub

   ' verifico si hemos iniciado una operacion de drag-drop
   If Button <> vbLeftButton Then Exit Sub
   ' seteo el nodo que esta siendo dragged, me voy si no hay ninguno
   Set SourceNode = tvwClaves.HitTest(X, Y)
   If SourceNode Is Nothing Then Exit Sub
   
   ShiftState = Shift
   
   ' salvo valores para mas adelante
   Set SourceObject = tvwClaves
   ' inicio operacion de drag
   tvwClaves.OLEDrag

End Sub



Private Sub ConstruyeTreeView()
      Dim nodX          As Node
      Dim strKey        As String
      Dim strKeyParent  As String
'      Dim itmX          As ListItem

10       On Error GoTo GestErr
   
20       With Me
      
30          .tvwClaves.Nodes.Clear
      

80         .lvwParametros.ListItems.Clear

90          Set nodX = .tvwClaves.Nodes.Add(, , KeyROOT, KeyROOT)
100         nodX.Image = "Root"
      
110         If rstTablas.RecordCount = 0 Then Exit Sub
      
120         rstTablas.MoveFirst
130         Do
140           Select Case NumberOfKeys(rstTablas(FDRecord.Clave1), rstTablas(FDRecord.Clave2), rstTablas(FDRecord.Clave3), rstTablas(FDRecord.Clave4), rstTablas(FDRecord.Clave5))
                Case 1
150               strKey = KeyROOT & "\" & rstTablas(FDRecord.Clave1)
160               strKeyParent = KeyROOT
            
170               If Not FindKey(strKeyParent) Then
180                 CreateParent strKeyParent
190               End If
200               If Not FindKey(strKey) Then
210                  Set nodX = .tvwClaves.Nodes.Add(strKeyParent, tvwChild, strKey, rstTablas(FDRecord.Clave1))
220               End If
          
230             Case 2
              
240               strKey = KeyROOT & "\" & rstTablas(FDRecord.Clave1)
250               strKeyParent = KeyROOT
            
260               If Not FindKey(strKeyParent) Then
270                 CreateParent strKeyParent
280               End If
290               If Not FindKey(strKey) Then
300                  Set nodX = .tvwClaves.Nodes.Add(strKeyParent, tvwChild, strKey, rstTablas(FDRecord.Clave1))
310               End If
            
320             Case 3

330               strKey = KeyROOT & "\" & rstTablas(FDRecord.Clave1) & "\" & rstTablas(FDRecord.Clave2)
340               strKeyParent = KeyROOT & "\" & rstTablas(FDRecord.Clave1)
            
350               If Not FindKey(strKeyParent) Then
360                 CreateParent strKeyParent
370               End If
380               If Not FindKey(strKey) Then
390                  Set nodX = .tvwClaves.Nodes.Add(strKeyParent, tvwChild, strKey, rstTablas(FDRecord.Clave2))
400               End If

410             Case 4

420               strKey = KeyROOT & "\" & rstTablas(FDRecord.Clave1) & "\" & rstTablas(FDRecord.Clave2) & "\" & rstTablas(FDRecord.Clave3)
430               strKeyParent = KeyROOT & "\" & rstTablas(FDRecord.Clave1) & "\" & rstTablas(FDRecord.Clave2)
            
440               If Not FindKey(strKeyParent) Then
450                 CreateParent strKeyParent
460               End If
470               If Not FindKey(strKey) Then
480                  Set nodX = .tvwClaves.Nodes.Add(strKeyParent, tvwChild, strKey, rstTablas(FDRecord.Clave3))
490               End If

500             Case 5

510               strKey = KeyROOT & "\" & rstTablas(FDRecord.Clave1) & "\" & rstTablas(FDRecord.Clave2) & "\" & rstTablas(FDRecord.Clave3) & "\" & rstTablas(FDRecord.Clave4)
520               strKeyParent = KeyROOT & "\" & rstTablas(FDRecord.Clave1) & "\" & rstTablas(FDRecord.Clave2) & "\" & rstTablas(FDRecord.Clave3)
            
530               If Not FindKey(strKeyParent) Then
540                  CreateParent strKeyParent
550               End If
560               If Not FindKey(strKey) Then
570                  Set nodX = .tvwClaves.Nodes.Add(strKeyParent, tvwChild, strKey, rstTablas(FDRecord.Clave4))
580               End If

590           End Select
                  
600           nodX.Image = "LibroCerrado"
610           nodX.ExpandedImage = "LibroAbierto"

620           rstTablas.MoveNext
630         Loop Until rstTablas.EOF
      
640         .tvwClaves.Nodes(1).Expanded = True
      
650     End With
  
660     Exit Sub
   
GestErr:
'670      LoadError ErrorLog, "ConstruyeTreeView" & Erl & vbCrLf & " Key " & strKey & _
'                                                         vbCrLf & " ParentKey: " & strKeyParent
'680      ShowErrMsg ErrorLog
End Sub

Private Function FindKey(strKey As String) As Boolean
Dim nodX As Node
  
   ' busco si existe la clave en el treeview
   
   On Error Resume Next
   
   Set nodX = tvwClaves.Nodes(strKey)
       
   If nodX Is Nothing Then
      FindKey = False
   Else
      FindKey = True
   End If

End Function

Private Sub CreateParent(strKeyParent As String)
Dim nodX       As Node             'nodo a crear
Dim strKey     As String           'valor de la clave para cada nodo que se visita
Dim nOffs      As Integer          'offset para la funcion Instr
Dim iPos       As Integer          '
Dim strParent  As String           'clave del padre para cada nodo que se visita
  
   '---------------------------------------------------------------------------------
   '  Crea la clave strKeyParent creando, si es necesario, los nodos no existentes
   '---------------------------------------------------------------------------------
   
   strKey = NullString
   strParent = NullString
   nOffs = 1
   iPos = 1
   Do While iPos <> 0
   
      iPos = InStr(nOffs, strKeyParent, "\")
      If iPos > 0 Then
         strKey = Left(strKeyParent, iPos - 1)
         If Not FindKey(strKey) Then
            If Len(strParent) = 0 Then
              Set nodX = tvwClaves.Nodes.Add(, , strKey, GetText(strKey))
            Else
              Set nodX = tvwClaves.Nodes.Add(strParent, tvwChild, strKey, GetText(strKey))
            End If
         End If
         strParent = strKey
         nOffs = iPos + 1
      Else
         strKey = strKeyParent
         If Not FindKey(strKey) Then
            If Len(strParent) = 0 Then
               Set nodX = tvwClaves.Nodes.Add(, , strKey, GetText(strKey))
            Else
               Set nodX = tvwClaves.Nodes.Add(strParent, tvwChild, strKey, GetText(strKey))
            End If
         End If
      End If
      
      If TypeName(nodX) <> "Nothing" Then
         ' el nodo fue creado
         nodX.Image = "LibroCerrado"
         nodX.ExpandedImage = "LibroAbierto"
      End If
   Loop
  
End Sub

Private Function GetText(ByVal strKey As String) As String
Dim iPos As Integer

   iPos = 1
   Do While iPos <> 0
   
      iPos = InStr(strKey, "\")
      If iPos > 0 Then
         strKey = Mid(strKey, iPos + 1)
      End If
     
   Loop
   
   GetText = strKey
  
End Function

Private Function NumberOfKeys(ByVal vValue1 As Variant, ByVal vValue2 As Variant, ByVal vValue3 As Variant, ByVal vValue4 As Variant, ByVal vValue5 As Variant) As Byte

   If Not IsNull(vValue5) Then NumberOfKeys = 5: Exit Function
   If Not IsNull(vValue4) Then NumberOfKeys = 4: Exit Function
   If Not IsNull(vValue3) Then NumberOfKeys = 3: Exit Function
   If Not IsNull(vValue2) Then NumberOfKeys = 2: Exit Function
   If Not IsNull(vValue1) Then NumberOfKeys = 1: Exit Function

End Function

Private Sub tvwClaves_NodeClick(ByVal Node As MSComctlLib.Node)
Dim itmx    As ListItem
Dim strKey  As String
Dim aKeys()
'Dim vValue  As Variant
   
   On Error GoTo GestErr
   
   'actualizo el nivel del nodo corriente
   objTabla.GetKeys Node.key, aKeys
   
   iLevel = UBound(aKeys) - 1
   
   If (TypeName(Node) = "Nothing") Or (Node = KeyROOT) Then
      lvwParametros.ListItems.Clear
      Exit Sub
   End If
   
   ' obtengo la clave del record descifrando la clave del nodo y la meto en aKeys
   
   strKey = GetNodeKey(Node)
   objTabla.GetKeys strKey, aKeys
   
   Dim s As String
'   S = "TAB_EMPRESA = '" & mvarControlData.Empresa & "' AND  "
   
   Select Case UBound(aKeys)
      Case 1
        rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "')"
      Case 2
        rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND (" & FDRecord.Clave2 & " = '" & aKeys(2) & "')"
      Case 3
        rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND (" & FDRecord.Clave3 & " = '" & aKeys(3) & "')"
      Case 4
        rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND (" & FDRecord.Clave4 & " = '" & aKeys(4) & "')"
   End Select
   
   itmsClear
   
   lvwParametros.ListItems.Clear
   Do While Not rstTablas.EOF
   
      If IsNull(rstTablas(UBound(aKeys) + 2)) Or UBound(aKeys) = 4 Then
         'actualizo el ListView
         If Not IsNull(rstTablas(Left(FDRecord.Clave1, 9) & UBound(aKeys) + 1)) Then
            If rstTablas(Left(FDRecord.Clave1, 9) & UBound(aKeys) + 1) <> "(Predeterminado)" Then
               Set itmx = lvwParametros.ListItems.Add()
               itmx.Text = rstTablas(Left(FDRecord.Clave1, 9) & UBound(aKeys) + 1)
               itmx.SubItems(1) = IIf(IsNull(rstTablas(FDRecord.Valor)), NullString, rstTablas(FDRecord.Valor))
               itmx.SubItems(2) = IIf(IsNull(rstTablas(FDRecord.COMENTARIO)), NullString, rstTablas(FDRecord.COMENTARIO))
                 itmx.SubItems(3) = IIf(IsNull(rstTablas(FDRecord.RESERVADA)), No, rstTablas(FDRecord.RESERVADA))
                 If itmx.SubItems(3) = Si Then
                    itmx.SubItems(3) = "Reservado"
                 Else
                    itmx.SubItems(3) = NullString
                End If

               itmx.key = "K" & CStr(itmx.Text)
            End If
         End If
      End If
     rstTablas.MoveNext
   Loop
   
   rstTablas.Filter = adFilterNone
    
   For Each itmx In lvwParametros.ListItems
     itms.Add itmx
   Next itmx
    
   stb1.Panels(1).Text = Node.FullPath
  
   Exit Sub
   
GestErr:
   
End Sub

Private Sub FindX()
Dim nod  As Node
Dim itmx As ListItem

   If itms.Count > 0 Then
      Do While itms.Count > 0
         Set itmx = itms(1)
         If (InStr(UCase(itmx.Text), strFindText) > 0) Or (InStr(UCase(itmx.SubItems(1)), strFindText) > 0) Then
            Set lvwParametros.SelectedItem = lvwParametros.FindItem(itmx.Text)
            Set LastItmFound = lvwParametros.SelectedItem
            itms.Remove (1)
            Exit Sub
         Else
           If itms.Count > 0 Then
              itms.Remove (1)
           End If
         End If
      Loop
      If nods.Count > 0 Then
         nods.Remove (1)
      End If
   Else
     If nods.Count > 0 Then
       nods.Remove (1)
     End If
   End If
   
   For Each nod In nods
   
      If (InStr(UCase(nod.Text), strFindText) > 0) Then
        
        Set tvwClaves.SelectedItem = tvwClaves.Nodes(nod.key)
        Set LastItmFound = Nothing
        tvwClaves_NodeClick nod
        nods.Remove (1)
        Exit Sub
        
      Else
      
        'inicio la busqueda en itms
        tvwClaves_NodeClick nod
        Do While itms.Count > 0
          Set itmx = itms(1)
          If (InStr(UCase(itmx.Text), strFindText) > 0) Or (InStr(UCase(itmx.SubItems(1)), strFindText) > 0) Then
            Set tvwClaves.SelectedItem = tvwClaves.Nodes(nod.key)
            Set lvwParametros.SelectedItem = lvwParametros.FindItem(itmx.Text)
            Set LastItmFound = lvwParametros.SelectedItem
            itms.Remove (1)
            Exit Sub
          Else
            If itms.Count > 0 Then
               itms.Remove (1)
            End If
          End If
        Loop
        nods.Remove (1)
      End If
       
   Next nod
   
   If nods.Count = 0 Then
      If TypeName(tvwClaves.SelectedItem) <> "Nothing" Then
         tvwClaves_NodeClick tvwClaves.SelectedItem
         If TypeName(LastItmFound) <> "Nothing" Then
            Set lvwParametros.SelectedItem = lvwParametros.FindItem(LastItmFound.Text)
         End If
      End If
      MsgBox "No fueron encontradas nuevas occurencias de la clave seleccionada", vbOKOnly, Title
      bEndOfFind = True
      Exit Sub
   End If

  
End Sub
Private Sub StartFind()
Dim nod As Node

   'inicia la busqueda
   
   bEndOfFind = False
   
   For Each nod In tvwClaves.Nodes
      nods.Add nod
   Next nod
   itmsClear
   FindX

End Sub

Private Sub itmsClear()

   'hace el clear de la colleccion itms
   Do While itms.Count > 0
      itms.Remove (1)
   Loop

End Sub

Private Sub ImprimirClaves()
Dim oldHeight As Integer
   
   strCurrentOperation = "Imprimir"
   
   If lvwImportExport.Visible Then Exit Sub
   lvwImportExport.LabelEdit = lvwManual
   
   oldHeight = tvwClaves.Height
   tvwClaves.Height = tvwClaves.Height / 2
   lvwParametros.Height = tvwClaves.Height
   With lvwImportExport
      
      .Height = oldHeight - tvwClaves.Height - 700
      .Top = tvwClaves.Top + tvwClaves.Height
      .Width = Me.ScaleWidth
      
      .ColumnHeaders(1).Width = lvwImportExport.Width
      .ColumnHeaders(2).Width = 0
      .ColumnHeaders(3).Width = 0
      
      .Visible = True
      .Left = 0
      .ZOrder 0
      
      .ColumnHeaders(1).Text = "Lista de claves a imprimir"
      
      .ListItems.Clear
   End With
   
   With picBasura
      .Left = 500
      .Top = lvwImportExport.Top + lvwImportExport.Height + 100
      .Visible = True
      .ZOrder 0
      .Picture = LoadPicture(BitMaps & "Cesto.bmp")
   End With
   
   With cmdImportExport
      .Caption = "&Imprimir"
'      .Top = picBasura.Top
'      .Left = lvwImportExport.Width - 1400
      .Visible = True
   End With
   
   With cmdCancelar
 '     .Top = cmdImportExport.Top
 '     .Left = cmdImportExport.Left - 1100
      .Visible = True
   End With

End Sub

Private Function RenameKey(ByVal strOldKey As String, ByVal iLevel As Integer, ByVal strNewString As String) As String
Dim aKeys()
Dim ix As Integer

   ' modifica la clave de un nodo del treeview cambiando el valor del nivel iLevel
   
   On Error GoTo GestErr
   
   RenameKey = NullString
   
   objTabla.GetKeys strOldKey, aKeys
   
   For ix = 1 To UBound(aKeys)
      If ix = iLevel Then
         RenameKey = RenameKey & "\" & strNewString
      Else
         RenameKey = RenameKey & "\" & aKeys(ix)
      End If
   Next ix
   
   If Left(RenameKey, 1) = "\" Then RenameKey = Mid(RenameKey, 2)
   
   Exit Function
   
GestErr:
'   LoadError ErrorLog, "RenameKey"
'   ShowErrMsg ErrorLog
   
End Function

Private Function GetParameterKey(ByVal Item As MSComctlLib.ListItem) As String

   GetParameterKey = GetNodeKey(tvwClaves.SelectedItem)
   GetParameterKey = GetParameterKey & "\" & Mid(Item.key, 2)

End Function

Private Function GetNodeKey(ByVal Node As MSComctlLib.Node)
                 
   GetNodeKey = Mid(Node.key, InStr(Node.key, "\") + 1)

End Function

Private Sub SetPerms()
'
'   If Not TaskIsEnabled(Me.MenuKey & ABMAlta, CUsuario) Then
'      tbr1.Buttons("NuevaClave").Enabled = False
'      tbr1.Buttons("NuevoParametro").Enabled = False
'      mnuNuevo.Enabled = False
'      mnuNuevaClave.Enabled = False
'      mnuNuevoParametro.Enabled = False
'      mnuClave.Enabled = False
'      mnuValor.Enabled = False
'   End If
'
'   If Not TaskIsEnabled(Me.MenuKey & ABMBaja, CUsuario) Then
'      mnuEliminar.Enabled = False
'   End If
'
'   If Not TaskIsEnabled(Me.MenuKey & ABMModificacion, CUsuario) Then
'      mnuModificar.Enabled = False
'      mnuRenombrar.Enabled = False
'   End If
'
End Sub
Public Sub PostInitForm()
   
   stb1.Panels(STB_PANEL1).Text = NullString
   Screen.MousePointer = vbDefault

End Sub


Public Sub InitForm()
      Dim nod As Node

10       On Error GoTo GestErr


         ' redimensiono el form en modo de ocupar el mayor espacio posible
'20       Me.Top = 0: Me.Left = 0
'30       Me.Height = mvarMDIForm.ScaleHeight
'40       Me.Width = mvarMDIForm.ScaleWidth
'
'50       Me.Show

      
70          KeyROOT = "TABLAS_DEL_SISTEMA"
80          Title = "Registro del Sistema"
      
'90          SetCaption Me, mvarControlData.Empresa, Title
      
100         CreateObjects
   
'110         If Not CUsuario.SysAdmin Then
'120            strSQL = "SELECT * FROM TABLAS WHERE TAB_TABLA_RESERVADA = 'No' ORDER BY TAB_CLAVE1, TAB_CLAVE2, TAB_CLAVE3, TAB_CLAVE4, TAB_CLAVE5"
'130         Else
'140            strSQL = "SELECT * FROM TABLAS ORDER BY TAB_CLAVE1, TAB_CLAVE2, TAB_CLAVE3, TAB_CLAVE4, TAB_CLAVE5"
140            strSQL = "SELECT * FROM TABLAS "
               strSQL = strSQL & "WHERE TAB_CLAVE1 <> 'Consultas (F3)' "
               strSQL = strSQL & "  AND TAB_CLAVE1 <> 'Dialogs' "
               strSQL = strSQL & "  AND TAB_CLAVE1 <> 'Exportar_A_Office' "
               strSQL = strSQL & "  AND TAB_CLAVE1 <> 'Filtros' "
               strSQL = strSQL & "  AND TAB_CLAVE1 <> 'Forms' "
               strSQL = strSQL & "  ORDER BY TAB_CLAVE1, TAB_CLAVE2, TAB_CLAVE3, TAB_CLAVE4, TAB_CLAVE5"
'150         End If
   
160         With FDRecord
170            .Clave1 = "TAB_CLAVE1"
180            .Clave2 = "TAB_CLAVE2"
190            .Clave3 = "TAB_CLAVE3"
200            .Clave4 = "TAB_CLAVE4"
210            .Clave5 = "TAB_CLAVE5"
220            .Valor = "TAB_VALOR"
230            .COMENTARIO = "TAB_COMENTARIO"
240            .RESERVADA = "TAB_TABLA_RESERVADA"
250         End With
   

   
         'inicializo los objetos
470      SetObjects
   

   
   
530      SetPerms
   
540      LoadTablas

550      For Each nod In tvwClaves.Nodes
560         If nod.Children > 0 Then nod.Expanded = False
570      Next nod
   
        'expande el primer nivel
580     tvwClaves.Nodes(1).Expanded = True

590   Exit Sub

GestErr:
'600      LoadError ErrorLog, "InitForm" & Erl
'610      ShowErrMsg ErrorLog


End Sub

Private Sub tvwClaves_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If strCurrentOperation <> "Importar" Or strCurrentOperation = "Exportar" Then Exit Sub
    
   ' chequeo que el destino es el mismo treeview
   If SourceObject Is tvwClaves Then
       MsgBox "No es posible efectuar esta operación", vbExclamation
   End If
   Set tvwClaves.DropHighlight = Nothing

End Sub

Private Sub tvwClaves_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    
    ' paso la propiedad Key de el nodo que esta siendo dragged
    ' (este valor no es usado, actualmente paso cualquier cosa)
    Data.SetData SourceNode.key
    If ShiftState And vbCtrlMask Then
        AllowedEffects = vbDropEffectCopy
    Else
        AllowedEffects = vbDropEffectMove
    End If

End Sub

Private Sub TirarBasura()
Dim itmx    As ListItem
Dim liItems As New Collection

   For Each itmx In lvwImportExport.ListItems
      If itmx.Selected Then
         liItems.Add itmx
      End If
   Next itmx
   
   For Each itmx In liItems
      If itmx.Selected Then
         lvwImportExport.ListItems.Remove itmx.key
      End If
   Next itmx
   
   picBasura.Picture = LoadPicture(BitMaps & "CestoFull.bmp")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   mvarMDIForm.MDITaskBar1.RemoveFormFromTaskBar Me.hWnd
End Sub

Public Property Let MenuKey(ByVal vData As String)
   mvarMenuKey = vData
End Property
Public Property Get MenuKey() As String
   MenuKey = mvarMenuKey
End Property

'Private Function Clave(ByVal Clave1 As Variant, ByVal Clave2 As Variant, _
'                       ByVal Clave3 As Variant, ByVal Clave4 As Variant, _
'                       ByVal Clave5 As Variant) As String
'
'   If IsNull(Clave1) Then Exit Function
'
'   Clave = Clave1
'
'   If IsNull(Clave2) Then Exit Function
'   Clave = Clave & "\" & Clave2
'   If IsNull(Clave3) Then Exit Function
'   Clave = Clave & "\" & Clave3
'   If IsNull(Clave4) Then Exit Function
'   Clave = Clave & "\" & Clave4
'   If IsNull(Clave5) Then Exit Function
'   Clave = Clave & "\" & Clave5
'
'End Function

Private Sub CreateObjects()

   Set objTabla = New BOGeneral.clsTablas
   
End Sub

Private Sub ReleaseObjects()

   Set objTabla = Nothing
   
'   Unload ReportObject.ReportForm
   
'   Set ReportObject = Nothing
   Set frmTablas = Nothing

End Sub

Private Sub OpenRecordset()
Dim strFilter As String
Dim strKey As String
Dim itmx As ListItem

   strFilter = NullString
   For Each itmx In lvwImportExport.ListItems
      
      ' convierte el path en clave
      strKey = itmx.Text
         
      objTabla.GetKeys strKey, aKeys
      Select Case UBound(aKeys)
        Case 1
          strFilter = strFilter & " OR (" & FDRecord.Clave1 & " = '" & aKeys(1) & "')"
        Case 2
          strFilter = strFilter & " OR ((" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "'))"
        Case 3
          strFilter = strFilter & " OR ((" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "'))"
        Case 4
          strFilter = strFilter & " OR ((" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND  (" & FDRecord.Clave4 & " = '" & aKeys(4) & "'))"
        Case 5
          strFilter = strFilter & " OR ((" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND  (" & FDRecord.Clave4 & " = '" & aKeys(4) & "') AND  (" & FDRecord.Clave5 & " = '" & aKeys(5) & "'))"
      End Select
     
   Next itmx

   strFilter = "(" & Mid(strFilter, 5) & ")"
   
   Set rstPrint = Fetch(mvarControlData.Empresa, "SELECT * FROM TABLAS WHERE " & strFilter)

End Sub

Public Property Let ControlData(ByRef vData As DataShare.udtControlData)
   mvarControlData = vData
   
   If mvarMenuKey <> "REGISTRO_SISTEMA" Then
      'Para el form de parametros internos se asigna la Empresa Primaria
       mvarControlData.Empresa = GetSPMProperty(DBSEmpresaPrimaria)
   End If
   
End Property


Public Property Get ControlData() As DataShare.udtControlData
    ControlData = mvarControlData
End Property

Private Sub SetObjects()

   DoEvents

   With ErrorLog
      .Form = Me.Name
      .Empresa = mvarControlData.Empresa
   End With

   objTabla.ControlData = mvarControlData

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   MDIExtend1.KeyDown KeyCode, Shift
End Sub

