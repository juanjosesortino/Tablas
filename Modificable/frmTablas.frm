VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTablas 
   Caption         =   "Registro del Sistema"
   ClientHeight    =   8325
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11715
   Icon            =   "frmTablas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2385
      ScaleHeight     =   180
      ScaleWidth      =   4980
      TabIndex        =   12
      Top             =   6165
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.PictureBox picTitles 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7845
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   585
      Width           =   7845
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Claves"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Tag             =   " Vista Árbol:"
         Top             =   0
         Width           =   2010
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parámetros"
         Height          =   270
         Index           =   1
         Left            =   2085
         TabIndex        =   10
         Tag             =   " Vista Lista:"
         Top             =   0
         Width           =   3210
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   5580
      TabIndex        =   7
      Top             =   1620
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdImportExport 
      Caption         =   "&Exporta"
      Height          =   360
      Left            =   5580
      TabIndex        =   6
      Top             =   2115
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picBasura 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   5445
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   600
      ScaleWidth      =   510
      TabIndex        =   5
      Top             =   2925
      Visible         =   0   'False
      Width           =   510
   End
   Begin MSComctlLib.ListView lvwImportExport 
      Height          =   1500
      Left            =   6345
      TabIndex        =   4
      Top             =   3465
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
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   6030
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1710
      Top             =   1365
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
            Picture         =   "frmTablas.frx":0442
            Key             =   "LibroCerrado"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":0894
            Key             =   "LibroAbierto"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":0CE6
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":10FA
            Key             =   "Import"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":120C
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":165E
            Key             =   "Clave"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":1AB0
            Key             =   "Parametro"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":1F02
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":2354
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":2968
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":2A7C
            Key             =   "Links"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTablas.frx":2BDC
            Key             =   "AnalisisRelaciones"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwParametros 
      Height          =   4395
      Left            =   1980
      TabIndex        =   1
      Top             =   1035
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
      Left            =   0
      TabIndex        =   2
      Top             =   1020
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
   Begin MSComctlLib.Toolbar tbr1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NuevaClave"
            Object.ToolTipText     =   "Agregar clave"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NuevoParametro"
            Object.ToolTipText     =   "Agregar Parámetro"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar clave o parámetro"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Importar"
            Object.ToolTipText     =   "Importar claves"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exportar"
            Object.ToolTipText     =   "Exportar claves"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualizar "
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir claves"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "VerReservadas"
            Object.ToolTipText     =   "Ver Claves Reservadas"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8055
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20161
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   7980
      Top             =   2370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   8010
      Top             =   1680
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   4275
      MousePointer    =   9  'Size W E
      Top             =   420
      Width           =   150
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevaClave 
         Caption         =   "Nueva &Clave"
      End
      Begin VB.Menu mnuNuevoParametro 
         Caption         =   "Nuevo &Parámetro"
      End
      Begin VB.Menu mnuEliminarClaveParametro 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuVerReservadas 
         Caption         =   "&Ver Claves Reservadas"
         Shortcut        =   ^V
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprimirArchivo 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Cerra&r"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edición"
      Begin VB.Menu mnuFind 
         Caption         =   "&Buscar ..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Buscar &Siguiente"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowItems 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuImport 
         Caption         =   "&Importar Claves"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Exportar Claves"
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActualizar 
         Caption         =   "&Actualizar"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuExpandir 
         Caption         =   "&Expandir"
      End
      Begin VB.Menu mnuNuevo 
         Caption         =   "&Nuevo"
         Begin VB.Menu mnuClave 
            Caption         =   "&Clave"
         End
         Begin VB.Menu mnuValor 
            Caption         =   "&Parámetro"
         End
      End
      Begin VB.Menu mnuModificar 
         Caption         =   "&Editar Parámetro"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnuRenombrar 
         Caption         =   "&Renombrar"
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir"
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

Private ReportObject          As clsReport                  'instancia de la clase report
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
     
   For ix = 1 To tbr1.Buttons.Count
      tbr1.Buttons(ix).Enabled = True
   Next ix
     
   SizeControls imgSplitter.Left
  
End Sub

Private Sub cmdImportExport_Click()
Dim iFile         As Integer
Dim itmx          As ListItem
Dim strKey        As String
Dim strRelatedKey As String
Dim strValor      As String
Dim strComentario As String
Dim strReservada  As String
Dim strFileName   As String
Dim aKeys()
Dim aKeys1()
   
   On Error GoTo GestErr

   Select Case cmdImportExport.Caption
      Case "&Exportar"
        
         On Error GoTo Cancel
         
         With cdg1
            .DialogTitle = Me.Caption
      
            .CancelError = True
               
               .Filter = "Archivos de Tablas *.tds|*.tds"
            cdg1.FilterIndex = 1
            .flags = cdlOFNHideReadOnly + cdlOFNExtensionDifferent + cdlOFNOverwritePrompt
               .DefaultExt = ".tds"
               .filename = "Tablas.tds"
            
            .ShowSave
         End With
        
         strFileName = cdg1.filename
         
         'no especifica el nombre del file
         If Len(strFileName) = 0 Then Exit Sub
         
         'verifica existencia del file
         If Len(Dir(strFileName)) > 0 Then
            If (MsgBox("El archivo ya existe, desea reemplazarlo ?", vbQuestion + vbYesNo, App.ProductName) = vbNo) Then Exit Sub
         End If
         iFile = FreeFile(0)
         Open strFileName For Output Lock Write As #iFile
           
         For Each itmx In lvwImportExport.ListItems
            Print #iFile, "[" & itmx.Text & "]&" & IIf(Not IsNull(itmx.SubItems(1)), "VALOR:" & itmx.SubItems(1), NullString) & IIf(Len(itmx.SubItems(2)) > 0, "&COMENTARIO:" & itmx.SubItems(2), "&COMENTARIO:") & IIf(IsNull(itmx.SubItems(3)), "&RESERVADO:", "&RESERVADO:" & itmx.SubItems(3))
         Next itmx
         
         Close #iFile
         
         MsgBox "Exportación Completada", vbOKOnly, Title
         
         'escondo listview inferior
         cmdCancelar_Click
       
      Case "&Importar"
     
        On Error Resume Next
        
        For Each itmx In lvwImportExport.ListItems
      
           ' convierte el path en clave
            strKey = itmx.Text
            strValor = itmx.SubItems(1)
            strComentario = itmx.SubItems(2)
            strReservada = itmx.SubItems(3)
            
            strReservada = IIf(strReservada = "Reservado", si, No)
            
            objTabla.GetKeys strKey, aKeys
                  
               objTabla.CreateKey mvarControlData.Empresa, itmx.Text, strValor, strComentario, strReservada
            If Err.Number <> 0 Then
              If (MsgBox("La clave " & itmx.Text & " ya existe." & Chr(13) & " Desea Actualizarla ?", vbYesNo, App.ProductName) = vbYes) Then
                 objTabla.UpdateKey EliminaAcentos(strKey), strValor, strComentario, strReservada
              End If
              Err.Clear
            End If
          
        Next itmx
        
        On Error GoTo GestErr
        
        mnuActualizar_Click
      
        MsgBox "Importación Completada", vbOKOnly, Title
      
'        cmdCancelar_Click
       
      Case "&Imprimir"
          
         If lvwImportExport.ListItems.Count = 0 Then
            MsgBox "Es necesario seleccionar las claves que se desean imprimir", vbInformation, Title
            Exit Sub
         End If
         
         With ReportObject
            
            If .ReportForm.WindowState <> vbMinimized Then
            
               OpenRecordset
               
               If rstPrint.RecordCount = 0 Then
                  MsgBox "No hay registros para imprimir", vbOKOnly, App.ProductName
                  Exit Sub
               End If
               
               stb1.Panels(STB_PANEL1).Text = STATE_PRINTING
            
               Set ReportFile = New cryTablas
               ReportFile.txtTitulo.SetText Title
               
               'seteo el titulo de la vista previa
               .Title = Title
               'asigno el report
               Set .Report = ReportFile
               'asigno el recordset
               Set .Recordset = rstPrint
               
               .Preview = True
               .ControlData = mvarControlData
               .ShowReport
               
               stb1.Panels(STB_PANEL1).Text = STATE_NONE
            
               Set ReportFile = Nothing
            Else
               'si el report esta minimizado, lo levanto
               .ReportForm.WindowState = vbNormal
               .ReportForm.Visible = True
            End If
         
         End With

      Case "&Aceptar"
      
         Dim objRelacion As Object
         Dim objRelaciones As Object
      
         strKey = Trim(Mid(lvwImportExport.ColumnHeaders(1).Text, InStr(lvwImportExport.ColumnHeaders(1).Text, ":") + 1))
         objTabla.GetKeys strKey, aKeys
         
         Set objRelaciones = New clsTablasRelaciones
         
         For Each itmx In lvwImportExport.ListItems
         
            Set objRelacion = New clsTablasRelacion
            
            ' convierte el path en clave
            
            strRelatedKey = itmx.Text
            objTabla.GetKeys strRelatedKey, aKeys1
            
            Select Case UBound(aKeys)
              Case 1
                objRelacion.Clave1 = aKeys(1)
              Case 2
                objRelacion.Clave1 = aKeys(1)
                objRelacion.Clave2 = aKeys(2)
              Case 3
                objRelacion.Clave1 = aKeys(1)
                objRelacion.Clave2 = aKeys(2)
                objRelacion.Clave3 = aKeys(3)
              Case 4
                objRelacion.Clave1 = aKeys(1)
                objRelacion.Clave2 = aKeys(2)
                objRelacion.Clave3 = aKeys(3)
                objRelacion.Clave4 = aKeys(4)
              Case 5
                objRelacion.Clave1 = aKeys(1)
                objRelacion.Clave2 = aKeys(2)
                objRelacion.Clave3 = aKeys(3)
                objRelacion.Clave4 = aKeys(4)
                objRelacion.Clave5 = aKeys(5)
            End Select
            
            Select Case UBound(aKeys1)
              Case 1
                objRelacion.LinkClave1 = aKeys1(1)
              Case 2
                objRelacion.LinkClave1 = aKeys1(1)
                objRelacion.LinkClave2 = aKeys1(2)
              Case 3
                objRelacion.LinkClave1 = aKeys1(1)
                objRelacion.LinkClave2 = aKeys1(2)
                objRelacion.LinkClave3 = aKeys1(3)
              Case 4
                objRelacion.LinkClave1 = aKeys1(1)
                objRelacion.LinkClave2 = aKeys1(2)
                objRelacion.LinkClave3 = aKeys1(3)
                objRelacion.LinkClave4 = aKeys1(4)
              Case 5
                objRelacion.LinkClave1 = aKeys1(1)
                objRelacion.LinkClave2 = aKeys1(2)
                objRelacion.LinkClave3 = aKeys1(3)
                objRelacion.LinkClave4 = aKeys1(4)
                objRelacion.LinkClave5 = aKeys1(5)
            End Select
            objRelacion.TipoRelacion = IIf(itmx.SubItems(1) = "Directa", "D", "I")
            objRelaciones.Add objRelacion
            
         Next itmx
         
         Set objTabla.Relaciones = objRelaciones
         objTabla.SaveRelations
         
         cmdCancelar_Click
         
      Case "&Interrumpir", "&Cerrar"
   
         cmdCancelar_Click
         
   End Select
  
   Exit Sub

Cancel:
   Exit Sub
GestErr:
   LoadError ErrorLog, "cmdImportExport_Click"
   ShowErrMsg ErrorLog
End Sub

Private Sub Form_Load()
Dim retval As Long

   On Error GoTo GestErr
   
   Me.Icon = LoadPicture(Icons & "Forms.ico")
   DoEvents
   
   ' En el evento Load definir procedimientos que sean independientes de la Empresa y/o MenuKey

   Set pBar.Canvas = Picture1

   Set ReportObject = New clsReport
   
   ' contruyo el Menu Ventana
   Me.mnuWindowItems(0).Caption = "&Cascada"
   Load Me.mnuWindowItems(1)
   Me.mnuWindowItems(1).Caption = "&Mosaico Horizontal"
   Load Me.mnuWindowItems(2)
   Me.mnuWindowItems(2).Caption = "Mosaico &Vertical"
   
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
      ElseIf sglPos > Me.Width - sglSplitLimit Then
         picSplitter.Left = Me.Width - sglSplitLimit
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
   If X > (Me.Width - 1500) Then X = Me.Width - 1500
   tvwClaves.Width = X
   imgSplitter.Left = X
   lvwParametros.Left = X + 40
   lvwParametros.Width = Me.Width - (tvwClaves.Width + 140)
   lblTitle(0).Width = tvwClaves.Width
   lblTitle(1).Left = lvwParametros.Left + 20
   lblTitle(1).Width = lvwParametros.Width - 40
   
   picTitles.Width = ScaleWidth
   picTitles.Top = IIf(tbr1.Visible, tbr1.Top + tbr1.Height, 0)
   tvwClaves.Top = picTitles.Top + picTitles.Height
   
   lvwParametros.Top = tvwClaves.Top
   
   'establecer el alto
   If stb1.Visible Then
      tvwClaves.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + stb1.Height)
   Else
      tvwClaves.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
   End If
   
   lvwParametros.Height = tvwClaves.Height
   imgSplitter.Top = tvwClaves.Top
   imgSplitter.Height = tvwClaves.Height
   
   With cmdImportExport
      .Top = Me.ScaleHeight - 800
      .Left = Me.ScaleWidth - 1400
   End With
   
   With cmdCancelar
      .Top = cmdImportExport.Top
      .Left = cmdImportExport.Left - 1100
   End With
   
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

Private Sub lvwImportExport_OLEDragDrop(data As MSComCtlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
            If CUsuario.SysAdmin Then
               itmx.SubItems(3) = IIf(IsNull(rstTablas(FDRecord.RESERVADA)), No, rstTablas(FDRecord.RESERVADA))
               If itmx.SubItems(3) = No Then
                  itmx.SubItems(3) = NullString
               Else
                  itmx.SubItems(3) = "Reservado"
               End If
            End If
         End If
         itmx.Key = "K" & CStr(strKey)
      Else
         MsgBox "La clave " & strKey & " ya existe en la lista", vbOKOnly, App.ProductName
      End If
      
      rstTablas.MoveNext
   Loop
         
   rstTablas.Filter = adFilterNone
   
   Exit Sub
   
GestErr:
   LoadError ErrorLog, "lvwImportExport_OLEDragDrop"
   ShowErrMsg ErrorLog

End Sub

Private Sub lvwImportExport_OLEStartDrag(data As MSComCtlLib.DataObject, AllowedEffects As Long)

   ' paso la propiedad Key de el nodo que esta siendo dragged
   ' (este valor no es usado, actualmente paso cualquier cosa)
   data.SetData SourceItem.Key
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
   itmx.Key = "K" & CStr(NewString)
   
   rstTablas.Filter = adFilterNone
  
   Exit Sub
   
GestErr:
   LoadError ErrorLog, "lvwParametros_AfterLabelEdit"
   ShowErrMsg ErrorLog
   
   Cancel = True
   rstTablas.Filter = adFilterNone
End Sub

'Private Sub lvwParametros_Click()
'
''   If strCurrentOperation = "Relacion" Then mnuRelaciones_Click
'
'   lvwParametros_DblClick
'
'End Sub

Private Sub lvwParametros_DblClick()
Dim strKey  As String
Dim frm     As New frmDialog
Dim itmx    As ListItem
Dim aKeys()

   If TypeName(lvwParametros.SelectedItem) = "Nothing" Then Exit Sub
   
   On Error GoTo GestErr
      
   ' obtengo la clave del parametro
   strKey = GetParameterKey(lvwParametros.SelectedItem)
   
   objTabla.GetKeys strKey, aKeys
   
   'obtengo el record relativo a la clave seleccionada
   Select Case UBound(aKeys)
    Case 2
      rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND (" & FDRecord.Clave2 & " = '" & aKeys(2) & "')"
    Case 3
      rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND (" & FDRecord.Clave3 & " = '" & aKeys(3) & "')"
    Case 4
      rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND (" & FDRecord.Clave4 & " = '" & aKeys(4) & "')"
    Case 5
      rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND (" & FDRecord.Clave4 & " = '" & aKeys(4) & "') AND (" & FDRecord.Clave5 & " = '" & aKeys(5) & "')"
   End Select
      
   frm.DialogTitle = "Edición de Parámetros"
   frm.ShowPrintButtons = False
   
   frm.AddFrame "F1", ""

   frm.AddLabel "lbl1", "Nombre", 100, 200
   frm.AddText "txtNombre", "", 0, 400, 0, 100, 6000, , lvwParametros.SelectedItem.Text, , , , , True, False
   frm.AddLabel "lbl2", "Valor", 100, 1000
   frm.AddText "txtValor", "", 0, 1200, 0, 100, 6000, , lvwParametros.SelectedItem.ListSubItems(1)
   frm.AddLabel "lbl3", "Comentario", 100, 1800
   frm.AddText "txtComentario", "", 0, 2000, 0, 100, 6000, , lvwParametros.SelectedItem.ListSubItems(2)
   
   If CUsuario.SysAdmin Then
      frm.AddCheck "chkReservado", "Reservado", 2600, 100, IIf(lvwParametros.SelectedItem.ListSubItems(3) = "Reservado", vbChecked, vbUnchecked)
   End If
   
   frm.ShowDialog
   
   If frm.ButtonPressed = Cancel Then Exit Sub
   
   Me.Refresh
   
   If TypeName(objTabla) = "clsParametrosInternos" Then
      objTabla.UpdateKey strKey, frm.Value("txtValor"), frm.Value("txtComentario"), IIf(frm.Value("chkReservado") = vbChecked, si, No)
   Else
      objTabla.UpdateKey "", strKey, frm.Value("txtValor"), frm.Value("txtComentario"), IIf(frm.Value("chkReservado") = vbChecked, si, No)
   End If

   'actualizo el listview
   Set itmx = lvwParametros.SelectedItem
   itmx.Text = frm.Value("txtNombre")
   itmx.SubItems(1) = frm.Value("txtValor")
   itmx.SubItems(2) = frm.Value("txtComentario")
   itmx.SubItems(3) = IIf(frm.Value("chkReservado") = vbChecked, "Reservado", NullString)
   itmx.Key = "K" & CStr(itmx.Text)
   
   'actualizo el recordset
   rstTablas(FDRecord.Valor) = itmx.SubItems(1)
   rstTablas(FDRecord.COMENTARIO) = itmx.SubItems(2)
   If itmx.SubItems(3) = NullString Then
      rstTablas(FDRecord.RESERVADA) = No
   Else
      rstTablas(FDRecord.RESERVADA) = si
   End If
   rstTablas.Update
   
   Exit Sub
   
GestErr:
   LoadError ErrorLog, "lvwParametros_DblClick"
   ShowErrMsg ErrorLog
End Sub

Private Sub lvwParametros_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then
      mnuEliminar_Click
   End If
End Sub

Private Sub lvwParametros_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
   If strCurrentOperation <> "Relacion" Then Exit Sub
   
   ' verifico si hemos iniciado una operacion de drag-drop
   If Button <> vbLeftButton Then Exit Sub
   ' seteo el nodo que esta siendo dragged, me voy si no hay ninguno
   Set SourceItem = lvwParametros.HitTest(X, Y)
   If SourceItem Is Nothing Then Exit Sub
   
   ShiftState = Shift
   
   ' salvo valores para mas adelante
   Set SourceObject = lvwParametros
   ' inicio operacion de drag
   lvwParametros.OLEDrag

End Sub

Private Sub lvwParametros_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
   If Button = vbRightButton Then
      mnuExpandir.Visible = False
      mnuNuevo.Visible = False
      mnuModificar.Visible = True
      
      mnuEliminar.Enabled = True
      mnuRenombrar.Enabled = True
      mnuModificar.Enabled = True
      
      SetPerms
      
      PopupMenu mnuMenu, vbPopupMenuRightButton
   End If

End Sub

Private Sub lvwParametros_OLEStartDrag(data As MSComCtlLib.DataObject, AllowedEffects As Long)
   
   ' paso la propiedad Key de el nodo que esta siendo dragged
   ' (este valor no es usado, actualmente paso cualquier cosa)
   data.SetData SourceItem.Key
   If ShiftState And vbCtrlMask Then
       AllowedEffects = vbDropEffectCopy
   Else
       AllowedEffects = vbDropEffectMove
   End If

End Sub

Private Sub mnuActualizar_Click()
   
   LoadTablas
   
End Sub
Private Sub LoadTablas()
Dim SQL As String

   Set rstTablas = Fetch("ALG", strSQL)
   
   ConstruyeTreeView
      
End Sub

Private Sub mnuClave_Click()
Dim strKey     As String
Dim strInput   As String
'Dim itmX       As ListItem
Dim nodX       As Node
Dim aKeys()

   On Error GoTo GestErr

   If tvwClaves.SelectedItem Is Nothing Then
      Set tvwClaves.SelectedItem = tvwClaves.Nodes(1)
   End If
   
   'si estoy en el 4 nivel, no permite la creacion de nuevas claves
   strKey = GetNodeKey(tvwClaves.SelectedItem)
   
   objTabla.GetKeys strKey, aKeys
   If UBound(aKeys) = 4 Then
      MsgBox "No se permite la creaciòn de nuevas claves. El nivel máximo permitido es 4", vbOKOnly, Title
      Exit Sub
   End If
   
   strInput = InputBox("Ingrese nombre de la clave ", Title)
   If Len(strInput) = 0 Then Exit Sub
   
   strInput = EliminaAcentos(strInput)
   
   If InStr(tvwClaves.SelectedItem.Key, "\") > 0 Then
      strKey = GetNodeKey(tvwClaves.SelectedItem) & "\" & strInput
   Else
      strKey = strInput
   End If
   'actualizo la base
   If TypeName(objTabla) = "clsParametrosInternos" Then
      objTabla.CreateKey strKey & "\" & "(Predeterminado)", "NuevoValor", "", "No"
   Else
      objTabla.CreateKey "", strKey & "\" & "(Predeterminado)", "NuevoValor", "", "No"
   End If
   
   'actualizo el árbol
   If TypeName(tvwClaves.SelectedItem) = "Nothing" Then
      'agrego nodo a la raiz
      Set nodX = tvwClaves.Nodes.Add(KeyROOT, tvwChild, strKey, strKey)
   Else
      'agrego nodo hijo al nodo seleccionado
      If tvwClaves.SelectedItem.Key = KeyROOT Then
         Set nodX = tvwClaves.Nodes.Add(KeyROOT, tvwChild, tvwClaves.SelectedItem.Key & "\" & strInput, strInput)
      Else
         On Error Resume Next
         Set nodX = tvwClaves.Nodes.Add(tvwClaves.SelectedItem.Key, tvwChild, tvwClaves.SelectedItem.Key & "\" & strInput, strInput)
         If Err.Number <> 0 Then
            MsgBox "Clave duplicada.", vbExclamation, Title
            Exit Sub
         End If
      End If
   End If
   
   nodX.Image = "LibroCerrado"
   nodX.ExpandedImage = "LibroAbierto"
   'defino nodo seleccionado al nodo apenas agregado
   Set tvwClaves.SelectedItem = nodX
   
   'actualizo el ListView
   lvwParametros.ListItems.Clear
     
   'actualizo el recordset
   strKey = GetNodeKey(nodX)
   objTabla.GetKeys strKey, aKeys
         
   rstTablas.AddNew
   Select Case UBound(aKeys)
     Case 1
       rstTablas(FDRecord.Clave1) = aKeys(1)
       rstTablas(FDRecord.Clave2) = "(Predeterminado)"
       rstTablas(FDRecord.Valor) = "NuevoValor"
   
     Case 2
       rstTablas(FDRecord.Clave1) = aKeys(1)
       rstTablas(FDRecord.Clave2) = aKeys(2)
       rstTablas(FDRecord.Clave3) = "(Predeterminado)"
       rstTablas(FDRecord.Valor) = "NuevoValor"
   
     Case 3
       rstTablas(FDRecord.Clave1) = aKeys(1)
       rstTablas(FDRecord.Clave2) = aKeys(2)
       rstTablas(FDRecord.Clave3) = aKeys(3)
       rstTablas(FDRecord.Clave4) = "(Predeterminado)"
       rstTablas(FDRecord.Valor) = "NuevoValor"
   
     Case 4
       rstTablas(FDRecord.Clave1) = aKeys(1)
       rstTablas(FDRecord.Clave2) = aKeys(2)
       rstTablas(FDRecord.Clave3) = aKeys(3)
       rstTablas(FDRecord.Clave4) = aKeys(4)
       rstTablas(FDRecord.Clave5) = "(Predeterminado)"
       rstTablas(FDRecord.Valor) = "NuevoValor"
   
   End Select
   rstTablas.Update
 
   Exit Sub
   
GestErr:
'   LoadError ErrorLog, "lvwParametros_AfterLabelEdit"
'   ShowErrMsg ErrorLog
End Sub

Private Sub mnuEliminar_Click()
'Dim nodX       As Node
Dim itmx       As ListItem
Dim strKey     As String
Dim strWhere   As String
Dim aKeys()
Dim alvKeys()  As String
Dim ix         As Integer

   On Error GoTo GestErr

   If Me.ActiveControl.Name = "lvwParametros" Then
   
      'elimina un parametro del listview
      If TypeName(lvwParametros.SelectedItem) = "Nothing" Then Exit Sub
       
      'determino si se eligio borrar uno o mas parametros
      For Each itmx In lvwParametros.ListItems
         If itmx.Selected = True Then ix = ix + 1
         If ix > 1 Then Exit For
      Next itmx
         
      If ix > 1 Then
         If (MsgBox("Confirma la eliminación de los parámetros seleccionados ?", vbQuestion + vbYesNo, App.ProductName) = vbNo) Then Exit Sub
      Else
         If (MsgBox("Confirma la eliminación del parámetro " & lvwParametros.SelectedItem.Text & " ?", vbQuestion + vbYesNo, App.ProductName) = vbNo) Then Exit Sub
      End If
       
      For Each itmx In lvwParametros.ListItems
         If itmx.Selected = True Then
           
            ' guardo las claves de los items que deben ser eliminados
            If IsArrayEmpty(alvKeys) Then
               ReDim alvKeys(0)
            Else
               ReDim Preserve alvKeys(UBound(alvKeys) + 1)
            End If
            alvKeys(UBound(alvKeys)) = itmx.Key
               
            strKey = GetParameterKey(itmx)
             
            'elimino la clave de la base
            If TypeName(objTabla) = "clsParametrosInternos" Then
               objTabla.DeleteKey strKey
            Else
               objTabla.DeleteKey "", strKey
            End If
             
            'obtengo la clave del record decifrando la clave del nodo y la meto en aKeys
             objTabla.GetKeys strKey, aKeys
             
             Select Case UBound(aKeys)
                 
               Case 2
                 rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "')"
                 rstTablas.Delete
             
               Case 3
                 rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "')"
                 rstTablas.Delete
             
               Case 4
                 rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND  (" & FDRecord.Clave4 & " = '" & aKeys(4) & "')"
                 rstTablas.Delete
             
               Case 5
                 rstTablas.Filter = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND  (" & FDRecord.Clave4 & " = '" & aKeys(4) & "') AND  (" & FDRecord.Clave5 & " = '" & aKeys(5) & "')"
                 rstTablas.Delete
             
             End Select
          
             rstTablas.Filter = adFilterNone
         
         End If
         
      Next itmx
      
      'elimino los items
      For ix = 0 To UBound(alvKeys)
         lvwParametros.ListItems.Remove alvKeys(ix)
      Next ix
   
   End If

   If Me.ActiveControl.Name = "tvwClaves" Then
       'elimina una clave del treeview
       
       If (MsgBox("Confirma la eliminación de la clave " & tvwClaves.SelectedItem.FullPath & " y eventuales subclaves ?", vbQuestion + vbYesNo, App.ProductName) = vbNo) Then Exit Sub
       
       strKey = GetNodeKey(tvwClaves.SelectedItem)
         
      'elimino la clave de la base
      If TypeName(objTabla) = "clsParametrosInternos" Then
         objTabla.DeleteKey strKey
      Else
         objTabla.DeleteKey "", strKey
      End If
      
      'actualizo el recordset
      objTabla.GetKeys strKey, aKeys
       
       Select Case UBound(aKeys)
         Case 1
           strWhere = FDRecord.Clave1 & " = '" & aKeys(1) & "'"
         Case 2
           strWhere = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "')"
         Case 3
           strWhere = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "')"
         Case 4
           strWhere = "(" & FDRecord.Clave1 & " = '" & aKeys(1) & "') AND  (" & FDRecord.Clave2 & " = '" & aKeys(2) & "') AND  (" & FDRecord.Clave3 & " = '" & aKeys(3) & "') AND  (" & FDRecord.Clave4 & " = '" & aKeys(4) & "')"
       End Select
       
       rstTablas.Filter = strWhere
       
       Do While Not rstTablas.EOF
          rstTablas.Delete
          rstTablas.MoveNext
       Loop
       
       'actualizo el treeView
       Do While (tvwClaves.SelectedItem.Children > 0)
         tvwClaves.Nodes.Remove tvwClaves.SelectedItem.Child.Key
       Loop
       tvwClaves.Nodes.Remove tvwClaves.SelectedItem.Key
       
       'actualizo el ListView
       lvwParametros.ListItems.Clear
       
   End If
      
   Exit Sub
   
GestErr:
   LoadError ErrorLog, "mnuEliminar_Click"
   ShowErrMsg ErrorLog
End Sub

Private Sub mnuEliminarClaveParametro_Click()
   mnuEliminar_Click
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuExport_Click()
Dim oldHeight As Integer

   strCurrentOperation = "Exportar"
   
   If lvwImportExport.Visible Then Exit Sub
   lvwImportExport.LabelEdit = lvwAutomatic
   
   oldHeight = tvwClaves.Height
   tvwClaves.Height = tvwClaves.Height / 2
   lvwParametros.Height = tvwClaves.Height
   
   With lvwImportExport
      .Height = oldHeight - tvwClaves.Height - 700
      .Top = tvwClaves.Top + tvwClaves.Height
      .Width = Me.ScaleWidth
      .ColumnHeaders(1).Width = 5000
      .ColumnHeaders(2).Width = 4000
      .ColumnHeaders(3).Width = 3000
      .Visible = True
      .Left = 0
      .ZOrder 0
   
      .ColumnHeaders(1).Text = "Claves exportadas"
      .ColumnHeaders(2).Text = "Valor"
      .ColumnHeaders(3).Text = "Comentario"
      .ColumnHeaders(4).Text = "Uso"
   
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
      .Caption = "&Exportar"
      .Visible = True
   End With
   
   cmdCancelar.Visible = True
  
End Sub

'Private Sub mnuRelaciones_Click()
'Dim oldHeight As Integer
'Dim itmX As ListItem
'Dim strKey As String
'
'   On Error GoTo GestErr
'
'   If lvwParametros.SelectedItem Is Nothing Then
'      MsgBox "Las claves no admiten relaciones", vbOKOnly, App.ProductName
'      Exit Sub
'   End If
'
'   strCurrentOperation = "Relacion"
'   lvwImportExport.LabelEdit = lvwManual
'
'   If lvwImportExport.Visible Then Exit Sub
'
'   oldHeight = tvwClaves.Height
'   tvwClaves.Height = tvwClaves.Height / 2
'   lvwParametros.Height = tvwClaves.Height
'
'   With lvwImportExport
'
'      .Height = oldHeight - tvwClaves.Height - 700
'      .Top = tvwClaves.Top + tvwClaves.Height
'      .Width = Me.ScaleWidth
'      .ColumnHeaders(1).Width = 10000
'      .ColumnHeaders(2).Width = 1500
'      .ColumnHeaders(3).Width = 0
'      .ColumnHeaders(4).Width = 0
'      .Visible = True
'      .Left = 0
'      .ZOrder 0
'
'      With picBasura
'         .Left = 500
'         .Top = lvwImportExport.Top + lvwImportExport.Height + 100
'         .Visible = True
'         .ZOrder 0
'         .Picture = LoadPicture(BitMaps & "Cesto.bmp")
'      End With
'
'      With cmdImportExport
'         .Caption = "&Aceptar"
'         .Visible = True
'      End With
'
'      cmdCancelar.Visible = True
'
'      strKey = Replace(tvwClaves.SelectedItem.FullPath, KeyROOT & "\", "") & "\" & lvwParametros.SelectedItem.Text
'
'      .ColumnHeaders(1).Text = "Parámetros relacionados con: " & strKey
'      .ColumnHeaders(2).Text = "Relación"
'
'      .ListItems.Clear
'
'      'cargo la clave y sus relaciones
'      If TypeName(objTabla) = "clsParametrosInternos" Then
'         objTabla.LoadKey strKey
'      Else
'         objTabla.LoadKey "", strKey
'      End If
'      objTabla.LoadRelations
'
'      Dim objRelacion As Object
'
'      For Each objRelacion In objTabla.Relaciones
'         Set itmX = .ListItems.Add
'         itmX.Text = objRelacion.ClaveRelacionada
'         itmX.SubItems(1) = IIf(objRelacion.TipoRelacion = "D", "Directa", "Indirecta")
'         itmX.Key = "K" & CStr(itmX.Text)
'      Next objRelacion
'
'   End With
'
'   Exit Sub
'
'GestErr:
'   LoadError ErrorLog, "mnuRelaciones"
'   ShowErrMsg ErrorLog
'End Sub
'
'Private Sub mnuTodasRelaciones_Click()
'Dim oldHeight As Integer
'Dim itmX As ListItem
'Dim strKey As String
'Dim ix As Integer
'
'   strCurrentOperation = ""
'   lvwImportExport.LabelEdit = lvwManual
'
'   If lvwImportExport.Visible Then Exit Sub
'
'   oldHeight = tvwClaves.Height
'
'   tvwClaves.Visible = False
'   lvwParametros.Visible = False
'
'   For ix = 1 To 13
'      tbr1.Buttons(ix).Enabled = False
'   Next ix
'
'   With lvwImportExport
'
'      .Height = picTitles.Top + picTitles.Height + tvwClaves.Height - 1000
'      .Top = picTitles.Top
'      .Width = Me.ScaleWidth
'      .ColumnHeaders(1).Width = 5000
'      .ColumnHeaders(2).Width = 5000
'      .ColumnHeaders(3).Width = 1500
'      .ColumnHeaders(4).Width = 0
'      .Visible = True
'      .Left = 0
'      .ZOrder 0
'
'      With cmdImportExport
'         .Caption = "&Interrumpir"
''         .Top = lvwImportExport.Height + 500
''         .Left = lvwImportExport.Width - 1400
'         .Visible = True
'      End With
'
'      Picture1.Height = stb1.Height - 80
'      Picture1.Top = stb1.Top + 40
'      Picture1.Left = TextWidth("Analizando relaciones...") + 100
'      Picture1.Width = stb1.Panels(1).Width - TextWidth("Analizando relaciones...")
'
'      With cmdCancelar
'         .Visible = False
'      End With
'
'      .ColumnHeaders(1).Text = "Parámetro"
'      .ColumnHeaders(2).Text = "Parámetros Relacionados"
'      .ColumnHeaders(3).Text = "Relación"
'
'      .ListItems.Clear
'
'      pBar.Visible = True
'      pBar.Min = 0
'      pBar.Max = IIf(rstTablas.RecordCount > 0, rstTablas.RecordCount, 1)
'      stb1.Panels(1).Text = "Analizando relaciones..."
'
'      rstTablas.MoveFirst
'      bEndOfAnalisis = False
'      Do While Not rstTablas.EOF
'
'         If bEndOfAnalisis Then Exit Do
'
'         strKey = Clave(rstTablas(FDRecord.Clave1), rstTablas(FDRecord.Clave2), rstTablas(FDRecord.Clave3), _
'                        rstTablas(FDRecord.Clave4), rstTablas(FDRecord.Clave5))
'
'         If InStr(strKey, "(Predeterminado)") = 0 Then
'
'
'            'cargo la clave y sus relaciones
'            objTabla.LoadKey "", strKey
'            If objTabla.ObjectIsLoaded Then
'               objTabla.LoadRelations
'
'               Dim objRelacion As Object
'
'               If objTabla.Relaciones.Count > 0 Then
'                  Set itmX = .ListItems.Add
'                  itmX.Text = strKey
'                  ix = 1
'                  For Each objRelacion In objTabla.Relaciones
'                     If ix > 1 Then
'                        Set itmX = .ListItems.Add
'                     End If
'                     itmX.SubItems(1) = objRelacion.ClaveRelacionada
'                     itmX.SubItems(2) = IIf(objRelacion.TipoRelacion = "D", "Directa", "Indirecta")
'                     ix = ix + 1
'                  Next objRelacion
'
'                  Set itmX = .ListItems.Add
'
'               End If
'            End If
'         End If
'
'         pBar.Value = rstTablas.AbsolutePosition
'
'         rstTablas.MoveNext
'
'         DoEvents
'      Loop
'
'   End With
'
'   pBar.Value = 0
'   pBar.Visible = False
'   stb1.Panels(1).Text = NullString
'
'   cmdImportExport.Caption = "&Cerrar"
'
'End Sub

Private Sub mnuFind_Click()

   strFindText = InputBox("Indicar el texto que se desea buscar", "Registro del Sistema", strFindText)
   If Len(strFindText) = 0 Then Exit Sub
   
   strFindText = UCase(strFindText)
   
   'inicia la busqueda del primero
   StartFind
  
End Sub

Private Sub mnuFindNext_Click()

   If Len(strFindText) = 0 Then
     MsgBox "No se indicó un texto de busqueda", vbOKOnly, Title
     Exit Sub
   End If
   
   If bEndOfFind Then
     MsgBox "Fin de la búsqueda", vbOKOnly, Title
     Exit Sub
   End If
   
   FindX

End Sub

Private Sub mnuImport_Click()
Dim iOldHeight    As Integer
Dim iFile         As Integer
Dim strBuffer     As String
Dim strComentario As String
Dim strValor      As String
Dim strReservada  As String
Dim strKey        As String
Dim itmx          As ListItem
Dim aArray()      As String


   strCurrentOperation = "Importar"
   
   If lvwImportExport.Visible Then Exit Sub
   lvwImportExport.LabelEdit = lvwAutomatic
      
   'solicito archivo a importar
   On Error GoTo Cancel
   
   cdg1.CancelError = True
   cdg1.Filter = "Archivos de Tablas *.tds|*.tds"
   cdg1.FilterIndex = 1
   cdg1.flags = cdlOFNHideReadOnly + cdlOFNExtensionDifferent + cdlOFNOverwritePrompt
   cdg1.DefaultExt = ".tds"
   cdg1.filename = NullString

   cdg1.ShowOpen

   'verifica existencia del file
   If Len(Dir(cdg1.filename)) = 0 Then
      MsgBox "El archivo no existe", vbExclamation, Title
      Exit Sub
   End If
   
   'leo el archivo y cargo el listview1
   iFile = FreeFile(0)
   Open cdg1.filename For Input As #iFile
   
   lvwImportExport.ListItems.Clear
   Do Until EOF(iFile)
      Line Input #iFile, strBuffer
      Set itmx = lvwImportExport.ListItems.Add()
       
      If Len(strBuffer) > 0 Then
         aArray = Split(strBuffer, "&")
         If Len(aArray(0)) > 0 Then
            strKey = Mid(aArray(0), 2, Len(aArray(0)) - 2)
            strValor = Right(aArray(1), Len(aArray(1)) - 6)
            strComentario = Right(aArray(2), Len(aArray(2)) - 11)
            strReservada = Right(aArray(3), Len(aArray(3)) - 10)
         
            itmx.Text = strKey
            itmx.SubItems(1) = strValor
            itmx.SubItems(2) = strComentario
            itmx.SubItems(3) = strReservada
             
            itmx.Key = "K" & CStr(strKey)
         
         End If
      End If
     
   Loop
   
   'cierro archivo
   Close #iFile
   
   'redimensiono los objetos necesarios para visualizar las claves importadas
   iOldHeight = tvwClaves.Height
   tvwClaves.Height = tvwClaves.Height / 2
   lvwParametros.Height = tvwClaves.Height
   With lvwImportExport
      .Height = iOldHeight - tvwClaves.Height - 700
      .Top = tvwClaves.Top + tvwClaves.Height
      .Width = Me.ScaleWidth
      .ColumnHeaders(1).Width = 5000
      .ColumnHeaders(2).Width = 4000
      .ColumnHeaders(3).Width = 3000
      .Visible = True
      .Left = 0
      .ZOrder 0
      .ColumnHeaders(1).Text = "Claves importadas"
      .ColumnHeaders(2).Text = "Valor"
      .ColumnHeaders(3).Text = "Comentario"
      .ColumnHeaders(4).Text = "Uso"
      
   End With
   
   With picBasura
      .Left = 500
      .Top = lvwImportExport.Top + lvwImportExport.Height + 100
      .Visible = True
      .ZOrder 0
      .Picture = LoadPicture(BitMaps & "Cesto.bmp")
   End With
   
   With cmdImportExport
      .Caption = "&Importar"
'      .Top = picBasura.Top
'      .Left = lvwImportExport.Width - 1400
      .Visible = True
   End With
   
   With cmdCancelar
'      .Top = cmdImportExport.Top
'      .Left = cmdImportExport.Left - 1100
      .Visible = True
   End With
   
   
Cancel:
   Exit Sub
End Sub

Private Sub mnuImprimir_Click()
   ImprimirClaves
End Sub

Private Sub mnuModificar_Click()
   lvwParametros_DblClick
End Sub

Private Sub mnuImprimirArchivo_Click()
   ImprimirClaves
End Sub

Private Sub mnuNuevaClave_Click()
   mnuClave_Click
End Sub

Private Sub mnuNuevoParametro_Click()
   mnuValor_Click
End Sub

Private Sub mnuRenombrar_Click()

   If Me.ActiveControl.Name = "lvwParametros" Then
      If TypeName(lvwParametros.SelectedItem) = "Nothing" Then Exit Sub
      lvwParametros.StartLabelEdit
   Else
      If TypeName(tvwClaves.SelectedItem) = "Nothing" Then Exit Sub
      tvwClaves.StartLabelEdit
   End If
  
End Sub

Private Sub mnuValor_Click()
Dim strKey       As String
Dim strParametro As String
Dim itmx         As ListItem
Dim aKeys()

   On Error GoTo GestErr

   If TypeName(tvwClaves.SelectedItem) = "Nothing" Then
      MsgBox "Es necesario seleccionar una clave", vbOKOnly, Title
      Exit Sub
   End If
   
   Set frm = New frmDialog
   
   frm.DialogTitle = "Nuevo Parametro"
   frm.ShowPrintButtons = False
   frm.AddFrame "Frame", ""
   frm.AddLabel "lbl1", "Nombre", 100, 200
   frm.AddText "txtNombre", "", 0, 400, 0, 100, 6000
   frm.AddLabel "lbl2", "Valor", 100, 1000
   frm.AddText "txtValor", "", 0, 1200, 0, 100, 6000
   frm.AddLabel "lbl3", "Comentario", 100, 1800
   frm.AddText "txtComentario", "", 0, 2000, 0, 100, 6000
   
   frm.AddCheck "chkReservado", "Reservado", 2600, 100, vbUnchecked

   
   frm.ShowDialog
   
   If frm.ButtonPressed = Cancel Then
      Unload frm
      DoEvents
      Set frm = Nothing
      Exit Sub
   End If
   
   If frm.Value("txtNombre") = NullString Then Exit Sub
   
   strKey = Mid(tvwClaves.SelectedItem.Key, InStr(tvwClaves.SelectedItem.Key, "\") + 1)
   strKey = strKey & "\" & frm.Value("txtNombre")
   objTabla.GetKeys strKey, aKeys

   If TypeName(objTabla) = "clsParametrosInternos" Then
      objTabla.CreateKey strKey, frm.Value("txtValor"), frm.Value("txtComentario"), IIf(frm.Value("chkReservado") = vbChecked, si, No)
   Else
      objTabla.CreateKey "", strKey, frm.Value("txtValor"), frm.Value("txtComentario"), IIf(frm.Value("chkReservado") = vbChecked, si, No)
   End If
   
   'actualizo el listview
   Set itmx = lvwParametros.ListItems.Add()
   itmx.Text = frm.Value("txtNombre")
   itmx.SubItems(1) = frm.Value("txtValor")
   itmx.SubItems(2) = frm.Value("txtComentario")
   itmx.SubItems(3) = IIf(frm.Value("chkReservado") = vbChecked, "Reservado", NullString)
   itmx.Key = "K" & CStr(itmx.Text)
   
   strParametro = frm.Value("txtNombre")

   'actualizo el recordset
   strKey = GetNodeKey(tvwClaves.SelectedItem)
   strKey = strKey & "\" & strParametro
   objTabla.GetKeys strKey, aKeys
         
   rstTablas.AddNew
   Select Case UBound(aKeys)
     Case 2
       rstTablas(FDRecord.Clave1) = aKeys(1)
       rstTablas(FDRecord.Clave2) = aKeys(2)
   
     Case 3
       rstTablas(FDRecord.Clave1) = aKeys(1)
       rstTablas(FDRecord.Clave2) = aKeys(2)
       rstTablas(FDRecord.Clave3) = aKeys(3)
   
     Case 4
       rstTablas(FDRecord.Clave1) = aKeys(1)
       rstTablas(FDRecord.Clave2) = aKeys(2)
       rstTablas(FDRecord.Clave3) = aKeys(3)
       rstTablas(FDRecord.Clave4) = aKeys(4)
   
     Case 5
       rstTablas(FDRecord.Clave1) = aKeys(1)
       rstTablas(FDRecord.Clave2) = aKeys(2)
       rstTablas(FDRecord.Clave3) = aKeys(3)
       rstTablas(FDRecord.Clave4) = aKeys(4)
       rstTablas(FDRecord.Clave5) = aKeys(5)
   
   End Select
   rstTablas(FDRecord.Valor) = itmx.SubItems(1)
   rstTablas(FDRecord.COMENTARIO) = itmx.SubItems(2)
   rstTablas(FDRecord.RESERVADA) = itmx.SubItems(3)
   
   rstTablas.Update
         
   Exit Sub
   
GestErr:
'   LoadError ErrorLog, "mnuValor_Click"
'   ShowErrMsg ErrorLog
End Sub



Private Sub picBasura_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If strCurrentOperation = "Relacion" And lvwImportExport.SelectedItem.SubItems(1) = "Indirecta" Then
      MsgBox "No es posible eliminar una relación indirecta", vbOKOnly, App.ProductName
      Exit Sub
   End If

   TirarBasura

End Sub

Private Sub tbr1_ButtonClick(ByVal Button As MSComCtlLib.Button)

   Select Case Button.Index
     Case 1 'Nueva clave
       mnuClave_Click
     Case 2 'nuevo parametro
       mnuValor_Click
     Case 4 'buscar
       mnuFind_Click
     Case 6 'Import
       mnuImport_Click
     Case 7  'Export
       mnuExport_Click
     Case 9
       mnuActualizar_Click
     Case 11
       ImprimirClaves
       
       ' quito esto de relaciones que no esta en uso y no funciona de manera estable
'     Case 13
'       mnuRelaciones_Click
'     Case 14
'       mnuTodasRelaciones_Click

      Case 13
         mnuVerReservadas_Click
   End Select
  
End Sub

Private Sub mnuVerReservadas_Click()
Static bViendoClavesReservadas As Boolean

   If bViendoClavesReservadas Then Exit Sub
   
   If (MsgBox("Desea cargar Todas las Claves Reservadas?" & vbCrLf & "(Esta operación puede demorar unos minutos...)", vbQuestion + vbYesNo, App.ProductName) = vbNo) Then
      Exit Sub
   End If
   
   bViendoClavesReservadas = True
   
   strSQL = "SELECT * FROM TABLAS ORDER BY TAB_CLAVE1, TAB_CLAVE2, TAB_CLAVE3, TAB_CLAVE4, TAB_CLAVE5"
   
   mnuActualizar_Click
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
   tvwClaves.SelectedItem.Key = CStr(strNewKey)
   
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
   LoadError ErrorLog, "tvwClaves_AfterLabelEdit"
   ShowErrMsg ErrorLog
End Sub


Private Sub tvwClaves_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyDelete Then
      mnuEliminar_Click
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

Private Sub tvwClaves_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
   If tvwClaves.Nodes.Count = 0 Then Exit Sub
   
   mnuExpandir.Visible = True
   mnuNuevo.Visible = True
   
   If Button = vbRightButton Then
      If tvwClaves.SelectedItem.Text = KeyROOT Then
         mnuExpandir.Enabled = True
         mnuNuevo.Enabled = True
         mnuValor.Enabled = False
         mnuEliminar.Enabled = False
         mnuRenombrar.Enabled = False
         mnuModificar.Enabled = False
      Else
         mnuExpandir.Enabled = True
         mnuNuevo.Enabled = True
         mnuValor.Enabled = True
         mnuEliminar.Enabled = True
         mnuRenombrar.Enabled = True
         mnuModificar.Enabled = True
      End If
      
      SetPerms
   
      PopupMenu mnuMenu, vbPopupMenuRightButton
   End If
   
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

Private Sub tvwClaves_NodeClick(ByVal Node As MSComCtlLib.Node)
Dim itmx    As ListItem
Dim strKey  As String
Dim aKeys()
'Dim vValue  As Variant
   
   On Error GoTo GestErr
   
   'actualizo el nivel del nodo corriente
   objTabla.GetKeys Node.Key, aKeys
   
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
                 If itmx.SubItems(3) = si Then
                    itmx.SubItems(3) = "Reservado"
                 Else
                    itmx.SubItems(3) = NullString
                End If

               itmx.Key = "K" & CStr(itmx.Text)
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
        
        Set tvwClaves.SelectedItem = tvwClaves.Nodes(nod.Key)
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
            Set tvwClaves.SelectedItem = tvwClaves.Nodes(nod.Key)
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
   LoadError ErrorLog, "RenameKey"
   ShowErrMsg ErrorLog
   
End Function

Private Function GetParameterKey(ByVal Item As MSComCtlLib.ListItem) As String

   GetParameterKey = GetNodeKey(tvwClaves.SelectedItem)
   GetParameterKey = GetParameterKey & "\" & Mid(Item.Key, 2)

End Function

Private Function GetNodeKey(ByVal Node As MSComCtlLib.Node)
                 
   GetNodeKey = Mid(Node.Key, InStr(Node.Key, "\") + 1)

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

Private Sub tvwClaves_OLEDragDrop(data As MSComCtlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If strCurrentOperation <> "Importar" Or strCurrentOperation = "Exportar" Then Exit Sub
    
   ' chequeo que el destino es el mismo treeview
   If SourceObject Is tvwClaves Then
       MsgBox "No es posible efectuar esta operación", vbExclamation
   End If
   Set tvwClaves.DropHighlight = Nothing

End Sub

Private Sub tvwClaves_OLEStartDrag(data As MSComCtlLib.DataObject, AllowedEffects As Long)
    
    ' paso la propiedad Key de el nodo que esta siendo dragged
    ' (este valor no es usado, actualmente paso cualquier cosa)
    data.SetData SourceNode.Key
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
         lvwImportExport.ListItems.Remove itmx.Key
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
   
   Set ReportObject = Nothing
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

