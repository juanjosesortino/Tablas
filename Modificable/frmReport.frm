VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRVIEWER.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport 
   Caption         =   "Vista Previa"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7200
      Left            =   180
      TabIndex        =   37
      Top             =   765
      Width           =   8880
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   0   'False
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
   End
   Begin VB.ComboBox cmbZoom 
      Height          =   315
      ItemData        =   "frmReport.frx":0000
      Left            =   7470
      List            =   "frmReport.frx":0025
      TabIndex        =   10
      Text            =   "53%"
      ToolTipText     =   "Zoom"
      Top             =   180
      Width           =   1590
   End
   Begin VB.Frame Frame3 
      Caption         =   "Propiedades impresion"
      Height          =   5085
      Left            =   9180
      TabIndex        =   30
      Top             =   2880
      Width           =   2760
      Begin VB.CheckBox chkComentarios 
         Caption         =   "Imprime Comentarios"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4680
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   315
         ItemData        =   "frmReport.frx":0076
         Left            =   405
         List            =   "frmReport.frx":0089
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Seleccionar el tipo de papel"
         Top             =   2295
         Width           =   1905
      End
      Begin VB.ComboBox cmbOrientation 
         Height          =   315
         ItemData        =   "frmReport.frx":00B8
         Left            =   405
         List            =   "frmReport.frx":00C2
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Seleccionar la orientación"
         Top             =   2835
         Width           =   1905
      End
      Begin VB.ComboBox cmbPrinters 
         Height          =   315
         Left            =   405
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Seleccionar la impresora"
         Top             =   810
         Width           =   2220
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Imprime"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   315
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Exporta"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   20
         Top             =   3375
         Width           =   1005
      End
      Begin VB.TextBox txtPathExportFile 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   330
         Left            =   75
         TabIndex        =   22
         ToolTipText     =   "Elegir un nombre para el archivo exportado"
         Top             =   4140
         Width           =   2310
      End
      Begin VB.ComboBox cmbTypeFile 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmReport.frx":00DC
         Left            =   75
         List            =   "frmReport.frx":00DE
         TabIndex        =   21
         ToolTipText     =   "Seleccionar tipo de formato de exportaciòn"
         Top             =   3690
         Width           =   2625
      End
      Begin VB.CommandButton cmdFindFile 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   2385
         TabIndex        =   31
         ToolTipText     =   "Elegir un nombre para el archivo exportado"
         Top             =   4140
         Width           =   330
      End
      Begin VB.Label Label5 
         Caption         =   "Ubicacion:"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   36
         Top             =   1440
         Width           =   2265
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   1215
         Width           =   2310
      End
      Begin VB.Label Label4 
         Caption         =   "Orientación:"
         Height          =   240
         Index           =   2
         Left            =   405
         TabIndex        =   34
         Top             =   2610
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "Papel:"
         Height          =   240
         Index           =   1
         Left            =   405
         TabIndex        =   33
         Top             =   2115
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "Impresora:"
         Height          =   240
         Index           =   0
         Left            =   405
         TabIndex        =   32
         Top             =   585
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   10530
      TabIndex        =   1
      Top             =   8130
      Width           =   960
   End
   Begin VB.Frame Frame2 
      Caption         =   " Copias "
      Height          =   1140
      Left            =   9225
      TabIndex        =   28
      Top             =   1620
      Width           =   2670
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   2160
         TabIndex        =   38
         Top             =   405
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCopias 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1620
         TabIndex        =   15
         Text            =   "1"
         ToolTipText     =   "Número de copias"
         Top             =   405
         Width           =   555
      End
      Begin VB.Image Img2 
         Height          =   735
         Left            =   180
         Stretch         =   -1  'True
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.TextBox txtCurrent 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1755
      TabIndex        =   5
      Top             =   180
      Width           =   690
   End
   Begin VB.TextBox txtBuscar 
      Height          =   330
      Left            =   4695
      TabIndex        =   8
      ToolTipText     =   "Texto para búsqueda"
      Top             =   180
      Width           =   1590
   End
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   5
      Left            =   6300
      Picture         =   "frmReport.frx":00E0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Buscar"
      Top             =   180
      Width           =   375
   End
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   90
      Picture         =   "frmReport.frx":022A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Elimina vista"
      Top             =   180
      Width           =   375
   End
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   4
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ultima página"
      Top             =   180
      Width           =   375
   End
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   3
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Página siguiente"
      Top             =   180
      Width           =   330
   End
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Primera página"
      Top             =   180
      Width           =   375
   End
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   2
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Página precedente"
      Top             =   180
      Width           =   330
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   9450
      TabIndex        =   0
      Top             =   8130
      Width           =   960
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   10800
      Top             =   405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   " Intervalo de páginas "
      Height          =   1500
      Left            =   9180
      TabIndex        =   24
      Top             =   45
      Width           =   2760
      Begin VB.TextBox txt2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1755
         TabIndex        =   14
         ToolTipText     =   "Página final"
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txt2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   630
         TabIndex        =   13
         ToolTipText     =   "Página inicial"
         Top             =   1080
         Width           =   600
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Páginas"
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   720
         Width           =   1410
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Todo"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   405
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "De                 hasta"
         Height          =   285
         Left            =   315
         TabIndex        =   25
         Top             =   1080
         Width           =   1410
      End
   End
   Begin VB.Label lbl1 
      Caption         =   "de 0+"
      Height          =   285
      Left            =   2475
      TabIndex        =   29
      Top             =   225
      Width           =   645
   End
   Begin VB.Image img1 
      Height          =   330
      Left            =   495
      Stretch         =   -1  'True
      Tag             =   "TreeDisabled"
      ToolTipText     =   "Visualiza/Esconde area de grupo"
      Top             =   180
      Width           =   345
   End
   Begin VB.Label Label3 
      Caption         =   "Buscar:"
      Height          =   240
      Left            =   4005
      TabIndex        =   27
      Top             =   225
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Zoom:"
      Height          =   240
      Left            =   6885
      TabIndex        =   26
      Top             =   225
      Width           =   510
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iMaxPage        As Integer            'numero máximo de páginas
Private bFormIsLoaded   As Boolean            'indicador de carga completada

Private mvarReport      As CRAXDRT.Report

Private Sub chkComentarios_Click()
   
   If chkComentarios = vbChecked Then
      mvarReport.txtComments.Suppress = False
   Else
      mvarReport.txtComments.Suppress = True
   End If
   
   If CRViewer1.IsBusy Then Exit Sub
   
   CRViewer1.Refresh
   
End Sub

Private Sub cmbOrientation_Click()

   If Not bFormIsLoaded Then Exit Sub
  
   Select Case cmbOrientation.Text
      Case "Horizontal"
        mvarReport.PaperOrientation = crLandscape
      Case "Vertical"
        mvarReport.PaperOrientation = crPortrait
   End Select
   
   If Not CRViewer1.IsBusy Then
      CRViewer1.Refresh
   End If
   
End Sub

Private Sub cmbPaper_Click()
  
  If Not bFormIsLoaded Then Exit Sub
  
  Select Case cmbPaper.Text
    Case "A4"
      mvarReport.PaperSize = crPaperA4
    Case "Legal"
      mvarReport.PaperSize = crPaperLegal
    Case "Letter"
      mvarReport.PaperSize = crPaperLetter
    Case "Executive"
      mvarReport.PaperSize = crPaperExecutive
  End Select
  
  If Not CRViewer1.IsBusy Then
      CRViewer1.Refresh
  End If

End Sub

Private Sub cmbPrinters_Click()
Dim driverName As String, printerName As String, portName As String
Dim strCurrentPageName As String
Dim ix As Integer
Dim strPrinterReg As String

  If Not bFormIsLoaded Then Exit Sub

  strPrinterReg = Replace(cmbPrinters.Text, "\", ",")
  'selecciono la impresora para la impresion
  printerName = GetRegistryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Print\Printers\" & strPrinterReg, "Name")
  portName = GetRegistryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Print\Printers\" & strPrinterReg, "Port")
  driverName = GetRegistryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Print\Printers\" & strPrinterReg, "Printer Driver")

  
  mvarReport.SelectPrinter driverName, printerName, portName
  
  CRViewer1.Refresh
     
  strCurrentPageName = PageName(mvarReport.PaperSize)
  For ix = 0 To cmbPaper.ListCount - 1
     If cmbPaper.List(ix) = strCurrentPageName Then Exit For
  Next
  cmbPaper.ListIndex = ix
   
  cmbOrientation.ListIndex = IIf(mvarReport.PaperOrientation = crLandscape, 0, 1)
  
  
  'actualizo labels
  Label5(0).Caption = "Tipo: " & driverName
  Label5(1).Caption = "Ubicación: " & portName
  
End Sub

Private Sub cmbPrinters_LostFocus()
Dim printerName As String
Dim strPrinterReg As String

  strPrinterReg = Replace(cmbPrinters.Text, "\", ",")
  
  printerName = GetRegistryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Print\Printers\" & strPrinterReg, "Name")
  If Len(printerName) = 0 Then
    MsgBox "La impresora seleccionada no esta definida en el sistema", vbExclamation
  End If

End Sub

Private Sub cmbTypeFile_Click()

  txtPathExportFile = NullString
  
End Sub

Private Sub cmbZoom_Click()

  If Right(cmbZoom.Text, 1) = "%" Then
    CRViewer1.Zoom Val(cmbZoom.Text)
  End If

  If cmbZoom.Text = "Ancho Página" Then
    CRViewer1.Zoom 1
  End If
  If cmbZoom.Text = "Página Entera" Then
    CRViewer1.Zoom 2
  End If
  SendKeys "{HOME}"

End Sub

Private Sub cmbZoom_LostFocus()

  If (cmbZoom.Text = "Ancho página") Or (cmbZoom.Text = "Página entera") Then Exit Sub

  Select Case Val(cmbZoom.Text)
    Case 1 To 25
      cmbZoom.Text = cmbZoom.Text & IIf(Right(cmbZoom.Text, 1) = "%", NullString, "%")
      MsgBox "El número debe estar comprendido entre 25 y 200", vbOKOnly
      cmbZoom.SetFocus
    Case 25 To 200
      cmbZoom.Text = cmbZoom.Text & IIf(Right(cmbZoom.Text, 1) = "%", NullString, "%")
      CRViewer1.Zoom Val(cmbZoom.Text)
    Case Else
      MsgBox "El número debe estar comprendido entre 25 y 200", vbOKOnly
      cmbZoom.SetFocus
  End Select

End Sub

Private Sub cmd1_Click(Index As Integer)
  DoEvents
  Select Case Index
    Case 0
      CRViewer1_CloseButtonClicked True
    Case 1
      If CRViewer1.IsBusy Then Exit Sub
      CRViewer1.ShowFirstPage
      DoEvents
      txtCurrent.Text = 1
      If iMaxPage = 1 Then
        lbl1.Caption = " de " & "1+"
      End If
    Case 2
      If CRViewer1.IsBusy Then Exit Sub
      CRViewer1.ShowPreviousPage
      DoEvents
      txtCurrent.Text = CRViewer1.GetCurrentPageNumber
    Case 3
      If CRViewer1.IsBusy Then Exit Sub
      CRViewer1.ShowNextPage
      DoEvents
      If CRViewer1.GetCurrentPageNumber > iMaxPage Then
        iMaxPage = CRViewer1.GetCurrentPageNumber
        lbl1.Caption = " de " & iMaxPage & "+"
      ElseIf CRViewer1.GetCurrentPageNumber = iMaxPage Then
        lbl1.Caption = " de " & iMaxPage
      End If
      txtCurrent.Text = CRViewer1.GetCurrentPageNumber
      
    Case 4
      If CRViewer1.IsBusy Then Exit Sub
      CRViewer1.ShowLastPage
      DoEvents
      iMaxPage = CRViewer1.GetCurrentPageNumber
      
      If CRViewer1.GetCurrentPageNumber > iMaxPage Then
        iMaxPage = CRViewer1.GetCurrentPageNumber
        lbl1.Caption = " de " & iMaxPage & "+"
      ElseIf CRViewer1.GetCurrentPageNumber = iMaxPage Then
        lbl1.Caption = " de " & iMaxPage
      End If
      txtCurrent.Text = CRViewer1.GetCurrentPageNumber
  
    
    Case 5
      If CRViewer1.IsBusy Then Exit Sub
      If Len(txtBuscar.Text) > 0 Then
        CRViewer1.SearchForText txtBuscar.Text
      Else
        MsgBox "Es necesario una expresión de bùsqueda", vbOKOnly
        txtBuscar.SetFocus
      End If
  End Select
  
      
End Sub
Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  
  'controlo intervalo de impresion
  If Opt1(1) Then
    If Len(txt2(0)) = 0 And Len(txt2(1)) = 0 Then
      MsgBox "No fue indicado el intervalo de impresión", vbOKOnly
      txt2(0).SetFocus
      Exit Sub
    End If
    If Len(txt2(1)) > 0 Then
      If Val(txt2(1)) < Val(txt2(0)) Then
        MsgBox "No es un intervalo de impresión válido", vbOKOnly
        txt2(0).SetFocus
        Exit Sub
      End If
    End If
  End If

  If opt2(0) Then
    '
    '    imprime en impresora definida por el metodo SelectPrinter
    '
    If Opt1(0) Then
      'imprime todo
      mvarReport.PrintOut False, CInt(txtCopias.Text)
    Else
      'imprime un rango de paginas
      mvarReport.PrintOut False, CInt(txtCopias.Text), True, CInt(txt2(0).Text), CInt(txt2(1).Text)
    End If
  Else
    '
    '   exporta
    '
    If Len(cmbTypeFile.Text) = 0 Then
      MsgBox "No fue especificado un formato para la exportación", vbOKOnly
      cmbTypeFile.SetFocus
      Exit Sub
    End If
    
    ' controlo el nombre del archivo
    If Len(txtPathExportFile.Text) = 0 Then
      MsgBox "No fue especificado el nombre del archivo", vbOKOnly
      txtPathExportFile.SetFocus
      Exit Sub
    End If
  
    mvarReport.ExportOptions.DestinationType = crEDTDiskFile
    mvarReport.ExportOptions.DiskFileName = txtPathExportFile.Text
    mvarReport.Export False
  End If
  
End Sub
Private Sub CRViewer1_CloseButtonClicked(UseDefault As Boolean)
  If CRViewer1.ActiveViewIndex > 1 Then
    CRViewer1.CloseView CRViewer1.ActiveViewIndex
  End If
End Sub

Private Sub CRViewer1_GroupTreeButtonClicked(ByVal Visible As Boolean)
  Select Case Visible
    Case True
      CRViewer1.DisplayGroupTree = False
    Case False
      CRViewer1.DisplayGroupTree = True
  End Select
End Sub


Private Sub Form_Load()
Dim ix                  As Integer
Dim nPos                As Integer
Dim strCurrentPrinter   As String
Dim strCurrentPageName  As String
Dim vRoot               As Variant
Dim vKey                As Variant
Dim strKey              As String
Dim strParametro        As String
Dim prt                 As Printer

   CRViewer1.ShowGroupTree = False
   bFormIsLoaded = False
   
   'obtengo lista de impresoras
   If Printers.Count = 0 Then
      MsgBox "No hay impresoras definidas"
   Else
      For Each prt In Printers
         cmbPrinters.AddItem prt.DeviceName
      Next prt
   End If
   
   With cmbTypeFile
      .AddItem "Texto (delimitado por caracter) *.txt"
      .AddItem "Hoja de cálculo de Microsoft Excel 2.1 (*.xls)"
      .AddItem "Hoja de cálculo de Microsoft Excel 3.0 (*.xls)"
      .AddItem "Hoja de cálculo de Microsoft Excel 4.0 (*.xls)"
      .AddItem "Hoja de cálculo de Microsoft Excel 5.0 (*.xls)"
      .AddItem "Ritch text formt (*.rtf)"
      .AddItem "Texto standard (*.txt)"
      .AddItem "Texto (delimitado por tabulacion) (*.txt)"
      .AddItem "Microsoft WinWord (*.doc)"
      .AddItem "WK1"
      .AddItem "WK3"
      .AddItem "WKS"
      .AddItem "Texto delimitado por coma (*.csv)"
      .AddItem "HTML 32 standard"
   End With
   
   img1.Picture = LoadPicture(Icons & "TreeDisabled.ico")
   Img2.Picture = LoadPicture(BitMaps & "Copias.bmp")
   cmd1(0).Picture = LoadPicture(Icons & "Cerrar.ico")
   cmd1(1).Picture = LoadPicture(Icons & "First.ico")
   cmd1(2).Picture = LoadPicture(Icons & "Prev.ico")
   cmd1(3).Picture = LoadPicture(Icons & "Next.ico")
   cmd1(4).Picture = LoadPicture(Icons & "Last.ico")
   cmd1(5).Picture = LoadPicture(Icons & "Find.ico")
    
   'obtengo impresora definida como default (registro de parametros)
   Select Case Left(WinVersion, 10)
     Case Is = "Windows NT"
       vRoot = GetKeyValuePI("Windows NT\Default Printer\Root")
       vKey = GetKeyValuePI("Windows NT\Default Printer\Clave")
     Case Is = "Windows 95"
       vRoot = GetKeyValuePI("Windows 95\Default Printer\Root")
       vKey = GetKeyValuePI("Windows 95\Default Printer\Clave")
     Case Is = "Windows 98"
       vRoot = GetKeyValuePI("Windows 98\Default Printer\Root")
       vKey = GetKeyValuePI("Windows 98\Default Printer\Clave")
   End Select
   nPos = InStr(vKey, ";")
   If nPos = 0 Then
       MsgBox "Error en la codificación de la clave de acceso al registro de windows" _
       & " para la localización de la impresora predefinida. Separar con ';' la clave" _
       & " y el parámetro que contiene dicho valor", vbInformation
    Else
       strKey = Trim(Left(vKey, nPos - 1))
       strParametro = Trim(Mid(vKey, nPos + 1))
   End If
  
   'recien aqui leo el registro de windows
   Select Case vRoot
      Case "HKEY_CLASSES_ROOT"
         strCurrentPrinter = GetRegistryValue(HKEY_CLASSES_ROOT, strKey, strParametro)
      Case "HKEY_CURRENT_USER"
         strCurrentPrinter = GetRegistryValue(HKEY_CURRENT_USER, strKey, strParametro)
      Case "HKEY_LOCAL_MACHINE"
         strCurrentPrinter = GetRegistryValue(HKEY_LOCAL_MACHINE, strKey, strParametro)
      Case "HKEY_USERS"
         strCurrentPrinter = GetRegistryValue(HKEY_USERS, strKey, strParametro)
   End Select
   nPos = InStr(strCurrentPrinter, ",")
   If nPos > 0 Then
      strCurrentPrinter = Trim(Left(strCurrentPrinter, nPos - 1))
   End If
     

   'visualizo informacion predefinida
   cmbZoom.Text = SystemOptions.iZoom & "%"
   For ix = 0 To cmbPrinters.ListCount - 1
      If cmbPrinters.List(ix) = strCurrentPrinter Then Exit For
   Next
   cmbPrinters.ListIndex = ix
   
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMaximized Then
      CRViewer1.Top = 800
      CRViewer1.Left = 0
      CRViewer1.Height = ScaleHeight - 900
      CRViewer1.Width = ScaleWidth - 2900
   End If
   
End Sub
Private Function PageName(iPaperSize As Integer) As String

  ' Convierte código del tamaño del papel en texto
  
  Select Case iPaperSize
    Case crPaperA4
      PageName = "A4"
    Case crPaperExecutive
      PageName = "Executive"
    Case crPaperLegal
      PageName = "Legal"
    Case crPaperLetter
      PageName = "Letter"
    Case crDefaultPaperSize
      PageName = "Predefinido"
  End Select

End Function
Private Sub Img1_Click()

  If img1.Tag = "TreeDisabled" Then
    img1.Picture = LoadPicture(Icons & "TreeEnabled.ico")
    img1.Tag = "TreeEnabled"
  Else
    img1.Picture = LoadPicture(Icons & "TreeDisabled.ico")
    img1.Tag = "TreeDisabled"
  End If
  CRViewer1_GroupTreeButtonClicked CRViewer1.DisplayGroupTree
End Sub


Private Sub Opt1_Click(Index As Integer)
  
  'abilito/desabilito controles del rango de impresion
  Select Case Index
    Case 0
      txt2(0).Enabled = False
      txt2(1).Enabled = False
      txt2(0).BackColor = DisabledColor
      txt2(1).BackColor = DisabledColor
    Case 1
      txt2(0).Enabled = True
      txt2(1).Enabled = True
      txt2(0).BackColor = EnabledColor
      txt2(1).BackColor = EnabledColor
  End Select

End Sub


Private Sub opt2_Click(Index As Integer)

  'abilito/desabilito controles de impresion/exportacion
  Select Case Index
    Case 0
    'desabilito controles export
      txtPathExportFile.Enabled = False
      cmbTypeFile.Enabled = False
      cmdFindFile.Enabled = False
      txtPathExportFile.BackColor = DisabledColor
      cmbTypeFile.BackColor = DisabledColor
      'abitlito controles impresion
      cmbPrinters.Enabled = True
      cmbPaper.Enabled = True
      cmbOrientation.Enabled = True
      cmbPrinters.BackColor = EnabledColor
      cmbPaper.BackColor = EnabledColor
      cmbOrientation.BackColor = EnabledColor
    Case 1
            
      'desabilito controles impresion
      cmbPrinters.Enabled = False
      cmbPaper.Enabled = False
      cmbOrientation.Enabled = False
      cmbPrinters.BackColor = DisabledColor
      cmbPaper.BackColor = DisabledColor
      cmbOrientation.BackColor = DisabledColor
      'abitlito controles export
      txtPathExportFile.Enabled = True
      cmbTypeFile.Enabled = True
      cmdFindFile.Enabled = True
      txtPathExportFile.BackColor = EnabledColor
      cmbTypeFile.BackColor = EnabledColor
      cmdFindFile.BackColor = EnabledColor
  End Select

End Sub

Private Sub txtCopias_LostFocus()
  If Len(Trim(txtCopias.Text)) = 0 Then
    txtCopias.SetFocus
    SendKeys "^Z"
    Exit Sub
  End If
  
  If Not IsNumeric(Trim(txtCopias.Text)) Then
    txtCopias.SetFocus
    SendKeys "^Z"
    Exit Sub
  End If

End Sub

Private Sub txtCurrent_LostFocus()
  
  If Len(Trim(txtCurrent.Text)) = 0 Then
    txtCurrent.SetFocus
    SendKeys "^Z"
    Exit Sub
  End If
  
  If Not IsNumeric(Trim(txtCurrent.Text)) Then
    txtCurrent.SetFocus
    SendKeys "^Z"
    Exit Sub
  End If
  
  CRViewer1.ShowNthPage (CInt(txtCurrent.Text))
  DoEvents
  If CRViewer1.GetCurrentPageNumber < Val(txtCurrent.Text) Then
    iMaxPage = CRViewer1.GetCurrentPageNumber
    lbl1.Caption = " de " & iMaxPage
  ElseIf CRViewer1.GetCurrentPageNumber > iMaxPage Then
    iMaxPage = CRViewer1.GetCurrentPageNumber
    lbl1.Caption = " de " & iMaxPage & "+"
  End If
  txtCurrent.Text = CRViewer1.GetCurrentPageNumber

End Sub

Private Sub cmdFindFile_Click()
Dim nPos    As Integer
Dim strChar As String

  'solicito nombre per archivo exportado
  txtPathExportFile.Text = Dialog("ShowSave", NullString, "*.*", 2500, 2500)
  If Len(txtPathExportFile.Text) = 0 Then
    txtPathExportFile.SetFocus
    Exit Sub
  End If
  
  nPos = InStr(txtPathExportFile.Text, ".")
  If nPos > 0 Then
    txtPathExportFile = Left(txtPathExportFile, nPos - 1)
  End If
  
  'agrego la extension del file
  Select Case cmbTypeFile
    Case "Texto (delimitado por caracter) *.txt"
      txtPathExportFile = txtPathExportFile & ".txt"
      mvarReport.ExportOptions.FormatType = crEFTCharSeparatedValues
      mvarReport.ExportOptions.CharFieldDelimiter = "@"
      strChar = InputBox("Ingrese el carácter para la separación de campos")
      If Len(strChar) > 0 Then
        mvarReport.ExportOptions.CharFieldDelimiter = Left(strChar, 1)
      End If
    Case "Hoja de cálculo de Microsoft Excel 2.1 (*.xls)"
      txtPathExportFile = txtPathExportFile & ".xls"
      mvarReport.ExportOptions.FormatType = crEFTExcel21
    Case "Hoja de cálculo de Microsoft Excel 3.0 (*.xls)"
      txtPathExportFile = txtPathExportFile & ".xls"
      mvarReport.ExportOptions.FormatType = crEFTExcel30
    Case "Hoja de cálculo de Microsoft Excel 4.0 (*.xls)"
      txtPathExportFile = txtPathExportFile & ".xls"
      mvarReport.ExportOptions.FormatType = crEFTExcel40
    Case "Hoja de cálculo de Microsoft Excel 5.0 (*.xls)"
      txtPathExportFile = txtPathExportFile & ".xls"
      mvarReport.ExportOptions.FormatType = crEFTExcel50
    Case "Ritch text formt (*.rtf)"
      txtPathExportFile = txtPathExportFile & ".rtf"
      mvarReport.ExportOptions.FormatType = crEFTRichText
    Case "Texto standard (*.txt)"
      txtPathExportFile = txtPathExportFile & ".txt"
      mvarReport.ExportOptions.FormatType = crEFTText
      mvarReport.ExportOptions.CharFieldDelimiter = NullString
    Case "Texto (delimitado por tabulacion) (*.txt)"
      txtPathExportFile = txtPathExportFile & ".txt"
      mvarReport.ExportOptions.FormatType = crEFTTabSeparatedText
    Case "Microsoft WinWord (*.doc)"
      txtPathExportFile = txtPathExportFile & ".doc"
      mvarReport.ExportOptions.FormatType = crEFTWordForWindows
    Case "WK1"
      txtPathExportFile = txtPathExportFile & ".wk1"
      mvarReport.ExportOptions.FormatType = crEFTLotus123WK1
    Case "WK3"
      txtPathExportFile = txtPathExportFile & ".wk3"
      mvarReport.ExportOptions.FormatType = crEFTLotus123WK3
    Case "WKS"
      txtPathExportFile = txtPathExportFile & ".wks"
      mvarReport.ExportOptions.FormatType = crEFTLotus123WKS
    Case "Texto delimitado por coma (*.csv)"
      txtPathExportFile = txtPathExportFile & ".csv"
      mvarReport.ExportOptions.FormatType = crEFTCommaSeparatedValues
    Case "HTML 32 standard"
      txtPathExportFile = txtPathExportFile & ".htm"
      mvarReport.ExportOptions.HTMLFileName = txtPathExportFile.Text
      mvarReport.ExportOptions.FormatType = crEFTHTML32Standard
  End Select
  
End Sub
Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKey txtBuscar.hWnd, KeyCode
End Sub
Private Sub txtCopias_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKey txtCopias.hWnd, KeyCode
End Sub
Private Sub txtCurrent_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKey txtCurrent.hWnd, KeyCode
End Sub
Private Sub txtPathExportFile_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKey txtPathExportFile.hWnd, KeyCode
End Sub
Private Sub txt2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKey txt2(Index).hWnd, KeyCode
End Sub

Private Function GetForm(ByVal lnghWnd As Long) As Form
Dim ix As Integer

   For ix = 0 To Forms.Count - 1
   
      If Forms(ix).hWnd = lnghWnd Then
         Set GetForm = Forms(ix)
         Exit For
      End If
   Next ix
   
End Function

Public Property Let Report(ByVal vData As Variant)
    mvarReport = vData
End Property

Public Property Set Report(ByVal vData As CRAXDRT.Report)
Dim strCurrentPageName  As String
Dim ix                  As Integer

   Set mvarReport = vData

   CRViewer1.ReportSource = vData
   
End Property

Public Property Get Report() As CRAXDRT.Report
   Set Report = mvarReport
End Property
Public Property Set Parameters(ByVal vData As Collection)
Dim ix As Integer
    
'    If CRViewer1.ReportSource Then Exit Property
    
   For ix = 1 To vData.Count
      Select Case mvarReport.ParameterFields(ix).ValueType
         Case crDateField, crDateTimeField, crTimeField
            mvarReport.ParameterFields(ix).SetCurrentValue CDate(vData(ix))
         Case Else
            mvarReport.ParameterFields(ix).SetCurrentValue vData(ix)
      End Select
   Next
      
End Property
Public Property Get Parameters() As Collection
    Set Parameters = mvarReport.ParameterFields
End Property

Public Sub ShowReport(ByVal bPreview As Boolean)
Dim strCurrentPageName  As String
Dim ix                  As Integer

On Error Resume Next

   If bPreview Then
   
      CRViewer1.ViewReport
      DoEvents
      
      CRViewer1.Zoom Val(cmbZoom.Text)
   
      chkComentarios.Value = IIf(Not mvarReport.txtComments.Suppress, vbChecked, vbUnchecked)
      
      strCurrentPageName = PageName(mvarReport.PaperSize)
      For ix = 0 To cmbPaper.ListCount - 1
         If cmbPaper.List(ix) = strCurrentPageName Then Exit For
      Next
      cmbPaper.ListIndex = ix
      
      cmbOrientation.ListIndex = IIf(mvarReport.PaperOrientation = crPortrait, 1, 0)
   
      ' me desplazo a la pagina 1
      iMaxPage = 1
      
      cmd1_Click 1
   
      bFormIsLoaded = True
      
      Me.Show
      
   Else
   
      mvarReport.PrintOut False
      
   End If

End Sub


Private Sub UpDown1_DownClick()
  If txtCopias.Text > 1 Then txtCopias.Text = Val(txtCopias.Text) - 1
End Sub

Private Sub UpDown1_UpClick()
  txtCopias.Text = Val(txtCopias.Text) + 1
End Sub
