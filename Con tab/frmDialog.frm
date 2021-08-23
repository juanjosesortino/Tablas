VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PowerMask1 
      Height          =   285
      Index           =   0
      Left            =   4395
      ScaleHeight     =   225
      ScaleWidth      =   1020
      TabIndex        =   14
      Top             =   3060
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar Opciones"
      Height          =   375
      Left            =   4260
      TabIndex        =   13
      ToolTipText     =   "Guardar Opciones"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame freButtons 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   4500
      TabIndex        =   7
      Top             =   90
      Width           =   1125
      Begin VB.CommandButton cmdFiltros 
         Caption         =   "&Otros Filtros"
         Height          =   375
         Left            =   45
         TabIndex        =   11
         ToolTipText     =   "Otros Filtros"
         Top             =   1755
         Width           =   1035
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Vista Previa"
         Height          =   375
         Left            =   60
         TabIndex        =   10
         ToolTipText     =   "Vista Previa "
         Top             =   1320
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   60
         TabIndex        =   9
         ToolTipText     =   "Cancela"
         Top             =   450
         Width           =   1035
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   60
         TabIndex        =   8
         ToolTipText     =   "Acepta"
         Top             =   30
         Width           =   1035
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   930
      TabIndex        =   3
      Top             =   3045
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   2610
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Index           =   0
      Left            =   270
      TabIndex        =   1
      Top             =   2190
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   12
      Top             =   3465
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmDialog.frx":0000
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00ECFAFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1710
      TabIndex        =   6
      Top             =   3030
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      DrawMode        =   4  'Mask Not Pen
      Index           =   0
      Visible         =   0   'False
      X1              =   990
      X2              =   4260
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   3090
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Menu mnuContextMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuContextItem 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const FRAME_LEFT = 100                          ' left del frame
Private Const FRAME_DISTANCE = 100                      ' distancia entre frames
Private Const OPTION_LEFT = 200                         ' valor del Left para todos los controles check, option dentro de un frame
Private Const OPTION_TOP = 200                          ' el top del primer option / check en el frame
Private Const OPTION_DISTANCE = 50                      ' distancia entre options/checks
Private Const BUTTON_DISTANCE = 100                     ' distancia entre los botones Acepta y Cancela

'  posicion de los datos en coleccion collSearch
Private Const TABLE_FIELD = 0
Private Const CONTROL_LIST = 1
Private Const FIELD_LIST = 2
Private Const LABEL_CONTROL = 3
Private Const LABEL_FIELDNAME = 4

Private WithEvents frmHook As AlgStdFunc.MsgHook
Attribute frmHook.VB_VarHelpID = -1
Private WithEvents tmr1    As AlgStdFunc.clsTimer
Attribute tmr1.VB_VarHelpID = -1

Private Const WM_DESTROY = &H2                          'mensaje que todas las windows reciben antes de ser cerradas
Private Const WM_SHOWWINDOW = &H18

Public mvarDialogTitle                 As String            ' Titulo del Dialog
Private WithEvents objControls         As clsControls       'instancia de la clase clsControls
Attribute objControls.VB_VarHelpID = -1

Public Enum EnumButtonPressedDialog
   Cancel = 0
   Accept = 1
   Preview = 3
End Enum
   
Public ButtonPressed                   As EnumButtonPressedDialog

Public Enum EnumEnabledDisabledControls
   Disabled = 0
   Enabled = 1
End Enum

Private nextBtnTop               As Single            ' es el Top para proximo control
Private maxWidth                 As Single            ' es el ancho del frame
Private TabIndexCounter          As Integer           ' es el contador del TabIndex

Private Values                   As New Collection    ' coleccion devuelta al llamador con todos los valores de los controles
Private colFormat                As New Collection    ' coleccion de objetos textbox que debe ser formateados. Tienen un evento LostFocus para ejecutar
Private colKeyPress              As New Collection    ' coleccion de objetos que tienen un evento KeyPress para ejecutar
Private colSearch                As New Collection    ' coleccion de datos para QueryDB
Private colProperties            As New Collection    ' coleccion de propiedades
Private colDataType              As New Collection    ' coleccion de tipos para cada control que necesitan validación
Private colControlsDisabled      As New Collection    ' coleccion de tipos para cada control que necesitan validación
Private colControlsEnabled       As New Collection    ' coleccion de tipos para cada control que necesitan validación
Private colControlsCanSave       As New Collection    ' coleccion de controles que pueden salvar su valor de default

Private mvarfrmFiltros           As frmFilter         ' form frmFilter
Private mvarCaptionsFilter       As String            ' Lista de captions para el filtro avanzado
Private mvarFieldsFilter         As String            ' Lista de campos para el filtro avanzado
Private mvarShowPrintButton      As Boolean
Private mvarShowSaveButton       As Boolean           ' determina si será visible el boton para "Salvar Opciones"
Private mvarShowFilterButton     As Boolean
Private mvarEmpresa              As String            ' Empresa con la que estoy trabajando

Dim objTabla                     As New BOGeneral.clsTablas
Dim aKeys()                      As Variant
Dim aArray()                     As Variant
Dim vValue                       As Variant
Dim strKey                       As String

' definicion de los eventos de la clase
Public Event ValidateDialog(ByRef Response As String)

Private Const WM_CONTEXTMENU = &H7B
Private Const WM_KEYDOWN = &H100
Private Const WM_KILLFOCUS = &H8
Private Const WM_RBUTTONDOWN = &H204

Public Sub AddFrame(ByVal key As String, ByVal Caption As String, Optional ByVal showBorder As Boolean = True, Optional ByVal ShowSaveButton As Boolean, Optional ByVal KeyParent As String)
Dim frameIndex    As Integer
Dim thisFrame     As Frame
Dim prevFrame     As Frame
    
    ' agrego un nuevo frame
    
    
    ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
    ' devuelve el indice del proximo frame que se creara
    frameIndex = Frame1.UBound + 1
    
    ' cargo un nuevo frame
    Load Frame1(frameIndex)
    Set thisFrame = Frame1(frameIndex)
    Set prevFrame = Frame1(frameIndex - 1)
    ' si este no es el primer frame, lo muevo debajo del precedente
    If frameIndex > 1 Then
        If Len(KeyParent) = 0 Then
           thisFrame.Move FRAME_LEFT, prevFrame.Top + prevFrame.Height + FRAME_DISTANCE, prevFrame.Width, thisFrame.Height
        Else
           thisFrame.Move FRAME_LEFT, GetControl(KeyParent).Height, prevFrame.Width, thisFrame.Height
           thisFrame.Width = thisFrame.Width - 300
           
        End If
    End If
    
    
    ' setea las propiedades del frame
    thisFrame.Caption = Caption
    thisFrame.Tag = key
    ' valores posibles para el BorderStyle son 0 (none) o 1 (visible)
    thisFrame.BorderStyle = -(showBorder)
    
    ' seteo la posicion del proximo control
    nextBtnTop = OPTION_TOP
    
    thisFrame.Visible = True
    
    If Len(KeyParent) > 0 Then
      Set thisFrame.Container = GetControl(KeyParent)
    End If
    
   If ShowSaveButton Then
      mvarShowSaveButton = True
      Set cmdGuardar.Container = thisFrame
      cmdGuardar.Visible = True
   End If
    
End Sub

Public Sub AddOption(ByVal key As String, ByVal Caption As String, Optional ByVal Top As Long, Optional ByVal Left As Long, Optional Value As Boolean, Optional ByVal ControlListDisabled As String, Optional ByVal ControlListEnabled As String, Optional ByVal CanSaveValue As Boolean = True, Optional ByVal Width As Integer)
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim thisBtn       As OptionButton
Dim thisFrame     As Frame
Dim FrameContainer As Frame
    
   ' agreo un Option al grupo corriente
   
   ' este es el numero del corriente grupo de controles
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   
   ' agrago un nuovo option
   optionIndex = Option1.UBound + 1
   Load Option1(optionIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisBtn = Option1(optionIndex)
   thisBtn.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   ' pongo el control dentro del frame adecuado
   Set thisBtn.Container = thisFrame
   ' lo muevo a la posicion correcta
   If Top > 0 Then
      thisBtn.Move OPTION_LEFT, Top
   End If
   If Left > 0 Then
      thisBtn.Move Left, thisBtn.Top
   End If
   If Top = 0 And Left = 0 Then
      thisBtn.Move OPTION_LEFT, nextBtnTop
   End If
   
   If Width > 0 Then
      thisBtn.Width = Width
   End If
   
   ' calculo la posicion del proximo control
   nextBtnTop = thisBtn.Top + thisBtn.Height + OPTION_DISTANCE
   
   ' recalculo las dimensione del frame y el form
   thisFrame.Height = nextBtnTop
   If Not (thisFrame.Container Is Nothing) Then
      If thisFrame.Container.Name <> "frmDialog" Then
         Set FrameContainer = thisFrame.Container
         FrameContainer.Height = thisFrame.Top + nextBtnTop
      End If
   End If
   
   If Len(ControlListDisabled) > 0 Then
      colControlsDisabled.Add ControlListDisabled, key        'no necesito el item, solo la clave
   End If
   
   If Len(ControlListEnabled) > 0 Then
      colControlsEnabled.Add ControlListEnabled, key          'no necesito el item, solo la clave
   End If
   
   ' seteo la propiedades del option
   thisBtn.Tag = key
   thisBtn.Caption = Caption
   thisBtn.Value = Value
   
   ' lo hago visible
   thisBtn.Visible = True
   
   If CanSaveValue Then
      colControlsCanSave.Add thisBtn, key
   End If
   
End Sub

Public Sub AddCheck(ByVal key As String, ByVal Caption As String, Optional ByVal Top As Long, Optional ByVal Left As Long, Optional Value As Integer, Optional ByVal ControlListDisabled As String, Optional ByVal ControlListEnabled As String, Optional ByVal CanSaveValue As Boolean = True, Optional ByVal Width As Integer)
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim thisCheck     As CheckBox
Dim thisFrame     As Frame
Dim FrameContainer As Frame
    
    ' agrego un check al grupo corriente
    
    ' este es el numero del corriente grupo de controles
    frameIndex = Frame1.UBound
    Set thisFrame = Frame1(frameIndex)
   
    ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
    ' devuelve el indice del proximo frame que se creara
    optionIndex = Check1.UBound + 1
    Load Check1(optionIndex)
    
    ' me creo una referencia al control para simplificar el codigo
    Set thisCheck = Check1(optionIndex)
    thisCheck.TabIndex = TabIndexCounter
    TabIndexCounter = TabIndexCounter + 1
    
    ' pongo el control dentro del frame adecuado
    Set thisCheck.Container = thisFrame
    ' lo muevo a la posicion correcta
   If Top > 0 Then
      thisCheck.Move OPTION_LEFT, Top
   End If
   If Left > 0 Then
      thisCheck.Move Left, thisCheck.Top
   End If
   If Top = 0 And Left = 0 Then
      thisCheck.Move OPTION_LEFT, nextBtnTop
   End If
   
   If Width > 0 Then
      thisCheck.Width = Width
   End If
   
   ' calculo la posicion del proximo control
   nextBtnTop = thisCheck.Top + thisCheck.Height + OPTION_DISTANCE
    
   ' recalculo las dimensione del frame y el form
   thisFrame.Height = nextBtnTop
   If Not (thisFrame.Container Is Nothing) Then
      If thisFrame.Container.Name <> "frmDialog" Then
         Set FrameContainer = thisFrame.Container
         FrameContainer.Height = FrameContainer.Height + nextBtnTop
      End If
   End If
   
   If Len(ControlListDisabled) > 0 Then
      colControlsDisabled.Add ControlListDisabled, key        'no necesito el item, solo la clave
   End If
   
   If Len(ControlListEnabled) > 0 Then
      colControlsEnabled.Add ControlListEnabled, key          'no necesito el item, solo la clave
   End If
   
    
    ' seteo la propiedades del check
    thisCheck.Tag = key
    thisCheck.Caption = Caption
    
    thisCheck.Value = 1
    thisCheck.Value = Abs(Value)
    
    ' lo hago visible
    thisCheck.Visible = True

   If CanSaveValue Then
      colControlsCanSave.Add thisCheck, key
   End If
   
End Sub

Public Sub AddLabel(ByVal key As String, ByVal Caption As String, Optional ByVal LeftCaption As Integer, Optional ByVal TopCaption As Integer, Optional ByVal color As Long)
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim thisLabel     As Label
Dim thisFrame     As Frame
    
   ' agrego un label al grupo corriente
   
   If TopCaption = 0 Then
      TopCaption = nextBtnTop
   End If
   
   If LeftCaption = 0 Then
      LeftCaption = OPTION_LEFT
   End If
   
   
   ' este es el numero del corriente grupo de controles
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = Label1.UBound + 1
   Load Label1(optionIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisLabel = Label1(optionIndex)
   thisLabel.Width = TextWidth(Caption)
   
   'controlo si es necesario aumentar el ancho del frame y del form
   If maxWidth < LeftCaption + thisLabel.Width Then
      maxWidth = LeftCaption + thisLabel.Width
      
      thisFrame.Width = maxWidth + 100
      freButtons.Left = FRAME_LEFT + thisFrame.Width + FRAME_LEFT
      
      'calculo el ancho del form
      Width = freButtons.Left + freButtons.Width + 2 * FRAME_LEFT
       
      ' muevo los botones a la derecha del ultimo frame creado
      freButtons.Move freButtons.Left, freButtons.Top
   
   End If
   
   
   ' calculo la posicion del proximo control y modifico si es necesario el alto del frame y del form
   nextBtnTop = TopCaption + thisLabel.Height + 200
   
   ' recalculo las dimensione del frame y el form
   thisFrame.Visible = True
   thisFrame.Height = nextBtnTop
   
   ' recalculo la altrura del form
   If freButtons.Top + freButtons.Height < thisFrame.Top + thisFrame.Height Then
      Height = 2 * thisFrame.Top + thisFrame.Height + 500
   Else
      Height = 2 * freButtons.Top + freButtons.Height + 500
   End If
   
   ' pongo el control dentro del frame adecuado
   Set thisLabel.Container = thisFrame
   
   ' lo muevo a la posicion correcta
   thisLabel.Move LeftCaption, TopCaption
   
   ' calculo la posicion del proximo control
   nextBtnTop = nextBtnTop + thisLabel.Height + OPTION_DISTANCE
   
   ' seteo la propiedades del check
   thisLabel.Tag = key
   thisLabel.Caption = Caption
   If color > 0 Then
    thisLabel.ForeColor = color
   End If
   
   ' lo hago visible
   thisLabel.Visible = True

End Sub

Public Sub AddHLine(ByVal Top As Long, ByVal Left As Long, ByVal Lenght As Long)
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim thisLine      As Line
Dim thisFrame     As Frame
    
   ' agrego un check al grupo corriente
   
   ' este es el numero del corriente grupo de controles
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = Line1.UBound + 1
   Load Line1(optionIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisLine = Line1(optionIndex)
   thisLine.Y1 = Top
   thisLine.Y2 = Top
   thisLine.X1 = Left
   thisLine.X2 = Left + Lenght
   thisLine.BorderColor = &H0&
   
   ' pongo el control dentro del frame adecuado
   Set thisLine.Container = thisFrame
   
   ' lo hago visible
   thisLine.Visible = True
   
   'creo la segunda linea un pixel mas abajo
   
   optionIndex = Line1.UBound + 1
   Load Line1(optionIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisLine = Line1(optionIndex)
   
   thisLine.Y1 = Top - 15
   thisLine.Y2 = thisLine.Y1
   thisLine.X1 = Left
   thisLine.X2 = Left + Lenght
   thisLine.BorderColor = &H808080
   
   ' pongo el control dentro del frame adecuado
   Set thisLine.Container = thisFrame
   
   ' lo hago visible
   thisLine.Visible = True

End Sub

Public Sub AddVLine(ByVal Top As Long, ByVal Left As Long, ByVal Height As Long)
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim thisLine      As Line
Dim thisFrame     As Frame
    
   ' agrego un check al grupo corriente
   
   ' este es el numero del corriente grupo de controles
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = Line1.UBound + 1
   Load Line1(optionIndex)
   
   ' pongo el control dentro del frame adecuado
   Set thisLine.Container = thisFrame
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisLine = Line1(optionIndex)
   thisLine.Y1 = Top
   thisLine.Y2 = Top + Height
   thisLine.X1 = Left
   thisLine.X2 = Left
   thisLine.BorderColor = &H0&
   
   ' lo hago visible
   thisLine.Visible = True
   
   'creo la segunda linea un pixel mas abajo
   
   optionIndex = Line1.UBound + 1
   Load Line1(optionIndex)
   
   ' pongo el control dentro del frame adecuado
   Set thisLine.Container = thisFrame
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisLine = Line1(optionIndex)
   
   thisLine.Y1 = Top
   thisLine.Y2 = Top + Height
   thisLine.X1 = Left - 10
   thisLine.X2 = Left - 10
   thisLine.BorderColor = &H808080
   
   ' lo hago visible
   thisLine.Visible = True

End Sub

Public Sub AddTextLabel(ByVal key As String, ByVal Caption As String, ByVal LeftCaption As Integer, _
                        ByVal Top As Integer, ByVal LeftText As Integer, ByVal LeftDescription As Integer, _
                        ByVal WidthText As Integer, ByVal WidthDescription As Integer, ByVal strTableField As String, _
                        ByVal strBoundLabel As String, Optional ByVal FindControlList As String, _
                        Optional ByVal FindBoundFieldList As String, Optional ByVal CanSaveValue As Boolean)
                        
Dim optionIndex         As Integer
Dim frameIndex          As Integer
Dim thisText            As TextBox
Dim thisLabel           As Label
Dim thisDescription     As Label
Dim thisFrame           As Frame
Dim aTableProperties    As Variant

   'Este metodo se usa para los Bound TextBox por eso es que no es necesario. objControls
   'obtiene la mayor parte de la informacion desde el diccionario

   ' agrego un textbox al grupo actual
   
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = Text1.UBound + 1
   If Len(Caption) > 0 Then
     'creo un label
     Load Label1(Label1.UBound + 1)
     Set thisLabel = Label1(Label1.UBound)
     Set thisLabel.Container = thisFrame
     thisLabel.Caption = Caption & ":"
     thisLabel.Width = TextWidth(thisLabel.Caption)
     thisLabel.Move LeftCaption, Top + 40
     thisLabel.Tag = "Label" & key
     thisLabel.Visible = True
   End If
   ' lo muevo a la posicion correcta
   thisLabel.Move LeftCaption, Top
   
   ' me creo una referencia al control para simplificar el codigo
   Load Text1(optionIndex)
   Set thisText = Text1(optionIndex)
   Set thisText.Container = thisFrame
   thisText.Width = WidthText
   thisText.Text = NullString
   thisText.Tag = key
   thisText.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   ' lo muevo a la posicion correcta
   thisText.Move LeftText, Top
   
   ' me creo una referencia al control para simplificar el codigo
   Load lblDescripcion(lblDescripcion.UBound + 1)
   Set thisDescription = lblDescripcion(lblDescripcion.UBound)
   Set thisDescription.Container = thisFrame
   thisDescription.Width = WidthDescription
   thisDescription.Caption = ""
   ' lo muevo a la posicion correcta
   thisDescription.Move LeftDescription, Top
   
   ' lo hago visible
   thisLabel.Visible = True
   thisText.Visible = True
   thisDescription.Visible = True
   
   '-- Información necesaria para la actualización:
   '--    Entrada diccionario
   '--    Nombres de TextBox actualizado por QueryDB (lista separada por ";")
   '--    Nombres de campos de QueryDB que actualizan los TextBox (lista separada por ";")
   '--    Control Label Actualizado por QueryDB
   '--    Campo del QueryDB que actualiza el Label descriptivo
   '--    Opcional: Tabla a leer si es distinta a la de la entrada en el diccionario (ej. Transportistas)
   '--
   
   If maxWidth < LeftDescription + WidthDescription Then
     maxWidth = LeftDescription + WidthDescription
     thisFrame.Width = maxWidth + 100
     freButtons.Left = FRAME_LEFT + thisFrame.Width + FRAME_LEFT
     
     'calculo el ancho del form
     Width = freButtons.Left + freButtons.Width + 2 * FRAME_LEFT
      
     ' muevo los botones a la derecha del ultimo frame creado
     freButtons.Move freButtons.Left, freButtons.Top
   
   End If
   
   ' calculo la posicion del proximo control
   nextBtnTop = Top + thisText.Height + 200
   
   ' recalculo las dimensione del frame y el form
   thisFrame.Height = nextBtnTop
   
   ' recalculo la altrura del form
   If freButtons.Top + freButtons.Height < thisFrame.Top + thisFrame.Height Then
     Height = 2 * thisFrame.Top + thisFrame.Height + 500
   Else
     Height = 2 * freButtons.Top + freButtons.Height + 500
   End If
   
   If Len(strTableField) > 0 Then
   
      aTableProperties = GetFieldInformation(strTableField)
      
      'agrego datos para la búsqueda a la colección
      
      If Len(FindControlList) = 0 Then FindControlList = key
      
      colSearch.Add Array(strTableField, FindControlList, FindBoundFieldList, thisDescription, strBoundLabel), key
      
      
      colProperties.Add aTableProperties, key 'guardo la propiedades del campo o bien un
                                              'arreglo vacio si el control es UnBound
      
   End If
   
   ' seteo la propiedades del control
   thisText.Tag = key
      
   
   If CanSaveValue Then
      colControlsCanSave.Add thisText, key
   End If
   
   objControls.Add thisText, strTableField, , , FindControlList, FindBoundFieldList
   
End Sub

Public Sub AddText(ByVal key As String, ByVal Caption As String, ByVal LeftCaption As Integer, _
                   ByVal Top As Integer, ByVal LeftText As Integer, ByVal WidthText As Integer, _
                   Optional ByVal strTableField As String, Optional ByVal strDefaultValue As String, _
                   Optional ByVal CanSaveValue As Boolean, Optional ByVal Validar As ValidEnum, _
                   Optional ByVal Dimension As Integer, Optional Decimales As Integer)
                   
Dim optionIndex         As Integer
Dim frameIndex          As Integer
Dim thisText            As TextBox
Dim thisLabel           As Label
Dim thisFrame           As Frame
Dim strField            As String
Dim aTableProperties    As Variant

   ' agrego un textbox al grupo actual
   
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = Text1.UBound + 1
   If Len(Caption) > 0 Then
      'creo un label
      Load Label1(Label1.UBound + 1)
      Set thisLabel = Label1(Label1.UBound)
      Set thisLabel.Container = thisFrame
      thisLabel.Caption = Caption & ":"
      thisLabel.Width = TextWidth(thisLabel.Caption)
      thisLabel.Move LeftCaption, Top + 40
      thisLabel.Visible = True
      thisLabel.Tag = "Label" & key
      
   End If
   
   ' me creo una referencia al control para simplificar el codigo
   Load Text1(optionIndex)
   Set thisText = Text1(optionIndex)
   Set thisText.Container = thisFrame
   thisText.Width = WidthText
   thisText.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   If Len(strDefaultValue) = 0 Then
      thisText.Text = strDefaultValue
   End If
   
   SetModify thisText.hWnd, True
   thisText.Tag = key
   thisText.Visible = True
   ' lo muevo a la posicion correcta
   thisText.Move LeftText, Top
   
   
   If maxWidth < LeftText + WidthText Then
      maxWidth = LeftText + WidthText
      thisFrame.Width = maxWidth + 100
      freButtons.Left = FRAME_LEFT + thisFrame.Width + FRAME_LEFT
      
      'calculo el ancho del form
      Width = freButtons.Left + freButtons.Width + 2 * FRAME_LEFT
       
      ' muevo los botones a la derecha del ultimo frame creado
      freButtons.Move freButtons.Left, freButtons.Top
   End If
   
   ' calculo la posicion del proximo control
   nextBtnTop = Top + thisText.Height + 200
   
   ' recalculo las dimensione del frame y el form
   thisFrame.Height = nextBtnTop
   
   ' recalculo la altrura del form
   If freButtons.Top + freButtons.Height < thisFrame.Top + thisFrame.Height Then
      Height = 2 * thisFrame.Top + thisFrame.Height + 500
   Else
      Height = 2 * freButtons.Top + freButtons.Height + 500
   End If
   
   If Len(strTableField) > 0 Then
   
      aTableProperties = GetFieldInformation(strTableField)
      
      strField = FieldProperty(aTableProperties, strTableField, dsCampo)

      'agrego datos para la búsqueda a la colección
      colSearch.Add Array(strTableField, thisText.Tag, strField), key
      
      
      colProperties.Add aTableProperties, key 'guardo la propiedades del campo o bien un
                                              'arreglo vacio si el control es UnBound
      
   End If
   
   
   If Len(thisText.Text) > 0 Then
      Text1_LostFocus (thisText.Index)
   End If
   
   If CanSaveValue Then
      colControlsCanSave.Add thisText, key
   End If
   
   objControls.Add thisText, strTableField, , Validar, , , Dimension, Decimales
   
End Sub

Public Sub AddPowerMaskLabel(ByVal key As String, ByVal Caption As String, ByVal LeftCaption As Integer, _
                        ByVal Top As Integer, ByVal LeftText As Integer, ByVal LeftDescription As Integer, _
                        ByVal WidthText As Integer, ByVal WidthDescription As Integer, ByVal strTableField As String, _
                        ByVal strBoundLabel As String, Optional ByVal FindControlList As String, _
                        Optional ByVal FindBoundFieldList As String, Optional ByVal CanSaveValue As Boolean)
                        
Dim optionIndex         As Integer
Dim frameIndex          As Integer
Dim thisPowerMask       As PowerMask
Dim thisLabel           As Label
Dim thisDescription     As Label
Dim thisFrame           As Frame
Dim aTableProperties    As Variant

   'Este metodo se usa para los Bound PowerMask.
   'objControls obtiene la mayor parte de la informacion desde el diccionario

   ' agrego un PowerMask al grupo actual
   
   aTableProperties = GetFieldInformation(strTableField)
   
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = PowerMask1.UBound + 1
   If Len(Caption) > 0 Then
     'creo un label
     Load Label1(Label1.UBound + 1)
     Set thisLabel = Label1(Label1.UBound)
     Set thisLabel.Container = thisFrame
     thisLabel.Caption = Caption & ":"
     thisLabel.Width = TextWidth(thisLabel.Caption)
     thisLabel.Move LeftCaption, Top + 40
     thisLabel.Tag = "Label" & key
     thisLabel.Visible = True
   End If
   ' lo muevo a la posicion correcta
   thisLabel.Move LeftCaption, Top
   
   ' me creo una referencia al control para simplificar el codigo
   Load PowerMask1(optionIndex)
   Set thisPowerMask = PowerMask1(optionIndex)
   Set thisPowerMask.Container = thisFrame
   thisPowerMask.Mask = FieldProperty(aTableProperties, strTableField, dsCodigoEspecial)
   thisPowerMask.Width = WidthText
   thisPowerMask.Text = NullString
   thisPowerMask.Tag = key
   thisPowerMask.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   ' lo muevo a la posicion correcta
   thisPowerMask.Move LeftText, Top
   
   ' me creo una referencia al control para simplificar el codigo
   Load lblDescripcion(lblDescripcion.UBound + 1)
   Set thisDescription = lblDescripcion(lblDescripcion.UBound)
   Set thisDescription.Container = thisFrame
   thisDescription.Width = WidthDescription
   thisDescription.Caption = ""
   ' lo muevo a la posicion correcta
   thisDescription.Move LeftDescription, Top
   
   ' lo hago visible
   thisLabel.Visible = True
   thisPowerMask.Visible = True
   thisDescription.Visible = True
   
   '-- Información necesaria para la actualización:
   '--    Entrada diccionario
   '--    Nombres de TextBox actualizado por QueryDB (lista separada por ";")
   '--    Nombres de campos de QueryDB que actualizan los TextBox (lista separada por ";")
   '--    Control Label Actualizado por QueryDB
   '--    Campo del QueryDB que actualiza el Label descriptivo
   '--    Opcional: Tabla a leer si es distinta a la de la entrada en el diccionario (ej. Transportistas)
   '--
   
   If Len(FindControlList) = 0 Then FindControlList = key
   
   colSearch.Add Array(strTableField, FindControlList, FindBoundFieldList, thisDescription, strBoundLabel), key
   
   If maxWidth < LeftDescription + WidthDescription Then
     maxWidth = LeftDescription + WidthDescription
     thisFrame.Width = maxWidth + 100
     freButtons.Left = FRAME_LEFT + thisFrame.Width + FRAME_LEFT
     
     'calculo el ancho del form
     Width = freButtons.Left + freButtons.Width + 2 * FRAME_LEFT
      
     ' muevo los botones a la derecha del ultimo frame creado
     freButtons.Move freButtons.Left, freButtons.Top
   
   End If
   
   ' calculo la posicion del proximo control
   nextBtnTop = Top + thisPowerMask.Height + 200
   
   ' recalculo las dimensione del frame y el form
   thisFrame.Height = nextBtnTop
   
   ' recalculo la altrura del form
   If freButtons.Top + freButtons.Height < thisFrame.Top + thisFrame.Height Then
     Height = 2 * thisFrame.Top + thisFrame.Height + 500
   Else
     Height = 2 * freButtons.Top + freButtons.Height + 500
   End If
   
   
   ' seteo la propiedades del control
   thisPowerMask.Tag = key
      
   colProperties.Add aTableProperties, key 'guardo la propiedades del campo
   
   If CanSaveValue Then
      colControlsCanSave.Add thisPowerMask, key
   End If
   
   objControls.Add thisPowerMask, strTableField, , , FindControlList, FindBoundFieldList
   
End Sub

Public Sub AddPowerMask(ByVal key As String, ByVal Caption As String, ByVal LeftCaption As Integer, _
                   ByVal Top As Integer, ByVal LeftText As Integer, ByVal WidthText As Integer, _
                   Optional ByVal strTableField As String, Optional ByVal Mask As CodigosEspecialesEnum, _
                   Optional ByVal strDefaultValue As String, Optional ByVal CanSaveValue As Boolean)
                   
Dim optionIndex         As Integer
Dim frameIndex          As Integer
Dim thisPowerMask       As PowerMask
Dim thisLabel           As Label
Dim thisFrame           As Frame
Dim strField            As String
Dim aTableProperties    As Variant

   ' agrego un PowerMask al grupo actual
   
   If Len(strTableField) > 0 Then
   
      aTableProperties = GetFieldInformation(strTableField)
      
      strField = FieldProperty(aTableProperties, strTableField, dsCampo)

      'agrego datos para la búsqueda a la colección
      colSearch.Add Array(strTableField, thisPowerMask.Tag, strField), key
      
      
      colProperties.Add aTableProperties, key 'guardo la propiedades del campo o bien un
                                              'arreglo vacio si el control es UnBound
                                              
      Mask = FieldProperty(aTableProperties, strTableField, dsCodigoEspecial)
   
   End If
   
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = PowerMask1.UBound + 1
   If Len(Caption) > 0 Then
      'creo un label
      Load Label1(Label1.UBound + 1)
      Set thisLabel = Label1(Label1.UBound)
      Set thisLabel.Container = thisFrame
      thisLabel.Caption = Caption & ":"
      thisLabel.Width = TextWidth(thisLabel.Caption)
      thisLabel.Move LeftCaption, Top + 40
      thisLabel.Visible = True
      thisLabel.Tag = "Label" & key
      
   End If
   
   ' me creo una referencia al control para simplificar el codigo
   Load PowerMask1(optionIndex)
   Set thisPowerMask = PowerMask1(optionIndex)
   Set thisPowerMask.Container = thisFrame
   thisPowerMask.Mask = Mask
   thisPowerMask.Width = WidthText
   thisPowerMask.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   If Len(strDefaultValue) = 0 Then
      thisPowerMask.Text = strDefaultValue
   End If
   
   SetModify thisPowerMask.hWnd, True
   thisPowerMask.Tag = key
   thisPowerMask.Visible = True
   ' lo muevo a la posicion correcta
   thisPowerMask.Move LeftText, Top
   
   
   If maxWidth < LeftText + WidthText Then
      maxWidth = LeftText + WidthText
      thisFrame.Width = maxWidth + 100
      freButtons.Left = FRAME_LEFT + thisFrame.Width + FRAME_LEFT
      
      'calculo el ancho del form
      Width = freButtons.Left + freButtons.Width + 2 * FRAME_LEFT
       
      ' muevo los botones a la derecha del ultimo frame creado
      freButtons.Move freButtons.Left, freButtons.Top
   End If
   
   ' calculo la posicion del proximo control
   nextBtnTop = Top + thisPowerMask.Height + 200
   
   ' recalculo las dimensione del frame y el form
   thisFrame.Height = nextBtnTop
   
   ' recalculo la altrura del form
   If freButtons.Top + freButtons.Height < thisFrame.Top + thisFrame.Height Then
      Height = 2 * thisFrame.Top + thisFrame.Height + 500
   Else
      Height = 2 * freButtons.Top + freButtons.Height + 500
   End If
   
   If CanSaveValue Then
      colControlsCanSave.Add thisPowerMask, key
   End If
   
   objControls.Add thisPowerMask, strTableField
   
End Sub

Public Sub AddRichtText(ByVal key As String, ByVal Caption As String, ByVal LeftCaption As Integer, _
                        ByVal Top As Integer, ByVal LeftText As Integer, ByVal WidthText As Integer, _
                        ByVal HeightText As Integer, Optional ByVal MaxLength As Long, _
                        Optional ByVal CanSaveValue As Boolean, Optional ByVal strDefaultValue As String)
                        
Dim optionIndex         As Integer
Dim frameIndex          As Integer
Dim thisText            As RichTextBox
Dim thisLabel           As Label
Dim thisFrame           As Frame

    ' agrego un richtextbox al grupo actual
    
    frameIndex = Frame1.UBound
    Set thisFrame = Frame1(frameIndex)
   
    ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
    ' devuelve el indice del proximo frame que se creara
    optionIndex = RichTextBox1.UBound + 1
    Load Label1(Label1.UBound + 1)
    Load RichTextBox1(optionIndex)
    
    ' me creo una referencia al control para simplificar el codigo
    Set thisLabel = Label1(Label1.UBound)
    Set thisText = RichTextBox1(optionIndex)
    thisText.TabIndex = TabIndexCounter
    TabIndexCounter = TabIndexCounter + 1
    
    ' pongo el control dentro del frame adecuado
    Set thisLabel.Container = thisFrame
    Set thisText.Container = thisFrame
    
    thisLabel.Caption = Caption & ":"
    thisLabel.Width = TextWidth(thisLabel.Caption)
    thisLabel.Tag = "Label" & key
    
    thisText.Width = WidthText
    thisText.Height = HeightText
    thisText.Tag = key
    
   If Len(strDefaultValue) = 0 Then
      thisText.Text = strDefaultValue
   End If
   
    ' lo muevo a la posicion correcta
    thisLabel.Move LeftCaption, Top + 40
    thisText.Move LeftText, Top
   
   If maxWidth < LeftText + WidthText Then
      maxWidth = LeftText + WidthText
      thisFrame.Width = maxWidth + 100
      freButtons.Left = FRAME_LEFT + thisFrame.Width + FRAME_LEFT
      
      'calculo el ancho del form
      Width = freButtons.Left + freButtons.Width + 2 * FRAME_LEFT
       
      ' muevo los botones a la derecha del ultimo frame creado
      freButtons.Move freButtons.Left, freButtons.Top
   End If
   
   ' calculo la posicion del proximo control
   nextBtnTop = Top + thisText.Height + 200
   
   ' recalculo las dimensione del frame y el form
   thisFrame.Height = nextBtnTop
   thisFrame.Visible = True
   
   ' recalculo la altrura del form
   If freButtons.Top + freButtons.Height < thisFrame.Top + thisFrame.Height Then
      Height = 2 * thisFrame.Top + thisFrame.Height + 500
   Else
      Height = 2 * freButtons.Top + freButtons.Height + 500
   End If
   
   ' seteo la propiedades del control
   thisText.MaxLength = MaxLength
        
   ' lo hago visible
   thisLabel.Visible = True
   thisText.Visible = True
   
   If CanSaveValue Then
      colControlsCanSave.Add thisText, key
   End If
   
End Sub

Public Sub AddComboBox(ByVal key As String, ByVal Caption As String, ByVal LeftCaption As Integer, _
                       ByVal Top As Integer, ByVal LeftCombo As Integer, ByVal WidthCombo As Integer, _
                       ByVal strList As String, Optional ByVal DefaultValue As String, _
                       Optional ByVal CanSaveValue As Boolean)
                       
Dim ComboIndex  As Integer
Dim frameIndex       As Integer
Dim thisCombo        As ComboBox
Dim thisLabel        As Label
Dim thisFrame        As Frame
Dim ix               As Integer

   ' agrego un textbox al grupo actual
   
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   ComboIndex = Combo1.UBound + 1
   Load Label1(Label1.UBound + 1)
   Load Combo1(ComboIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisLabel = Label1(Label1.UBound)
   Set thisCombo = Combo1(ComboIndex)
   thisCombo.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   ' pongo el control dentro del frame adecuado
   Set thisLabel.Container = thisFrame
   Set thisCombo.Container = thisFrame
   
   thisLabel.Caption = Caption & ":"
   thisLabel.Width = TextWidth(thisLabel.Caption)
   thisCombo.Width = WidthCombo
   
   ' lo muevo a la posicion correcta
   thisLabel.Move LeftCaption, Top + 40
   thisCombo.Move LeftCombo, Top
   thisLabel.Tag = "Label" & key
   
   ' cargo la lista del combo
   ComboLoadList thisCombo, strList
   ' asigno valor inicial
   If Len(DefaultValue) > 0 Then
      thisCombo.ListIndex = ComboSearch(thisCombo, DefaultValue)
   Else
      thisCombo.ListIndex = -1
   End If
   
   If maxWidth < LeftCombo + thisCombo.Width Then
      maxWidth = LeftCombo + thisCombo.Width
      freButtons.Left = FRAME_LEFT + thisFrame.Width + FRAME_LEFT
      
      'calculo el ancho del form
      Width = freButtons.Left + freButtons.Width + 2 * FRAME_LEFT
      
      ' muevo los botones a la derecha del ultimo frame creado
      freButtons.Move freButtons.Left, freButtons.Top
   
   End If
   
   ' calculo la posicion del proximo control
   nextBtnTop = Top + thisCombo.Height + 200
   
   ' recalculo las dimensione del frame y el form
   thisFrame.Height = nextBtnTop
   
   ' recalculo la altrura del form
   If freButtons.Top + freButtons.Height < thisFrame.Top + thisFrame.Height Then
      Height = 2 * thisFrame.Top + thisFrame.Height + 500
   Else
      Height = 2 * freButtons.Top + freButtons.Height + 500
   End If
   
   ' seteo la propiedades del control
   thisCombo.Tag = key
       
   ' lo hago visible
   thisLabel.Visible = True
   thisCombo.Visible = True
   
   If CanSaveValue Then
      colControlsCanSave.Add thisCombo, key
   End If
   
End Sub

Function Value(ByVal key As String) As Variant
   ' devuelve el valor asociado a un option button,
   ' checkbox o frame (el valor del frame es la
   ' key del option button cuyo valor es true)
   On Error Resume Next
   Value = Values.Item(key)
    
End Function

Private Sub Check1_Click(Index As Integer)
Dim ix               As Integer
Dim oControl         As Control
Dim strKey           As String
Dim aControlsNames() As String
Dim aControls()      As Control
Dim ListControls     As String
Dim vValue           As Variant

   On Error GoTo GestErr
   
   '--
   '--   Deshabilitación de Controles
   '--
   Set oControl = Check1(Index)
   
   strKey = oControl.Tag
   
   'busco la posición del elemento buscado
   On Error Resume Next
   ListControls = colControlsDisabled.Item(strKey)
   If Err.Number = 0 Then
      If Len(ListControls) > 0 Then
         aControlsNames = Split(ListControls, ";")
         
         'completo el arreglo aControls con los controles que entran en juego
         ReDim aControls(UBound(aControlsNames))
         For ix = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(ix) = GetControl(aControlsNames(ix))
            aControls(ix).Enabled = Not (oControl.Value = 1)
            
            If TypeOf aControls(ix) Is TextBox Or _
               TypeOf aControls(ix) Is RichTextBox Or _
               TypeOf aControls(ix) Is ComboBox Then
               Set aControls(ix) = GetControl("Label" & aControlsNames(ix))
               aControls(ix).Enabled = Not (oControl.Value = 1)
            End If
            
         Next ix
   
      End If
   End If
   
   '--
   '-- Habilitación de Controles
   '--
   
   'busco la posición del elemento buscado
   On Error Resume Next
   ListControls = colControlsEnabled.Item(strKey)
   If Err.Number = 0 Then
      If Len(ListControls) > 0 Then
         aControlsNames = Split(ListControls, ";")
         
         'completo el arreglo aControls con los controles que entran en juego
         ReDim aControls(UBound(aControlsNames))
         For ix = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(ix) = GetControl(aControlsNames(ix))
            aControls(ix).Enabled = (oControl.Value = 1)
         
            If TypeOf aControls(ix) Is TextBox Or _
               TypeOf aControls(ix) Is RichTextBox Or _
               TypeOf aControls(ix) Is ComboBox Then
               Set aControls(ix) = GetControl("Label" & aControlsNames(ix))
               aControls(ix).Enabled = (oControl.Value = 1)
            End If
            
         Next ix
      End If
   End If
   
   '
   'Al volver a habilitar todos los controles, pueden quedar
   'habilitados controles que por el valor de otro control debería
   'estar deshabilitado. Por eso con los controles recientemente
   'habilitados fuerza un evento Click con el valor cambiado y el
   'valor original
   '
   
   'busco la posición del elemento buscado
   On Error Resume Next
   ListControls = colControlsEnabled.Item(strKey)
   If Err.Number = 0 Then
      If Len(ListControls) > 0 Then
         aControlsNames = Split(ListControls, ";")
         
         'completo el arreglo aControls con los controles que entran en juego
         ReDim aControls(UBound(aControlsNames))
         For ix = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(ix) = GetControl(aControlsNames(ix))
            
            vValue = aControls(ix).Value
            aControls(ix).Value = Abs(Not (vValue))
            aControls(ix).Value = vValue
            
            If TypeOf aControls(ix) Is TextBox Or _
               TypeOf aControls(ix) Is RichTextBox Or _
               TypeOf aControls(ix) Is ComboBox Then
               
               Set aControls(ix) = GetControl("Label" & aControlsNames(ix))
               
               vValue = aControls(ix).Value
               aControls(ix).Value = Abs(Not (vValue))
               aControls(ix).Value = vValue
            End If
            
         Next ix
      End If
   End If
   
   Exit Sub

GestErr:
   LoadError "Check1 (Click)"
   ShowErrMsg
End Sub

Private Sub Form_Activate()
Dim MinTop     As Single
Dim Ctrl       As Control
Dim thisFrame  As Frame
Dim frameIndex As Integer
Dim ix         As Integer

   On Error Resume Next
   
   MinTop = 999999
   
   For Each Ctrl In Controls
      If TypeName(Ctrl) <> "Menu" Then
         If Ctrl.Visible Then
            If Ctrl.Container.Name = Me.Name Then
               If Ctrl.Name <> "freButtons" Then
                   If Ctrl.Top < MinTop Then MinTop = Ctrl.Top
                End If
            End If
         End If
         
         If Ctrl.TabIndex = 0 Then Ctrl.SetFocus
         
      End If
   Next Ctrl
   
   freButtons.Top = MinTop
   
   If Not mvarShowPrintButton Then
      cmdFiltros.Top = cmdPreview.Top
   End If
   
   If Not mvarShowPrintButton And Not mvarShowFilterButton Then
      freButtons.Height = cmdCancel.Top + cmdCancel.Height + 100
   Else
      If Not mvarShowPrintButton And mvarShowFilterButton Then
         freButtons.Height = cmdFiltros.Top + cmdFiltros.Height + 100
      End If
      If mvarShowPrintButton And Not mvarShowFilterButton Then
         freButtons.Height = cmdPreview.Top + cmdPreview.Height + 100
      End If
   End If
   
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   If Not (thisFrame.Container Is Nothing) Then
      If thisFrame.Container.Name <> "frmDialog" Then
         Set thisFrame = thisFrame.Container
         thisFrame.Height = thisFrame.Height + FRAME_DISTANCE
      End If
   End If
   
   'si existe el boton guardar, lo juevo a la posicion Sup-DX
   If mvarShowSaveButton Then
      cmdGuardar.Move thisFrame.Width - cmdGuardar.Width - 100, 200
   End If
   
   Height = 600 + thisFrame.Top + thisFrame.Height
   
   If freButtons.Top + freButtons.Height > Height Then
      Height = freButtons.Top + freButtons.Height + 800
   End If

   'alinea los botones al primer contenedor
   CenterForm Me
   
End Sub


Private Sub Form_Initialize()
   
   Set mvarfrmFiltros = New frmFilter
   
   mvarShowPrintButton = False
   mvarShowSaveButton = False
   mvarShowFilterButton = False
   
   Set frmHook = New AlgStdFunc.MsgHook
   Set tmr1 = New AlgStdFunc.clsTimer
   
   Set objControls = New clsControls

End Sub

Private Sub Form_Load()

   ' reseteo todos los valores
   Set Values = Nothing
   
   Set objControls.Form = Me
'   Set objControls.FormFind = frmFind
'   Set objControls.Usuario = CUsuario
   
End Sub

Private Sub Form_Terminate()

   Set mvarfrmFiltros = Nothing
   Set objControls = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Ctrl As Control, ctrl2 As Control
Dim Response As String

    '--  antes de efectuar el unload, el form debe almacenar el valor corriente de
    '--  todos los controles en la coleccion local, asi en un futuro pueden ser
    '--  obtenidos por la  funcion Value
    

    On Error Resume Next

    For Each Ctrl In Controls
        ' si es un option button o check box
        ' almaceno el valor actual
        If TypeOf Ctrl Is OptionButton Then
            Values.Add Ctrl.Value, Ctrl.Tag
        ElseIf TypeOf Ctrl Is CheckBox Then
            Values.Add Ctrl.Value, Ctrl.Tag
        ElseIf TypeOf Ctrl Is TextBox Then
            Values.Add Trim(Ctrl.Text), Ctrl.Tag
        ElseIf TypeOf Ctrl Is PowerMask Then
            Values.Add Ctrl.Text, Ctrl.Tag
        ElseIf TypeOf Ctrl Is ComboBox Then
            Values.Add Trim(Ctrl.Text), Ctrl.Tag
        ElseIf TypeOf Ctrl Is RichTextBox Then
            Values.Add Trim(Ctrl.Text), Ctrl.Tag
        ElseIf TypeOf Ctrl Is Frame Then
            ' si es un frame,almaceno la key solo del
            ' option button cuyo valor es true
            For Each ctrl2 In Controls
                If TypeOf ctrl2 Is OptionButton Then
                    If (ctrl2.Container Is Ctrl) And ctrl2.Value = True Then
                        Values.Add ctrl2.Tag, Ctrl.Tag
                        Exit For
                    End If
                End If
            Next
        End If
    Next

   Values.Add mvarfrmFiltros.Filter1.SQLWhere, "SQLWhere"
   Values.Add mvarfrmFiltros.Filter1.FilterWhere, "FilterWhere"
   Values.Add mvarfrmFiltros.Filter1.ListConditions, "ListConditions"
   Values.Add mvarfrmFiltros.Filter1.ArrayFilterList, "ArrayFilterList"
   
   If ButtonPressed <> Cancel Then
      RaiseEvent ValidateDialog(Response)
      If Len(Response) > 0 Then
         MsgBox Response, vbExclamation, Me.Caption
         Set Values = Nothing
         Cancel = True
         Exit Sub
      End If
   End If
   
   Unload mvarfrmFiltros
   Set mvarfrmFiltros = Nothing
   
'   Unload frmFind
   
End Sub

Private Sub frmHook_AfterMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, retValue As Long)
      
   '-- si el form administrar no es MRU, al cerrarlo le llega el mensaje WM_DESTROY
   '-- si el form administrar es MRU, al cerrarlo le llega el mensaje WM_SHOWWINDOW con parametro false
   
   If uMsg = WM_DESTROY Or (uMsg = WM_SHOWWINDOW And wParam = False) Then
      frmHook.StopSubclass hWnd
      tmr1.StartTimer 100
   End If

End Sub

Private Sub mnuContextItem_Click(Index As Integer)
   objControls.mnuContextItem_Click Index
End Sub

Private Sub objControls_Messages(ByVal lngMessage As Long, Info As Variant)
'Dim hWndAdmin As Long
'
'   On Error GoTo GestErr
'
'   Select Case lngMessage
'      Case CTL_CALL_ADMIN
'
'         Dim sFormAdmin As String
'         Dim sModuloAdmin As String
'
'         If TypeName(Info) = "ControlType" Then
'            sFormAdmin = Info.FormAdmin
'            sModuloAdmin = Info.Modulo
'         End If
'
'         If sFormAdmin = NullString Then
'            MsgBox "Es probable, que en el diccionario de datos, no se haya definido un form para administrar dicho campo", vbOKOnly, App.ProductName
'            Exit Sub
'         End If
'
'         If Len(sModuloAdmin) = 0 Or UCase(sModuloAdmin) = UCase(App.ProductName) Then
'
'            '-- escondo la ventana modal, de esta manera puedo presentar otro form
'            Me.Hide
'
'            hWndAdmin = ShowForm(sFormAdmin, mvarEmpresa)
'
'            If hWndAdmin = 0 Then
'               '-- la apertura del form admin fallo
'               Me.Show vbModal
'               Exit Sub
'            End If
'
'            If Not frmHook.IsSubClassed(hWndAdmin) Then
'               frmHook.StartSubclass hWndAdmin
'            End If
'
'         Else
'            Dim objEXE As Object
'
'            Set objEXE = CreateObject(sModuloAdmin & ".Application")
'            Set objEXE.CurrentUser = CUsuario
'            objEXE.OpenForm sFormAdmin, mvarEmpresa
'
'         End If
'
'   End Select
'
'   Exit Sub
'
'GestErr:
'   LoadError "frmDialog [objControls_Messages]"
'   ShowErrMsg
End Sub

Private Sub Option1_Click(Index As Integer)
Dim ix               As Integer
Dim oControl         As Control
Dim strKey           As String
Dim aControlsNames() As String
Dim aControls()      As Control
Dim ListControls     As String
Dim vValue           As Variant

   On Error GoTo GestErr
   
   '
   'Deshabilitación de Controles
   '
   Set oControl = Option1(Index)
   
   strKey = oControl.Tag
   
   'busco la posición del elemento buscado
   On Error Resume Next
   ListControls = colControlsDisabled.Item(strKey)
   If Err.Number = 0 Then
      If Len(ListControls) > 0 Then
         aControlsNames = Split(ListControls, ";")
         
         'completo el arreglo aControls con los controles que entran en juego
         ReDim aControls(UBound(aControlsNames))
         For ix = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(ix) = GetControl(aControlsNames(ix))
            aControls(ix).Enabled = Not (oControl.Value = True)
            
            If TypeOf aControls(ix) Is TextBox Or _
               TypeOf aControls(ix) Is RichTextBox Or _
               TypeOf aControls(ix) Is ComboBox Then
               Set aControls(ix) = GetControl("Label" & aControlsNames(ix))
               aControls(ix).Enabled = Not (oControl.Value = True)
            End If
            
         Next ix
   
      End If
   End If
   
   '--
   '--   Habilitación de Controles
   '--
   
   'busco la posición del elemento buscado
   On Error Resume Next
   ListControls = colControlsEnabled.Item(strKey)
   If Err.Number = 0 Then
      If Len(ListControls) > 0 Then
         aControlsNames = Split(ListControls, ";")
         
         'completo el arreglo aControls con los controles que entran en juego
         ReDim aControls(UBound(aControlsNames))
         For ix = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(ix) = GetControl(aControlsNames(ix))
            aControls(ix).Enabled = (oControl.Value = True)
            
            If TypeOf aControls(ix) Is TextBox Or _
               TypeOf aControls(ix) Is RichTextBox Or _
               TypeOf aControls(ix) Is ComboBox Then
               Set aControls(ix) = GetControl("Label" & aControlsNames(ix))
               aControls(ix).Enabled = (oControl.Value = True)
            End If
            
         Next ix
      End If
   End If
   
   
   '
   'Al volver a habilitar todos los controles, pueden quedar
   'habilitados controles que por el valor de otro control debería
   'estar deshabilitado. Por eso con los controles recientemente
   'habilitados fuerza un evento Click con el valor cambiado y el
   'valor original
   '
   
   'busco la posición del elemento buscado
   On Error Resume Next
   ListControls = colControlsEnabled.Item(strKey)
   If Err.Number = 0 Then
      If Len(ListControls) > 0 Then
         aControlsNames = Split(ListControls, ";")
         
         'completo el arreglo aControls con los controles que entran en juego
         ReDim aControls(UBound(aControlsNames))
         For ix = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(ix) = GetControl(aControlsNames(ix))
            
            vValue = aControls(ix).Value
            aControls(ix).Value = Abs(Not (vValue))
            aControls(ix).Value = vValue
            
            If TypeOf aControls(ix) Is TextBox Or _
               TypeOf aControls(ix) Is RichTextBox Or _
               TypeOf aControls(ix) Is ComboBox Then
               Set aControls(ix) = GetControl("Label" & aControlsNames(ix))
               vValue = aControls(ix).Value
               aControls(ix).Value = Abs(Not (vValue))
               aControls(ix).Value = vValue
            End If
            
         Next ix
      End If
   End If
   
   Exit Sub

GestErr:
   LoadError "Option1 (Click)"
   ShowErrMsg

End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim thisText         As TextBox
Dim aProp()          As Variant
Dim aArray()         As Variant
Dim str              As String
Dim strExpresion     As String
Dim aControlsNames() As String
Dim aControls()      As Control
Dim aBoundField()    As String
Dim strFieldName     As String
Dim strTableName     As String
Dim ix               As Integer
Dim strCaption       As Variant

   'selecciono el textbox
   Set thisText = Text1(Index)
   
   If Not IsChanged(thisText.hWnd) Then Exit Sub
   
   On Error Resume Next
   aProp = colProperties(thisText.Tag)
   If Err.Number <> 0 Then
      ' el control es unbound
      thisText.Text = objControls.ControlValue(thisText)
      Exit Sub
   End If
      
   'si el arreglo esta lleno significa que el control es bound
   If Not IsArrayEmpty(aProp) Then

      'busco los datos para la búsqueda
      aArray = colSearch.Item(thisText.Tag)
      
      If UBound(aArray) <= 3 Then Exit Sub 'no tiene una etiqueta asociada
      
      ix = InStr(aArray(TABLE_FIELD), ".")
      
      strTableName = FieldProperty(aProp, aArray(TABLE_FIELD), dsTablaReferencia)
      strFieldName = FieldProperty(aProp, aArray(TABLE_FIELD), dsCampoReferencia)
      
      If Len(strTableName) = 0 Then
         strTableName = Left(aArray(TABLE_FIELD), ix - 1)
      End If
      
      If Len(strFieldName) = 0 Then
         strFieldName = Mid(aArray(TABLE_FIELD), ix + 1)
      End If
      
   
      aBoundField = Split(aArray(FIELD_LIST), ";")
      aControlsNames = Split(aArray(CONTROL_LIST), ";")
      
      'completo el arreglo aControls con los controles que entran en juego
      ReDim aControls(UBound(aControlsNames))
      For ix = LBound(aControlsNames) To UBound(aControlsNames)
         Set aControls(ix) = GetControl(aControlsNames(ix))
      Next ix
      
      'genero la expresion para la funcion dLookUp()
      strExpresion = NullString
      For ix = LBound(aBoundField) To UBound(aBoundField)
               
         If Len(Trim(aControls(ix))) > 0 Then
            Select Case FieldProperty(aProp, thisText.DataField, dsTipoDato)
               Case adNumeric
                  If Len(strExpresion) = 0 Then
                     strExpresion = strExpresion & aBoundField(ix) & " = " & aControls(ix)
                  Else
                     strExpresion = strExpresion & " AND " & aBoundField(ix) & " = " & aControls(ix)
                  End If
               
               Case adChar, adVarChar
                  If Len(strExpresion) = 0 Then
                     strExpresion = strExpresion & aBoundField(ix) & " = '" & aControls(ix) & "'"
                  Else
                     strExpresion = strExpresion & " AND " & aBoundField(ix) & " = '" & aControls(ix) & "'"
                  End If
         
            End Select
         
         End If
      Next ix
      
      strCaption = NullString
      If Len(strExpresion) > 0 Then
         strCaption = DLookUp(mvarEmpresa, aArray(LABEL_FIELDNAME), strTableName, strExpresion)
      End If
      If IsNull(strCaption) Then
         aArray(LABEL_CONTROL).Caption = NullString
      Else
         aArray(LABEL_CONTROL).Caption = strCaption
      End If
      aArray(LABEL_CONTROL).ZOrder 0
   End If

   thisText.Text = objControls.ControlValue(thisText)
   
End Sub

Public Function GetControl(ByVal strControlName As String) As Control
Dim cntl As Control

   For Each cntl In Me.Controls
      If UCase(cntl.Tag) = UCase(strControlName) Then
         Set GetControl = cntl
         Exit For
      End If
   Next cntl

End Function
Public Property Let Empresa(ByVal vData As String)
    mvarEmpresa = vData
End Property
Public Property Get Empresa() As String
    Empresa = mvarEmpresa
End Property
Public Property Let DialogTitle(ByVal vData As String)
    mvarDialogTitle = vData
    
    If Len(mvarDialogTitle) > 0 Then Me.Caption = mvarDialogTitle
    
End Property
Public Property Get DialogTitle() As String
    DialogTitle = mvarDialogTitle
End Property

Public Property Let CaptionsFilter(ByVal vData As String)
    mvarCaptionsFilter = vData
    
   If Len(mvarCaptionsFilter) > 0 Then
      mvarfrmFiltros.Filter1.CaptionsString = mvarCaptionsFilter
   End If
    
End Property
Public Property Let FieldsFilter(ByVal vData As String)
    mvarFieldsFilter = vData
    
   If Len(mvarFieldsFilter) > 0 Then
      mvarfrmFiltros.Filter1.ListField = mvarFieldsFilter
   End If
   
End Property
Public Property Let ShowPrintButtons(ByVal vData As Boolean)
    mvarShowPrintButton = vData
    
    cmdPreview.Visible = mvarShowPrintButton
    
End Property
Public Property Let ShowFilterButton(ByVal vData As Boolean)
    mvarShowFilterButton = vData
    
    cmdFiltros.Visible = mvarShowFilterButton
End Property

Public Sub AddFormulaToFilter(ByVal strCaption As String, Optional ByVal strConditionList As String, Optional ByVal strValueList As String)

   '-- permite agregar formulas al filtro
   
   Select Case True
      Case Len(strConditionList) = 0 And Len(strValueList) = 0
         mvarfrmFiltros.Filter1.AddFormula strCaption
      Case Len(strConditionList) > 0 And Len(strValueList) = 0
         mvarfrmFiltros.Filter1.AddFormula strCaption, strConditionList
      Case Len(strConditionList) > 0 And Len(strValueList) > 0
         mvarfrmFiltros.Filter1.AddFormula strCaption, strConditionList, strValueList
   End Select
   
End Sub

Public Sub SetFilterFormula(ByVal vExpresionFormula As Variant, ByVal strExpresionFormula As String, Optional ByVal iType As TranslateEnum = ForWhere)

   '-- permite setear las formulas del filtro

   mvarfrmFiltros.Filter1.SetFormula vExpresionFormula, strExpresionFormula, iType
   
End Sub

Private Sub tmr1_Timer()
   
   '-- vuelvo a presentar el frmFilter modal
   tmr1.StopTimer
   Me.Show vbModal

End Sub
