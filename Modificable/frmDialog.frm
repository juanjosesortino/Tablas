VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{955CADBF-E2F9-11D4-94E4-00E07D72826B}#38.0#0"; "PowerMaskControl.ocx"
Object = "{B97E3E11-CC61-11D3-95C0-00C0F0161F05}#172.0#0"; "ALGControls.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   3510
      Visible         =   0   'False
      Width           =   735
   End
   Begin ALGControls.ALGLine ALGLine1 
      Height          =   45
      Index           =   0
      Left            =   165
      TabIndex        =   18
      Top             =   1515
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      X1              =   50
      X2              =   1800
      Y1              =   10
      Y2              =   10
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   0
      Left            =   1530
      TabIndex        =   14
      Top             =   3465
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame freButtons 
      BorderStyle     =   0  'None
      Height          =   2760
      Left            =   4320
      TabIndex        =   11
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton cmdOK 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   60
         TabIndex        =   12
         ToolTipText     =   "Acepta"
         Top             =   30
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   60
         TabIndex        =   13
         ToolTipText     =   "Cancela"
         Top             =   450
         Width           =   1125
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Vista Previa"
         Height          =   375
         Left            =   60
         TabIndex        =   15
         ToolTipText     =   "Vista Previa "
         Top             =   1320
         Width           =   1125
      End
      Begin VB.CommandButton cmdFiltros 
         Caption         =   "&Otros Filtros"
         Height          =   375
         Left            =   45
         TabIndex        =   16
         ToolTipText     =   "Otros Filtros"
         Top             =   1755
         Width           =   1125
      End
      Begin VB.CommandButton cmdAvanzado 
         Caption         =   "&Avanzado ..."
         Height          =   375
         Left            =   45
         TabIndex        =   17
         ToolTipText     =   "Otros Filtros"
         Top             =   2340
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Index           =   0
      Left            =   390
      TabIndex        =   9
      Top             =   3900
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin PowerMaskControl.PowerMask PowerMask1 
      Height          =   315
      Left            =   1000
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar Opciones"
      Height          =   375
      Left            =   4260
      TabIndex        =   7
      ToolTipText     =   "Guardar Opciones"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   2805
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
      Left            =   255
      TabIndex        =   6
      Top             =   3420
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDialog.frx":0000
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   915
      Index           =   0
      Left            =   4290
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1614
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
      TabIndex        =   5
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
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   3
      Top             =   3090
      Visible         =   0   'False
      Width           =   645
      WordWrap        =   -1  'True
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

Private WithEvents m_TextControls As clsControlItems
Attribute m_TextControls.VB_VarHelpID = -1

Private Const FRAME_LEFT = 100                          ' left del frame
Private Const FRAME_DISTANCE = 100                      ' distancia entre frames
Private Const OPTION_LEFT = 200                         ' valor del Left para todos los controles check, option dentro de un frame
Private Const OPTION_DISTANCE = 50                      ' distancia entre options/checks

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

Private ErrorLog                As ErrType
Public mvarDialogTitle          As String            ' Titulo del Dialog
Public WithEvents objControls     As clsControls       'instancia de la clase clsControls
Attribute objControls.VB_VarHelpID = -1

Private nextBtnTop               As Single            ' es el Top para proximo control
Private maxWidth                 As Single            ' es el ancho del frame
Private TabIndexCounter          As Integer           ' es el contador del TabIndex
Private bControlsPlaced          As Boolean           ' indica que los controles fueron ubicados en su posicion y dimensionados

Private Values                   As New Collection    ' coleccion devuelta al llamador con todos los valores de los controles
Private ValuesFormatted          As New Collection    ' coleccion devuelta al llamador con todos los valores de los controles formateados
Private colFormat                As New Collection    ' coleccion de objetos textbox que debe ser formateados. Tienen un evento LostFocus para ejecutar
Private colKeyPress              As New Collection    ' coleccion de objetos que tienen un evento KeyPress para ejecutar
Private colSearch                As New Collection    ' coleccion de datos para QueryDB
Private colProperties            As New Collection    ' coleccion de propiedades
Private colDataType              As New Collection    ' coleccion de tipos para cada control que necesitan validación
Private colControlsDisabled      As New Collection    ' coleccion de tipos para cada control que necesitan validación
Private colControlsEnabled       As New Collection    ' coleccion de tipos para cada control que necesitan validación
Private colControlsCanSave       As New Collection    ' coleccion de controles que pueden salvar su valor de default
Private colFrames                As New Collection    ' coleccion frames. Por c/frame guardo el tipo de dimension (fija/Automatico)
Private colDlookUp               As New Collection    ' coleccion de expresiones Where para la función DlookUp

Public ButtonPressed             As EnumButtonPressedDialog

Private mvarfrmFiltros           As frmFilter         ' form frmFilter
Private mvarCaptionsFilter       As String            ' Lista de captions para el filtro avanzado
Private mvarFieldsFilter         As String            ' Lista de campos para el filtro avanzado
Private mvarShowPrintButton      As Boolean
Private mvarShowSaveButton       As Boolean           ' determina si será visible el boton para "Salvar Opciones"
Private mvarShowFilterButton     As Boolean
Private mvarShowAdvancedButton   As Boolean
Private mvarShowCancelButton     As Boolean
Private mvarShowOkButton         As Boolean
Private mvarControlData          As DataShare.udtControlData         'información de control
Private mvarMenuKey              As String                           'clave del menu

Private mvarDialogName           As String            ' Nombre del dialogo
Private mvarAcceptContinueEnabled As Boolean          ' indica si el boton Aceptar descarga el dialogo despues de Aceptar
Private mvarCallerObjName        As String            ' Nombre del Form/Clase que llama al dialog

Private objTabla                 As New BOGeneral.clsTablas
Private aKeys()                  As Variant
Private aArray()                 As Variant
Private vValue                   As Variant
Private strKey                   As String
Private lvwColumn                As ColumnHeader
Private IX                       As Integer

Private mvarSetDefaultEditData   As Boolean 'Bug #5958 LGA

' definicion de los eventos del form
Public Event Activate()
Public Event ValidateDialog(ByRef Response As String)
Public Event CommandButtonClick(ByVal Index As Integer, ByVal Key As String)
Public Event LostFocus(ByVal ControlKey As String)
Public Event ItemClick(ByVal ControlKey As String, ByVal Item As MSComCtlLib.ListItem)
Public Event ItemCheck(ByVal ControlKey As String, ByVal Item As MSComCtlLib.ListItem)
Public Event Click(ByVal ControlKey As String)
Public Event Change(ByVal Index As Integer, ByVal Key As String)
Public Event BeforeButtonClick(ByVal Button As EnumButtonPressedDialog)
Public Event ButtonClick(ByVal Button As EnumButtonPressedDialog)
Public Event AfterButtonClick(ByVal Button As EnumButtonPressedDialog)
Public Event Messages(ByVal Key As String, ByVal lngMessage As Long, ByRef Info As Variant)
Public Event ColumnClick(ByVal ControlKey As String, ByVal ColumnHeader As MSComCtlLib.ColumnHeader)
Public Event MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event ItemDblClick(ByVal ControlKey As String)
Public Event KeyDown(ByVal ControlKey As String, KeyCode As Integer, Shift As Integer)

'Agregados para poder establecer el orden de las especies en el reporte (TP 6136.)
Public Event OLEDragDrop(Index As Integer, data As MSComCtlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, ByRef objControl As Control)                   'TP 6136. SAMSA L4 - 30) Nuevo reporte para el financiero
Public Event OLEDragOver(Index As Integer, data As MSComCtlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer, ByRef objControl As Control) 'TP 6136. SAMSA L4 - 30) Nuevo reporte para el financiero
Public Event OLEStartDrag(Index As Integer, data As MSComCtlLib.DataObject, AllowedEffects As Long, ByRef objControl As Control)                                                                         'TP 6136. SAMSA L4 - 30) Nuevo reporte para el financiero

Public Enum eDlgAlignText
   AlineaDerecha = 0
   AlineaInferior = 1
End Enum
Private bSetDefault              As Boolean

Public Property Let SetDefaultEditData(ByVal vData As Boolean) 'Bug #5958 LGA
    mvarSetDefaultEditData = vData
End Property
Public Property Get SetDefaultEditData() As Boolean 'Bug #5958 LGA
    SetDefaultEditData = mvarSetDefaultEditData
End Property


Public Sub AddFrame(ByVal Key As String, ByVal Caption As String, Optional ByVal showBorder As Boolean = True, _
                    Optional ByVal ShowSaveButton As Boolean, Optional ByVal KeyParent As String, _
                    Optional ByVal FrameHeight As Integer, Optional ByVal FrameWidth As Integer)
                    
Dim frameIndex    As Integer
Dim thisFrame     As Frame
Dim prevFrame     As Frame
    
    ' agrego un nuevo frame
    
    
    ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
    ' devuelve el indice del proximo frame que se creara
   On Error GoTo GestErr

    frameIndex = Frame1.UBound + 1
    
    ' cargo un nuevo frame
    Load Frame1(frameIndex)
    Set thisFrame = Frame1(frameIndex)
    
    ' setea las propiedades del frame
    thisFrame.Caption = Caption
    thisFrame.Tag = Key
    ' valores posibles para el BorderStyle son 0 (none) o 1 (visible)
    thisFrame.BorderStyle = -(showBorder)
    thisFrame.Visible = True
    If Len(KeyParent) > 0 Then
      Set thisFrame.Container = GetControl(KeyParent)
    End If
    
    thisFrame.Width = 0
    thisFrame.Height = 0
    
    If FrameHeight > 0 Then
      thisFrame.Height = FrameHeight
    End If
    If FrameWidth > 0 Then
      thisFrame.Width = FrameWidth
    End If
    
    'esta coleccion si el frame dimension fija o es calculada dinamicamente
    If FrameHeight > 0 Or FrameWidth > 0 Then
      colFrames.Add "Fijo", "K" & frameIndex
    Else
      colFrames.Add "Automatico", "K" & frameIndex
    End If
    
    ' seteo la posicion del proximo control
    maxWidth = 0  ' ***
    
    'obtengo el frame anterior
    Set prevFrame = Frame1(frameIndex - 1)
    
    ' si este no es el primer frame, lo muevo debajo del precedente
    Select Case frameIndex
      Case Is = 1
         'es el primer Frame
         nextBtnTop = 0
         If colFrames(frameIndex) = "Automatico" Then
            thisFrame.Move FRAME_LEFT, FRAME_DISTANCE, thisFrame.Width, thisFrame.Height
         Else
            thisFrame.Move FRAME_LEFT, FRAME_DISTANCE, thisFrame.Width, thisFrame.Height
         End If
         
      Case Is > 1
         If Len(KeyParent) = 0 Then
            ' es un Frame padre
            nextBtnTop = 0
            If colFrames(frameIndex) = "Automatico" Then
               thisFrame.Move FRAME_LEFT, prevFrame.Top + prevFrame.Height + FRAME_DISTANCE, prevFrame.Width, thisFrame.Height
            Else
               thisFrame.Move FRAME_LEFT, prevFrame.Top + prevFrame.Height + FRAME_DISTANCE, thisFrame.Width, thisFrame.Height
            End If
         Else
            ' es un Frame Hijo. El nuevo frame lo coloco en debajo del ultimo control
            Dim ParentFrame As Frame
            Dim ctrl As Control
            Dim MaxTop As Integer
            Set ParentFrame = GetControl(KeyParent)
                        
            MaxTop = nextBtnTop
            For Each ctrl In Controls

               If ctrl.Tag <> NullString Then

                  'busco el control bas bajo
                  If ctrl.Container.Tag = KeyParent Then
                     If ctrl.Top + ctrl.Height > MaxTop Then
                        MaxTop = ctrl.Top + ctrl.Height
                     End If
                  End If
               End If
               
            Next ctrl
            
            thisFrame.Move FRAME_LEFT, MaxTop + FRAME_DISTANCE, ParentFrame.Width - 2 * FRAME_LEFT, thisFrame.Height
            
         End If
    End Select
    
   If ShowSaveButton Then
      mvarShowSaveButton = True
      Set cmdGuardar.Container = thisFrame
      cmdGuardar.Visible = True
   End If

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddFrame"
   ShowErrMsg ErrorLog
    
End Sub

Public Sub AddOption(ByVal Key As String, ByVal Caption As String, _
                     Optional ByVal Top As Long, Optional ByVal Left As Long, _
                     Optional Value As Boolean, Optional ByVal ControlListDisabled As String, _
                     Optional ByVal ControlListEnabled As String, Optional ByVal CanSaveValue As Boolean, _
                     Optional ByVal Width As Integer)
                     
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim thisBtn       As OptionButton
Dim thisFrame     As Frame
Dim FrameContainer As Frame
    
   ' agreo un Option al grupo corriente
   
   ' este es el numero del corriente grupo de controles
   On Error GoTo GestErr

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   
   ' agrago un nuovo option
   optionIndex = Option1.UBound + 1
   Load Option1(optionIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisBtn = Option1(optionIndex)
   thisBtn.TabIndex = TabIndexCounter
   thisBtn.Caption = Caption   '***
   thisBtn.Width = TextWidth(Caption) + 500
   
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
   If Top >= nextBtnTop Then
      nextBtnTop = Top + thisBtn.Height + OPTION_DISTANCE
   End If
   
   ' ajusto las dimensiones del frame '***
   If colFrames.Item("K" & frameIndex) = "Automatico" Then
      
      If thisFrame.Index = 1 Then
         If maxWidth < Left + thisBtn.Width Then
            maxWidth = Left + thisBtn.Width
            thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < Top + thisBtn.Height + OPTION_DISTANCE Then
         thisFrame.Height = Top + thisBtn.Height + OPTION_DISTANCE
      End If
      
   End If
   
   If Not (thisFrame.Container Is Nothing) Then
      If thisFrame.Container.Name <> "frmDialog" Then
         Set FrameContainer = thisFrame.Container
         FrameContainer.Height = thisFrame.Top + nextBtnTop
      End If
   End If
   
   If Len(ControlListDisabled) > 0 Then
      colControlsDisabled.Add ControlListDisabled, Key        'no necesito el item, solo la clave
   End If
   
   If Len(ControlListEnabled) > 0 Then
      colControlsEnabled.Add ControlListEnabled, Key          'no necesito el item, solo la clave
   End If
   
   ' seteo la propiedades del option
   thisBtn.Tag = Key
   thisBtn.Value = Value
   
   ' lo hago visible
   thisBtn.Visible = True
   
   If CanSaveValue Then
      colControlsCanSave.Add thisBtn, Key
   End If
   
'   CenterForm Me

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddOption"
   ShowErrMsg ErrorLog
   
End Sub

Public Sub AddCheck(ByVal Key As String, ByVal Caption As String, _
                    Optional ByVal Top As Long, Optional ByVal Left As Long, _
                    Optional Value As Variant, Optional ByVal ControlListDisabled As String, _
                    Optional ByVal ControlListEnabled As String, Optional ByVal CanSaveValue As Boolean, _
                    Optional ByVal WidthCheck As Integer, Optional ByVal Alignment As Integer)
                    
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim ThisCheck     As CheckBox
Dim thisFrame     As Frame
Dim FrameContainer As Frame
    
   ' agrego un check al grupo corriente
   
   ' este es el numero del corriente grupo de controles
   On Error GoTo GestErr

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = Check1.UBound + 1
   Load Check1(optionIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set ThisCheck = Check1(optionIndex)
   ThisCheck.TabIndex = TabIndexCounter
   ThisCheck.Caption = Caption    '***
   If WidthCheck > 0 Then
      ThisCheck.Width = WidthCheck
   Else
      ThisCheck.Width = TextWidth(Caption) + 500
   End If
   TabIndexCounter = TabIndexCounter + 1
   
   ' pongo el control dentro del frame adecuado
   Set ThisCheck.Container = thisFrame
   ' lo muevo a la posicion correcta
   If Top > 0 Then
      ThisCheck.Move OPTION_LEFT, Top
   End If
   If Left > 0 Then
      ThisCheck.Move Left, ThisCheck.Top
   End If
   If Top = 0 And Left = 0 Then
      ThisCheck.Move OPTION_LEFT, nextBtnTop
   End If
   'seteo propiedad de alineación del texto
   ThisCheck.Alignment = IIf(Alignment = 1, 1, 0)
   
   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + ThisCheck.Height + OPTION_DISTANCE
   End If
   
   ' ajusto el ancho del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then  '***
      
      If thisFrame.Index = 1 Then
         If maxWidth < Left + ThisCheck.Width Then   ' ***
            maxWidth = Left + ThisCheck.Width
            thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < Top + ThisCheck.Height + OPTION_DISTANCE Then
         thisFrame.Height = Top + ThisCheck.Height + OPTION_DISTANCE
      End If
      
   End If
   
   If Not (thisFrame.Container Is Nothing) Then
     If thisFrame.Container.Name <> "frmDialog" Then
        Set FrameContainer = thisFrame.Container
        FrameContainer.Height = FrameContainer.Height + nextBtnTop
     End If
   End If
   
   If Len(ControlListDisabled) > 0 Then
     colControlsDisabled.Add ControlListDisabled, Key        'no necesito el item, solo la clave
   End If
   
   If Len(ControlListEnabled) > 0 Then
     colControlsEnabled.Add ControlListEnabled, Key          'no necesito el item, solo la clave
   End If
   
   
   ' seteo la propiedades del check
   ThisCheck.Tag = Key
   ThisCheck.Value = vbChecked
   
   Select Case True
      Case IsMissing(Value)
         Value = vbUnchecked
      Case Value = si
         Value = vbChecked
      Case Value = No
         Value = vbUnchecked
      Case Value = True
         Value = vbChecked
      Case Value = False
         Value = vbUnchecked
      
      Case Value = vbChecked     'absi
         Value = vbChecked       'absi
      Case Value = vbUnchecked   'absi
         Value = vbUnchecked     'absi
         
      Case Else
         Value = vbUnchecked
      End Select
      
   ThisCheck.Value = Abs(Value)
   
   ' lo hago visible
   ThisCheck.Visible = True
   
   If CanSaveValue Then
     colControlsCanSave.Add ThisCheck, Key
   End If
   
'   CenterForm Me

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddCheck"
   ShowErrMsg ErrorLog
   
End Sub

Public Sub AddLabel(ByVal Key As String, ByVal Caption As String, _
                    Optional ByVal LeftCaption As Integer, _
                    Optional ByVal TopCaption As Integer, Optional ByVal color As Long, _
                    Optional ByVal WidthCaption As Integer, _
                    Optional ByVal Bold As Boolean, _
                    Optional ByVal iAlignment As Integer)
                    
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim ThisLabel     As Label
Dim thisFrame     As Frame

   ' agrego un label al grupo corriente

   On Error GoTo GestErr

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
   Set ThisLabel = Label1(optionIndex)
   ThisLabel.Caption = Caption
   Me.FontBold = Bold
   ThisLabel.FontBold = Bold
   ThisLabel.Alignment = iAlignment
   
   If WidthCaption <> 0 Then
      ThisLabel.Width = WidthCaption
   Else
      ThisLabel.Width = TextWidth(Caption)
   End If
   Me.FontBold = False
   
   

   ' calculo la posicion del proximo control
   If TopCaption >= nextBtnTop Then
      nextBtnTop = TopCaption + ThisLabel.Height + 200
   End If

   'ajusto el ancho del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then    '***

      If thisFrame.Index = 1 Then
         If maxWidth < LeftCaption + ThisLabel.Width Then
            maxWidth = LeftCaption + ThisLabel.Width
            thisFrame.Width = maxWidth + 100
         End If
      End If

      If thisFrame.Height < TopCaption + ThisLabel.Height + 200 Then
         thisFrame.Height = TopCaption + ThisLabel.Height + 200
      End If

   End If

   ' pongo el control dentro del frame adecuado
   Set ThisLabel.Container = thisFrame

   ' lo muevo a la posicion correcta
   ThisLabel.Move LeftCaption, TopCaption

   ' calculo la posicion del proximo control
   nextBtnTop = nextBtnTop + ThisLabel.Height + OPTION_DISTANCE

   ' seteo la propiedades del check
   ThisLabel.Tag = Key
   If color > 0 Then
      ThisLabel.ForeColor = color
   End If

   ' lo hago visible
   ThisLabel.Visible = True


   Exit Sub

GestErr:
   LoadError ErrorLog, "AddLabel"
   ShowErrMsg ErrorLog

End Sub

Public Sub AddLabelDescription(ByVal Key As String, Optional ByVal Left As Integer, Optional ByVal Top As Integer, Optional ByVal BackColor As Long, _
                                                    Optional ByVal Width As Integer, Optional ByVal strDefaultValue As String)
                    
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim ThisLabel     As Label
Dim thisFrame     As Frame

   ' agrego un label al grupo corriente

   On Error GoTo GestErr

   If Top = 0 Then
      Top = nextBtnTop
   End If

   If Left = 0 Then
      Left = OPTION_LEFT
   End If


   ' este es el numero del corriente grupo de controles
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)

   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = lblDescripcion.UBound + 1
   Load lblDescripcion(optionIndex)

   ' me creo una referencia al control para simplificar el codigo
   Set ThisLabel = lblDescripcion(optionIndex)
   If Width <> 0 Then
      ThisLabel.Width = Width
   End If

   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + ThisLabel.Height + 200
   End If

   'ajusto el ancho del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then    '***

      If thisFrame.Index = 1 Then
         If maxWidth < Left + ThisLabel.Width Then
            maxWidth = Left + ThisLabel.Width
            thisFrame.Width = maxWidth + 100
         End If
      End If

      If thisFrame.Height < Top + ThisLabel.Height + 200 Then
         thisFrame.Height = Top + ThisLabel.Height + 200
      End If

   End If

   ' pongo el control dentro del frame adecuado
   Set ThisLabel.Container = thisFrame

   ' lo muevo a la posicion correcta
   ThisLabel.Move Left, Top

   ' calculo la posicion del proximo control
   nextBtnTop = nextBtnTop + ThisLabel.Height + OPTION_DISTANCE

   ' seteo la propiedades del check
   ThisLabel.Tag = Key
   If BackColor > 0 Then
      ThisLabel.BackColor = BackColor
   End If

   ThisLabel.Caption = strDefaultValue
   
   ' lo hago visible
   ThisLabel.Visible = True

   'si ya viene con un valor, salvo su valor
   If Len(strDefaultValue) <> 0 Then SaveProperties ThisLabel


   Exit Sub

GestErr:
   LoadError ErrorLog, "AddLabelDescription"
   ShowErrMsg ErrorLog

End Sub

Public Sub AddListView(ByVal Key As String, ByVal Caption As String, _
                       ByVal Top As Integer, ByVal Left As Integer, _
                       ByVal lvwWidth As Integer, ByVal lvwHeight As Integer, _
                       Optional ByVal strColumnList As String, _
                       Optional ByVal strColumnWidth As String, _
                       Optional ByVal strColumnAlign As String, _
                       Optional ByVal strColumnDataType As String) 'Inc.SDP 91683 para ordenar el listView al clickear la cabecera de una columna
                   
Dim optionIndex         As Integer
Dim frameIndex          As Integer
Dim thisListView        As ListView
Dim ThisLabel           As Label
Dim thisFrame           As Frame
Dim aCaps()             As String
Dim aWidths()           As String
Dim aAlign()            As String
Dim aDataType()         As String

   ' agrego un textbox al grupo actual
   
   On Error GoTo GestErr

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = ListView1.UBound + 1
   If Len(Caption) > 0 Then
      'creo un label
      Load Label1(Label1.UBound + 1)
      Set ThisLabel = Label1(Label1.UBound)
      Set ThisLabel.Container = thisFrame
      ThisLabel.Caption = Caption & ":"
      ThisLabel.Width = TextWidth(ThisLabel.Caption)
      ThisLabel.Move Left, Top - 200
      ThisLabel.Visible = True
      ThisLabel.Tag = "Label" & Key
      
   End If
   
   ' me creo una referencia al control para simplificar el codigo
   Load ListView1(optionIndex)
   Set thisListView = ListView1(optionIndex)
   Set thisListView.Container = thisFrame
   thisListView.Width = lvwWidth
   thisListView.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   If Len(strColumnList) > 0 Then
   
      aCaps = Split(strColumnList, ";")
      aWidths = Split(strColumnWidth, ";")
      aAlign = Split(strColumnAlign, ";")
      aDataType = Split(strColumnDataType, ";")
         
      For IX = LBound(aCaps) To UBound(aCaps)
         Set lvwColumn = thisListView.ColumnHeaders.Add(, , aCaps(IX), aWidths(IX))
         If Len(strColumnAlign) > 0 Then
             Select Case aAlign(IX)
                Case NullString, "L"
                   lvwColumn.Alignment = lvwColumnLeft
                Case NullString, "R"
                   lvwColumn.Alignment = lvwColumnRight
                Case NullString, "C"
                   lvwColumn.Alignment = lvwColumnCenter
                Case Else
                  lvwColumn.Alignment = lvwColumnLeft
            End Select
         End If
         If Len(strColumnDataType) > 0 Then   'Inc.SDP 91683
            Select Case aDataType(IX)
               Case NullString, "NUMBER"
                  lvwColumn.Tag = "NUMBER"
               Case NullString, "DATE"
                  lvwColumn.Tag = "DATE"
               Case NullString, "STRING"
                  lvwColumn.Tag = "STRING"
               Case Else
                  lvwColumn.Tag = "NUMBER"
            End Select
         End If
      Next IX
      
   End If
      
   thisListView.Tag = Key
   thisListView.Visible = True
   ' lo muevo a la posicion correcta
   thisListView.Move Left, Top
   thisListView.Height = lvwHeight
   
   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + thisListView.Height + 200
   End If
   
   ' ajusto las dimensiones del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then  '***
  
      If thisFrame.Index = 1 Then
         If maxWidth < Left + lvwWidth Then
            maxWidth = Left + lvwWidth
            thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < Top + thisListView.Height + 200 Then
         thisFrame.Height = Top + thisListView.Height + 200
      End If
   
   End If
   
'   CenterForm Me

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddListView"
   ShowErrMsg ErrorLog
   
End Sub

Public Sub AddHLine(ByVal Top As Long, ByVal Left As Long, ByVal Lenght As Long)
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim thisLine      As Line
Dim thisFrame     As Frame
    
   ' agrego un check al grupo corriente
   
   ' este es el numero del corriente grupo de controles
   On Error GoTo GestErr

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
   
   ' ajusto el ancho del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then  '***
   
      If thisFrame.Index = 1 Then
         If maxWidth < Left + Lenght Then '***
           maxWidth = Left + Lenght
           thisFrame.Width = maxWidth + 100
         End If
      End If
      
   End If
   
   ' pongo el control dentro del frame adecuado
   Set thisLine.Container = thisFrame
   
   ' lo hago visible
   thisLine.Visible = True

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddHLine"
   ShowErrMsg ErrorLog

End Sub

Public Sub AddVLine(ByVal Top As Long, ByVal Left As Long, ByVal Height As Long)
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim thisLine      As Line
Dim thisFrame     As Frame
    
   ' agrego un check al grupo corriente
   
   ' este es el numero del corriente grupo de controles
   On Error GoTo GestErr

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

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddVLine"
   ShowErrMsg ErrorLog

End Sub

Public Sub AddLine(ByVal Key As String, ByVal Top As Long, _
                   ByVal Left As Long, ByVal Lenght As Long, _
                   Optional ByVal Orientation As Integer, _
                   Optional ByVal CaptionLine As String, _
                   Optional ByVal Bold As Boolean = False)
                  
Dim optionIndex   As Integer
Dim frameIndex    As Integer
Dim thisLine      As ALGLine
Dim thisFrame     As Frame
    
    
   ' agrego un check al grupo corriente
   
   ' este es el numero del corriente grupo de controles
   On Error GoTo GestErr

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   optionIndex = ALGLine1.UBound + 1
   Load ALGLine1(optionIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisLine = ALGLine1(optionIndex)
   thisLine.Top = Top
   thisLine.Left = Left
   If Orientation = 0 Then
      'horizontal
      thisLine.Width = Lenght
   Else
      thisLine.Height = Lenght
   End If
   
   thisLine.Caption = CaptionLine
   thisLine.FontBold = Bold
   
   ' pongo el control dentro del frame adecuado
   Set thisLine.Container = thisFrame
   
   ' lo hago visible
   thisLine.Visible = True
   
   ' ajusto el ancho del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then  '***
   
      If thisFrame.Index = 1 Then
         If maxWidth < Left + Lenght Then '***
           maxWidth = Left + Lenght
           thisFrame.Width = maxWidth + 100
         End If
      End If
      
   End If
   
   thisLine.Tag = Key

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddLine"
   ShowErrMsg ErrorLog

End Sub

Public Sub AddTextLabel(ByVal Key As String, ByVal Caption As String, _
                        ByVal LeftCaption As Integer, _
                        ByVal Top As Integer, _
                        Optional ByVal AlignText As eDlgAlignText = AlineaDerecha, _
                        Optional ByVal LeftText As Integer, _
                        Optional ByVal LeftDescription As Integer, _
                        Optional ByVal WidthText As Integer, Optional ByVal WidthDescription As Integer, _
                        Optional ByVal strTableField As String, _
                        Optional ByVal strBoundLabel As String, Optional ByVal FindControlList As String, _
                        Optional ByVal FindBoundFieldList As String, Optional ByVal CanSaveValue As Boolean, _
                        Optional ByVal strDefaultValue As String, Optional ByVal strLabelDefaultValue As String, _
                        Optional ByVal strWhereDlookUp As String, Optional ByVal Enabled As Boolean = True)
                        
                        
Dim frameIndex          As Integer
Dim thisText            As TextBox
Dim ThisLabel           As Label
Dim thisDescription     As Label
Dim thisFrame           As Frame
Dim aTableProperties    As Variant


'/
'  Ej:
'      .AddTextLabel Key:="txtProductor", Caption:="Productor", LeftCaption:=270, _
'                    Top:=300, LeftText:=1710, LeftDescription:=3375, WidthText:=1590, _
'                    WidthDescription:=3750, strTableField:="PRODUCTORES.PRO_PRODUCTOR", strBoundLabel:="ENT_NOMBRE", _
'                    FindBoundFieldList:="PRO_PRODUCTOR", CanSaveValue:=False, _
'                    strDefaultValue:="", strLabelDefaultValue:="", strWhereDlookUp:="PRO_TIPO_ENTIDAD,3;PRO_PRODUCTOR,%1"
'
' En este caso, al indicar strWhereDlookUp:="PRO_TIPO_ENTIDAD,3;PRO_PRODUCTOR,%1", cuando se produzca el evento LostFocus, el dato
' que se muestra en la etiqueta se busca en base a la query definida en el diccionario para el campo PRODUCTORES.PRO_PRODUCTOR y a
' esa query se le agrega a la Where :
' PRO_TIPO_ENTIDAD = 3 and PRO_PRODUCTOR = [valor de control]
' Si no se define el parametro strWhereDlookUp, entonces se usa la funcion dLookUp automáticamente
'/


   'Este metodo se usa para los Bound TextBox por eso es que no es necesario. objControls
   'obtiene la mayor parte de la informacion desde el diccionario

   ' agrego un textbox al grupo actual
   
   On Error GoTo GestErr

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   If Len(Caption) > 0 Then
     'creo un label
     Load Label1(Label1.UBound + 1)
     Set ThisLabel = Label1(Label1.UBound)
     Set ThisLabel.Container = thisFrame
     ThisLabel.Caption = Caption & ":"
     ThisLabel.Width = TextWidth(ThisLabel.Caption)
     ThisLabel.Move LeftCaption, Top + 40
     ThisLabel.Tag = "Label" & Key
     ThisLabel.Visible = True
   
     ' lo muevo a la posicion correcta
     ThisLabel.Move LeftCaption, Top
     ThisLabel.Visible = True
   
   End If
   
   ' me creo una referencia al control para simplificar el codigo
 
   ' agrego el Text a la colección de controles
   Set thisText = Me.Controls.Add("VB.TextBox", Key)
   thisText.Height = 315
   
   Dim ControlItem As New clsControlItem
   
   Set ControlItem.MyControl = thisText
   m_TextControls.Add ControlItem
 
 
 
   Set thisText.Container = thisFrame
   thisText.Width = WidthText
   thisText.Text = NullString
   thisText.Tag = Key
   thisText.TabIndex = TabIndexCounter
   thisText.Enabled = Enabled
   TabIndexCounter = TabIndexCounter + 1
   
   
   ' lo muevo a la posicion correcta
   
   If AlignText = AlineaDerecha Then
      'uso el mismo top del caption
      If LeftText = 0 Then
         thisText.Move LeftCaption + ThisLabel.Width + 100, Top - 50
      Else
         thisText.Move LeftText, Top - 50
      End If
   Else
      'uso el mismo left del caption
      If LeftText = 0 Then
         thisText.Move LeftCaption, (Top + ThisLabel.Height + 30)
      Else
         thisText.Move LeftText, (Top + ThisLabel.Height + 30)
      End If
   End If
   
   ' me creo una referencia al control para simplificar el codigo
   Load lblDescripcion(lblDescripcion.UBound + 1)
   Set thisDescription = lblDescripcion(lblDescripcion.UBound)
   Set thisDescription.Container = thisFrame
   thisDescription.Width = WidthDescription
   thisDescription.Tag = "lbl" & Key
   thisDescription.Caption = ""
   
   ' lo muevo a la posicion correcta
   If AlignText = AlineaDerecha Then
      'alineo el powermask a la derecha del caption, uso el mismo top del caption
      If LeftDescription = 0 Then
         thisDescription.Move (thisText.Left + thisText.Width + 30), Top - 50
      Else
         thisDescription.Move LeftDescription, Top - 50
      End If
   Else
      'powermask debajo del Caption
      If LeftDescription = 0 Then
         thisDescription.Move (LeftCaption + thisText.Width + 30), (Top + ThisLabel.Height + 30)
      Else
         thisDescription.Move LeftDescription, thisText.Top
      End If
   End If
   
   ' lo hago visible

   thisText.Visible = True
   thisDescription.Visible = True
   thisDescription.Enabled = Enabled
   
   '-- Información necesaria para la actualización:
   '--    Entrada diccionario
   '--    Nombres de TextBox actualizado por QueryDB (lista separada por ";")
   '--    Nombres de campos de QueryDB que actualizan los TextBox (lista separada por ";")
   '--    Control Label Actualizado por QueryDB
   '--    Campo del QueryDB que actualiza el Label descriptivo
   '--    Opcional: Tabla a leer si es distinta a la de la entrada en el diccionario (ej. Transportistas)
   '--
   
   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + thisText.Height + 200
   End If
   
   ' ajusto las dimensiones del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then  '***
   
      If thisFrame.Index = 1 Then
         If maxWidth < thisDescription.Left + thisDescription.Width Then
           maxWidth = thisDescription.Left + thisDescription.Width
           thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < Top + thisText.Height + 200 Then
         thisFrame.Height = Top + thisText.Height + 200
      End If
   
   End If
   
   If Len(strTableField) > 0 Then
   
      aTableProperties = GetFieldInformation(strTableField)
      
      'agrego datos para la búsqueda a la colección
      
      If Len(FindControlList) = 0 Then FindControlList = Key
      
      colSearch.Add Array(strTableField, FindControlList, FindBoundFieldList, thisDescription, strBoundLabel), Key
      
      
      colProperties.Add aTableProperties, Key 'guardo la propiedades del campo o bien un
                                              'arreglo vacio si el control es UnBound
      
   End If
   
'   If Len(strDefaultValue) > 0 Then
'      thisText.Text = strDefaultValue
'   End If
'   If Len(strDefaultValue) > 0 Then
'      thisDescription.Caption = strLabelDefaultValue
'   End If
   
   
   ' seteo la propiedades del control
   thisText.Tag = Key
      
   If CanSaveValue Then
      colControlsCanSave.Add thisText, Key
   End If
   
   If Len(strWhereDlookUp) > 0 Then
      colDlookUp.Add strWhereDlookUp, Key
   End If
   
   objControls.Add thisText, strTableField, , , FindControlList, FindBoundFieldList
   
   If Len(strDefaultValue) > 0 Then
      thisText.Text = strDefaultValue
   End If
   If Len(strDefaultValue) > 0 Then
      thisDescription.Caption = strLabelDefaultValue
   End If
   
   'si ya viene con un valor, salvo su valor
   If Len(strDefaultValue) <> 0 Then SaveProperties thisText
   
'   CenterForm Me

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddTextLabel"
   ShowErrMsg ErrorLog
   
End Sub

Public Sub AddText(ByVal Key As String, ByVal Caption As String, _
                   ByVal LeftCaption As Integer, _
                   ByVal Top As Integer, _
                   Optional ByVal AlignText As eDlgAlignText = AlineaDerecha, _
                   Optional ByVal LeftText As Integer, _
                   Optional ByVal WidthText As Integer, _
                   Optional ByVal strTableField As String, Optional ByVal strDefaultValue As String, _
                   Optional ByVal CanSaveValue As Boolean, Optional ByVal Validar As ValidEnum, _
                   Optional ByVal Dimension As Integer, Optional ByVal Decimales As Integer, _
                   Optional ByVal LOCKED As Boolean, Optional ByVal Enabled As Boolean = True, _
                   Optional ByVal FindControlList As String, Optional ByVal FindBoundFieldList As String, _
                   Optional ByVal Alignment As Integer, _
                   Optional ByVal Formatear As eFormatField = [_NoDefinido], _
                   Optional ByVal Multiline As Boolean, _
                   Optional ByVal Mascara As String = NullString)
                   
Dim frameIndex          As Integer
Dim thisText            As TextBox
Dim ThisLabel           As Label
Dim thisFrame           As Frame
Dim aTableProperties    As Variant


   ' agrego un textbox al grupo actual
   
   On Error GoTo GestErr

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara

   If Len(Caption) > 0 Then
      'creo un label
      Load Label1(Label1.UBound + 1)
      Set ThisLabel = Label1(Label1.UBound)
      Set ThisLabel.Container = thisFrame
      ThisLabel.Caption = Caption & ":"
      ThisLabel.Width = TextWidth(ThisLabel.Caption)
      ThisLabel.Move LeftCaption, Top + 40
      ThisLabel.Visible = True
      ThisLabel.Tag = "Label" & Key
      
   End If
   
   ' me creo una referencia al control para simplificar el codigo
   
   ' agrego el Text a la colección de controles
   If IsNumeric(Key) Then Key = "K" & Key
   Set thisText = Me.Controls.Add("VB.TextBox", Key)
   
   'Si es multiline trabajamos con el control creado en tpo de diseño Text1
   If Multiline Then
      Text1.Height = 1300
      Set thisText = Text1
      thisText.Tag = Text1.Name
   Else
      thisText.Height = 315
      thisText.Tag = Key
   End If
   
   
   Dim ControlItem As New clsControlItem
   Set ControlItem.MyControl = thisText
   m_TextControls.Add ControlItem
   
   
   Set thisText.Container = thisFrame
   thisText.Width = WidthText
   thisText.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
'   If Len(strDefaultValue) > 0 Then
'      thisText.Text = strDefaultValue
'   End If
   
   thisText.Alignment = Alignment  'vbLeftJustify (0) es el default
   
   SetModify thisText.hWnd, False
   
   thisText.Visible = True
   
   ' lo muevo a la posicion correcta
   
   If AlignText = AlineaDerecha Then
      'uso el mismo top del caption
      If LeftText = 0 Then
         thisText.Move LeftCaption + ThisLabel.Width + 30, Top - 20
      Else
         thisText.Move LeftText, Top
      End If
   Else
      'uso el mismo left del caption
      thisText.Move LeftCaption, (Top + ThisLabel.Height + 30)
   End If
   
   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + thisText.Height + 200
   End If
   
   ' ajusto las dimensiones del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then  '***
   
      If thisFrame.Index = 1 Then
         If maxWidth < LeftText + WidthText Then
            maxWidth = LeftText + WidthText
            thisFrame.Width = maxWidth + 100
         End If
      End If
      
      ' recalculo las dimensione del frame
      If thisFrame.Height < Top + thisText.Height + 200 Then
         thisFrame.Height = Top + thisText.Height + 200
      End If
   
   End If
   
   If Len(strTableField) > 0 Then

      aTableProperties = GetFieldInformation(strTableField)
      
      'agrego datos para la búsqueda a la colección
      
      If Multiline Then
      
         If Len(FindControlList) = 0 Then FindControlList = Text1.Name
         colSearch.Add Array(strTableField, FindControlList, FindBoundFieldList), Key
         colProperties.Add aTableProperties, Text1.Name
      
      Else

         If Len(FindControlList) = 0 Then FindControlList = Key
         colSearch.Add Array(strTableField, FindControlList, FindBoundFieldList), Key
         colProperties.Add aTableProperties, Key 'guardo la propiedades del campo o bien un
                                              'arreglo vacio si el control es UnBound
      End If
      
      Dimension = FieldProperty(aTableProperties, strTableField, dsDimension)
      Decimales = FieldProperty(aTableProperties, strTableField, dsDecimales)

   End If
   
   'Formatear = IIf(Len(Mascara) > 0, eFormatField.FormatSi, eFormatField.FormatNo)
   
   objControls.Add thisText, strTableField, , Validar, FindControlList, FindBoundFieldList, Dimension, Decimales, , Formatear, Mascara
   
   If CanSaveValue Then
      colControlsCanSave.Add thisText, Key
   End If
   
   thisText.LOCKED = LOCKED
   thisText.Enabled = Enabled
   
   If Len(strDefaultValue) > 0 Then
      thisText.Text = strDefaultValue
   End If
   
   'si ya viene con un valor, salvo su valor
   If Len(strDefaultValue) <> 0 Then SaveProperties thisText
   
   If Len(thisText.Text) > 0 Then
'      Text1_LostFocus (thisText.Index)
   End If
   
'   CenterForm Me

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddText"
   ShowErrMsg ErrorLog
   
End Sub

Public Sub AddPowerMaskLabel(ByVal Key As String, ByVal Caption As String, _
                              ByVal LeftCaption As Integer, ByVal Top As Integer, _
                              Optional ByVal AlignText As eDlgAlignText = AlineaDerecha, _
                              Optional ByVal LeftText As Integer, _
                              Optional ByVal LeftDescription As Integer, _
                              Optional ByVal WidthText As Integer, _
                              Optional ByVal WidthDescription As Integer, _
                              Optional ByVal strTableField As String, _
                              Optional ByVal strBoundLabel As String, _
                              Optional ByVal FindControlList As String, _
                              Optional ByVal FindBoundFieldList As String, _
                              Optional ByVal CanSaveValue As Boolean, _
                              Optional ByVal strDefaultValue As String, _
                              Optional ByVal strLabelDefaultValue As String, _
                              Optional ByVal strWhereDlookUp As String)
                        
Dim frameIndex          As Integer
Dim ThisPowerMask       As PowerMask
Dim ThisLabel           As Label
Dim thisDescription     As Label
Dim thisFrame           As Frame
Dim aTableProperties    As Variant

'/
'  Ej:
'      .AddPowerMaskLabel Key:="pmkProductor", Caption:="Productor", LeftCaption:=270, _
'                        Top:=300, LeftText:=1710, LeftDescription:=3375, WidthText:=1590, _
'                        WidthDescription:=3750, strTableField:="PRODUCTORES.PRO_PRODUCTOR", strBoundLabel:="ENT_NOMBRE", _
'                        FindBoundFieldList:="PRO_PRODUCTOR", CanSaveValue:=False, _
'                        strDefaultValue:="", strLabelDefaultValue:="", strWhereDlookUp:="PRO_TIPO_ENTIDAD,3;PRO_PRODUCTOR,%1"
'
' En este caso, al indicar strWhereDlookUp:="PRO_TIPO_ENTIDAD,3;PRO_PRODUCTOR,%1", cuando se produzca el evento LostFocus, el dato
' que se muestra en la etiqueta se busca en base a la query definida en el diccionario para el campo PRODUCTORES.PRO_PRODUCTOR y a
' esa query se le agrega a la Where :
' PRO_TIPO_ENTIDAD = 3 and PRO_PRODUCTOR = [valor de control]
' Si no se define el parametro strWhereDlookUp, entonces se usa la funcion dLookUp automáticamente
'/


   'Este metodo se usa para los Bound PowerMask.
   'objControls obtiene la mayor parte de la informacion desde el diccionario

   ' agrego un PowerMask al grupo actual
   
   On Error GoTo GestErr

   aTableProperties = GetFieldInformation(strTableField)
   
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara

   If Len(Caption) > 0 Then
     'creo un label
     Load Label1(Label1.UBound + 1)
     Set ThisLabel = Label1(Label1.UBound)
     Set ThisLabel.Container = thisFrame
     ThisLabel.Caption = Caption & ":"
     ThisLabel.Width = TextWidth(ThisLabel.Caption)
     ThisLabel.Move LeftCaption, Top + 40
     ThisLabel.Tag = "Label" & Key
     ThisLabel.Visible = True
   End If
   ' lo muevo a la posicion correcta
   ThisLabel.Move LeftCaption, Top
   
   ' me creo una referencia al control para simplificar el codigo
   
   ' agrego el Text a la colección de controles
   Set ThisPowerMask = Me.Controls.Add("PowerMaskControl.PowerMask", Key)
   ThisPowerMask.Height = 315
   
   Dim ControlItem As New clsControlItem
   
   Set ControlItem.MyControl = ThisPowerMask
   m_TextControls.Add ControlItem
   
   
   Set ThisPowerMask.Container = thisFrame
   ThisPowerMask.mask = FieldProperty(aTableProperties, strTableField, dsCodigoEspecial)
   ThisPowerMask.Width = WidthText
   ThisPowerMask.Text = NullString
   ThisPowerMask.Tag = Key
   ThisPowerMask.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   ' lo muevo a la posicion correcta
   If AlignText = AlineaDerecha Then
      'uso el mismo top del caption
      If LeftText = 0 Then
         ThisPowerMask.Move LeftCaption + ThisLabel.Width + 100, Top - 50
      Else
         ThisPowerMask.Move LeftText, Top
      End If
   Else
      'muevo el powermask debajo del Caption, uso el mismo left del caption
      If LeftText = 0 Then
         ThisPowerMask.Move LeftCaption, (Top + ThisLabel.Height + 30)
      Else
         ThisPowerMask.Move LeftText, (Top + ThisLabel.Height + 30)
      End If
   End If
   
   
   ' me creo una referencia al control para simplificar el codigo
   Load lblDescripcion(lblDescripcion.UBound + 1)
   Set thisDescription = lblDescripcion(lblDescripcion.UBound)
   Set thisDescription.Container = thisFrame
   thisDescription.Width = WidthDescription
   thisDescription.Tag = "lbl" & Key
   thisDescription.Caption = ""
   
   ' lo muevo a la posicion correcta
   If AlignText = AlineaDerecha Then
      'alineo el powermask a la derecha del caption, uso el mismo top del caption
      If LeftDescription = 0 Then
         thisDescription.Move (ThisPowerMask.Left + ThisPowerMask.Width + 30), Top
      Else
         thisDescription.Move LeftDescription, Top
      End If
   Else
      'powermask debajo del Caption
      If LeftDescription = 0 Then
         thisDescription.Move (LeftCaption + ThisPowerMask.Width + 30), (Top + ThisLabel.Height + 30)
      Else
         thisDescription.Move LeftDescription, ThisPowerMask.Top
      End If
   End If
   
   
   ' lo hago visible
   ThisLabel.Visible = True
   ThisPowerMask.Visible = True
   thisDescription.Visible = True
   
   '-- Información necesaria para la actualización:
   '--    Entrada diccionario
   '--    Nombres de TextBox actualizado por QueryDB (lista separada por ";")
   '--    Nombres de campos de QueryDB que actualizan los TextBox (lista separada por ";")
   '--    Control Label Actualizado por QueryDB
   '--    Campo del QueryDB que actualiza el Label descriptivo
   '--    Opcional: Tabla a leer si es distinta a la de la entrada en el diccionario (ej. Transportistas)
   '--
   
   If Len(FindControlList) = 0 Then FindControlList = Key
   
   colSearch.Add Array(strTableField, FindControlList, FindBoundFieldList, thisDescription, strBoundLabel), Key
   
   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + ThisPowerMask.Height + 200
   End If
   
   ' ajusto las dimensiones del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then  '***
   
      If thisFrame.Index = 1 Then
         If maxWidth < LeftDescription + WidthDescription Then
           maxWidth = LeftDescription + WidthDescription
           thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < Top + ThisPowerMask.Height + 200 Then
         thisFrame.Height = Top + ThisPowerMask.Height + 200
      End If
   
   End If
   
   ' seteo la propiedades del control
   ThisPowerMask.Tag = Key
   
   ThisPowerMask.Text = strDefaultValue
   thisDescription.Caption = strLabelDefaultValue
   
   colProperties.Add aTableProperties, Key 'guardo la propiedades del campo
   
   If CanSaveValue Then
      colControlsCanSave.Add ThisPowerMask, Key
   End If
   
   If Len(strWhereDlookUp) > 0 Then
      colDlookUp.Add strWhereDlookUp, Key
   End If
   
   objControls.Add ThisPowerMask, strTableField, , , FindControlList, FindBoundFieldList
   
   'si ya viene con un valor, salvo su valor
   If Len(strDefaultValue) <> 0 Then SaveProperties ThisPowerMask
   
'   CenterForm Me

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddPowerMaskLabel"
   ShowErrMsg ErrorLog
   
End Sub

Public Sub AddPowerMask(ByVal Key As String, ByVal Caption As String, _
                        ByVal LeftCaption As Integer, ByVal Top As Integer, _
                        Optional ByVal AlignText As eDlgAlignText = AlineaDerecha, _
                        Optional ByVal LeftText As Integer, _
                        Optional ByVal WidthText As Integer, _
                        Optional ByVal strTableField As String, _
                        Optional ByVal mask As CodigosEspecialesEnum, _
                        Optional ByVal strDefaultValue As String, _
                        Optional ByVal CanSaveValue As Boolean)
                   
Dim frameIndex          As Integer
Dim ThisPowerMask       As PowerMask
Dim ThisLabel           As Label
Dim thisFrame           As Frame
Dim strField            As String
Dim aTableProperties    As Variant

   'me fijo en base a la entrada del diccionario, su máscara
   On Error GoTo GestErr

   If Len(strTableField) > 0 Then
   
      aTableProperties = GetFieldInformation(strTableField)
      
      strField = FieldProperty(aTableProperties, strTableField, dsCampo)

      'agrego datos para la búsqueda a la colección
      colSearch.Add Array(strTableField, Key, strField), Key
      
      
      colProperties.Add aTableProperties, Key 'guardo la propiedades del campo o bien un
                                              'arreglo vacio si el control es UnBound
                                              
      mask = FieldProperty(aTableProperties, strTableField, dsCodigoEspecial)
   
   End If
   
   ' agrego un PowerMask al grupo actual
   
   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   If Len(Caption) > 0 Then
      'creo un label
      Load Label1(Label1.UBound + 1)
      Set ThisLabel = Label1(Label1.UBound)
      Set ThisLabel.Container = thisFrame
      ThisLabel.Caption = Caption & ":"
      ThisLabel.Width = TextWidth(ThisLabel.Caption)
      ThisLabel.Move LeftCaption, Top + 40
      ThisLabel.Visible = True
      ThisLabel.Tag = "Label" & Key
      
   End If
   
   ' me creo una referencia al control para simplificar el codigo
   
   ' agrego el Text a la colección de controles
   Set ThisPowerMask = Me.Controls.Add("PowerMaskControl.PowerMask", Key)
   ThisPowerMask.Height = 315
   
   Dim ControlItem As New clsControlItem
   
   Set ControlItem.MyControl = ThisPowerMask
   m_TextControls.Add ControlItem
   
   
   Set ThisPowerMask.Container = thisFrame
   ThisPowerMask.mask = mask
   ThisPowerMask.Width = WidthText
   ThisPowerMask.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   If Len(strDefaultValue) > 0 Then
      ThisPowerMask.Text = strDefaultValue
   End If
   
   ThisPowerMask.TabStop = True
   
   ThisPowerMask.SetModify False
   ThisPowerMask.Tag = Key
   ThisPowerMask.Visible = True
   
   ' lo muevo a la posicion correcta
   If AlignText = AlineaDerecha Then
      'uso el mismo top del caption
      If LeftText = 0 Then
         ThisPowerMask.Move LeftCaption + ThisLabel.Width + 30, Top - 20
      Else
         ThisPowerMask.Move LeftText, Top
      End If
   Else
      'uso el mismo left del caption
      ThisPowerMask.Move LeftCaption, (Top + ThisLabel.Height + 30)
   End If
   
   
   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + ThisPowerMask.Height + 200
   End If
   
   ' ajusto las dimensiones del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then  '***
  
      If thisFrame.Index = 1 Then
         If maxWidth < LeftText + WidthText Then
            maxWidth = LeftText + WidthText
            thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < Top + ThisPowerMask.Height + 200 Then
         thisFrame.Height = Top + ThisPowerMask.Height + 200
      End If
   
   End If
   
   If CanSaveValue Then
      colControlsCanSave.Add ThisPowerMask, Key
   End If
   
   objControls.Add ThisPowerMask, strTableField, , , ThisPowerMask.Tag
   
   'si ya viene con un valor, salvo su valor
   If Len(strDefaultValue) <> 0 Then SaveProperties ThisPowerMask
   
'   CenterForm Me

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddPowerMask"
   ShowErrMsg ErrorLog
   
End Sub

Public Sub AddRichtText(ByVal Key As String, ByVal Caption As String, ByVal LeftCaption As Integer, _
                        ByVal Top As Integer, ByVal LeftText As Integer, ByVal WidthText As Integer, _
                        ByVal HeightText As Integer, Optional ByVal MaxLength As Long, _
                        Optional ByVal CanSaveValue As Boolean, Optional ByVal strDefaultValue As String)
                        
Dim optionIndex         As Integer
Dim frameIndex          As Integer
Dim thisText            As RichTextBox
Dim ThisLabel           As Label
Dim thisFrame           As Frame

    ' agrego un richtextbox al grupo actual
    
    On Error GoTo GestErr

    frameIndex = Frame1.UBound
    Set thisFrame = Frame1(frameIndex)
   
    ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
    ' devuelve el indice del proximo frame que se creara
    optionIndex = RichTextBox1.UBound + 1
    Load Label1(Label1.UBound + 1)
    Load RichTextBox1(optionIndex)
    
    ' me creo una referencia al control para simplificar el codigo
    Set ThisLabel = Label1(Label1.UBound)
    Set thisText = RichTextBox1(optionIndex)
    thisText.TabIndex = TabIndexCounter
    TabIndexCounter = TabIndexCounter + 1
    
    ' pongo el control dentro del frame adecuado
    Set ThisLabel.Container = thisFrame
    Set thisText.Container = thisFrame
    
    If Caption <> NullString Then
      ThisLabel.Caption = Caption & ":"
      ThisLabel.Width = TextWidth(ThisLabel.Caption)
      ThisLabel.Tag = "Label" & Key
   Else
      ThisLabel.Caption = Caption
      ThisLabel.Visible = False
   End If
    
    thisText.Width = WidthText
    thisText.Height = HeightText
    thisText.Tag = Key
    
   If Len(strDefaultValue) > 0 Then
      thisText.Text = strDefaultValue
   End If
   
   ' lo muevo a la posicion correcta
   ThisLabel.Move LeftCaption, Top + 40
   thisText.Move LeftText, Top
   
   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + thisText.Height + 200
   End If
   
   ' ajusto las dimensiones del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then  '***
   
      If thisFrame.Index = 1 Then
         If maxWidth < LeftText + WidthText Then
            maxWidth = LeftText + WidthText
            thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < Top + thisText.Height + 200 Then
         thisFrame.Height = Top + thisText.Height + 200
      End If
      
   End If
   
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
   ThisLabel.Visible = True
   thisText.Visible = True
   
   If CanSaveValue Then
      colControlsCanSave.Add thisText, Key
   End If
   
   'si ya viene con un valor, salvo su valor
   If Len(strDefaultValue) <> 0 Then SaveProperties thisText
   
   Exit Sub

GestErr:
   LoadError ErrorLog, "AddRichtText"
   ShowErrMsg ErrorLog
   
End Sub


Public Sub AddCommand(ByVal Key As String, ByVal Caption As String, ByVal LeftButton As Integer, _
                      ByVal TopButton As Integer, ByVal WidthButton As Integer, Optional ByVal Enabled As Boolean = True, _
                      Optional ByVal Bold As Boolean = False)

Dim CommandIndex     As Integer
Dim frameIndex       As Integer
Dim thisCommand      As CommandButton
Dim thisFrame        As Frame

   ' agrego un textbox al grupo actual
   
   On Error GoTo GestErr

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   CommandIndex = Command1.UBound + 1
   Load Command1(CommandIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisCommand = Command1(CommandIndex)
   thisCommand.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   ' pongo el control dentro del frame adecuado
   Set thisCommand.Container = thisFrame
   
   thisCommand.Width = WidthButton
   thisCommand.FontBold = Bold
   
   
   ' lo muevo a la posicion correcta
   thisCommand.Move LeftButton, TopButton
   
   ' calculo la posicion del proximo control
   If TopButton >= nextBtnTop Then
      nextBtnTop = TopButton + thisCommand.Height + 200
   End If
   
   thisCommand.Caption = Caption
   
   ' ajusto las dimensiones del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then ' ***
   
      If thisFrame.Index = 1 Then
         If maxWidth < LeftButton + thisCommand.Width Then
            maxWidth = LeftButton + thisCommand.Width
            thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < TopButton + thisCommand.Height + 200 Then
         thisFrame.Height = TopButton + thisCommand.Height + 200
      End If
      
   End If
   
   ' seteo la propiedades del control
   thisCommand.Tag = Key
       
   ' lo hago visible
   thisCommand.Visible = True
   
   thisCommand.Enabled = Enabled
'   CenterForm Me

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddCommand"
   ShowErrMsg ErrorLog
                       
End Sub

Public Sub AddComboBox(ByVal Key As String, ByVal Caption As String, ByVal LeftCaption As Integer, _
                       ByVal Top As Integer, ByVal LeftCombo As Integer, ByVal WidthCombo As Integer, _
                       ByVal strList As String, Optional ByVal DefaultValue As String, _
                       Optional ByVal CanSaveValue As Boolean, Optional ByVal Enabled As Boolean = True)
                       
Dim ComboIndex  As Integer
Dim frameIndex       As Integer
Dim thisCombo        As ComboBox
Dim ThisLabel        As Label
Dim thisFrame        As Frame
Dim IX               As Integer

   ' agrego un textbox al grupo actual
   
   On Error GoTo GestErr

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   ComboIndex = Combo1.UBound + 1
   Load Label1(Label1.UBound + 1)
   Load Combo1(ComboIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set ThisLabel = Label1(Label1.UBound)
   Set thisCombo = Combo1(ComboIndex)
   thisCombo.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   ' pongo el control dentro del frame adecuado
   Set ThisLabel.Container = thisFrame
   Set thisCombo.Container = thisFrame
   
   ' seteo la propiedades del control
   thisCombo.Tag = Key
   
   ThisLabel.Caption = Caption & ":"
   ThisLabel.Width = TextWidth(ThisLabel.Caption)
   thisCombo.Width = WidthCombo
   
   ' lo muevo a la posicion correcta
   ThisLabel.Move LeftCaption, Top + 40
   thisCombo.Move LeftCombo, Top
   ThisLabel.Tag = "Label" & Key
   
   ' cargo la lista del combo
'   ComboLoadList thisCombo, strList
   
   Dim aSplit() As String
   aSplit = Split(strList, ";")
   If Len(strList) > 0 Then
      For IX = LBound(aSplit) To UBound(aSplit)
         thisCombo.AddItem aSplit(IX)
      Next IX
   End If
   
   ' asigno valor inicial
   If Len(DefaultValue) > 0 Then
      thisCombo.ListIndex = ComboSearch(thisCombo, DefaultValue)
   Else
      thisCombo.ListIndex = -1
   End If
   
   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + thisCombo.Height + 200
   End If
   
   ' ajusto las dimensiones del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then ' ***
   
      If thisFrame.Index = 1 Then
         If maxWidth < LeftCombo + thisCombo.Width Then
            maxWidth = LeftCombo + thisCombo.Width
            thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < Top + thisCombo.Height + 200 Then
         thisFrame.Height = Top + thisCombo.Height + 200
      End If
      
   End If
   
       
   ' lo hago visible
   ThisLabel.Visible = True
   thisCombo.Visible = True
   thisCombo.Enabled = Enabled
   
   If CanSaveValue Then
      colControlsCanSave.Add thisCombo, Key
   End If
   
'   CenterForm Me

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddComboBox"
   ShowErrMsg ErrorLog
   
End Sub
Public Sub AddDataCombo(ByVal Key As String, ByVal Caption As String, ByVal LeftCaption As Integer, _
                       ByVal Top As Integer, ByVal LeftDataCombo As Integer, ByVal WidthDataCombo As Integer, _
                       ByVal strTableName As String, ByVal strBoundText As String, ByVal strBoundColumn As String, Optional ByVal DefaultValue As String, _
                       Optional ByVal CanSaveValue As Boolean, Optional ByVal rstSource As ADODB.Recordset, Optional ByVal Enabled As Boolean = True)
                       
Dim DataComboIndex   As Integer
Dim frameIndex       As Integer
Dim thisDataCombo    As DataCombo
Dim ThisLabel        As Label
Dim thisFrame        As Frame

   ' agrego un DataCombo al grupo actual
   
   On Error GoTo GestErr

   frameIndex = Frame1.UBound
   Set thisFrame = Frame1(frameIndex)
   
   ' dado que los arreglo de controles se basan en el indice cero, el ubound + 1
   ' devuelve el indice del proximo frame que se creara
   DataComboIndex = DataCombo1.UBound + 1
   Load Label1(Label1.UBound + 1)
   Load DataCombo1(DataComboIndex)
   
   ' me creo una referencia al control para simplificar el codigo
   Set thisDataCombo = DataCombo1(DataComboIndex)
   thisDataCombo.TabIndex = TabIndexCounter
   TabIndexCounter = TabIndexCounter + 1
   
   ' pongo el control dentro del frame adecuado
   Set thisDataCombo.Container = thisFrame
   
   If Len(Caption) > 0 Then
      Set ThisLabel = Label1(Label1.UBound)
      Set ThisLabel.Container = thisFrame
   
      ThisLabel.Caption = Caption & ":"
      ThisLabel.Width = TextWidth(ThisLabel.Caption)
      ThisLabel.Tag = "Label" & Key
      
      ThisLabel.Move LeftCaption, Top + 40
      
      ThisLabel.Visible = True
      
   End If
   thisDataCombo.Width = WidthDataCombo
   
   ' lo muevo a la posicion correcta
   thisDataCombo.Move LeftDataCombo, Top
   
   ' cargo la lista del combo
   DataComboLoad thisDataCombo, mvarControlData.Empresa, strTableName, strBoundText, strBoundColumn, rstSource, DefaultValue
   
   ' calculo la posicion del proximo control
   If Top >= nextBtnTop Then
      nextBtnTop = Top + thisDataCombo.Height + 200
   End If
   
   ' ajusto las dimensiones del frame
   If colFrames.Item("K" & frameIndex) = "Automatico" Then ' ***
   
      If thisFrame.Index = 1 Then
         If maxWidth < LeftDataCombo + thisDataCombo.Width Then
            maxWidth = LeftDataCombo + thisDataCombo.Width
            thisFrame.Width = maxWidth + 100
         End If
      End If
      
      If thisFrame.Height < Top + thisDataCombo.Height + 200 Then
         thisFrame.Height = Top + thisDataCombo.Height + 200
      End If
   
   End If
   
   ' seteo la propiedades del control
   thisDataCombo.Tag = Key
       
   ' lo hago visible
   
   thisDataCombo.Visible = True
   
   If CanSaveValue Then
      colControlsCanSave.Add thisDataCombo, Key
   End If
   
   thisDataCombo.Enabled = Enabled
   
   objControls.Add thisDataCombo, strTableName & "." & strBoundText, , , , strBoundText
   
'   CenterForm Me

   Exit Sub

GestErr:

   LoadError ErrorLog, "AddDataCombo"
   ShowErrMsg ErrorLog
   
End Sub

Function Value(ByVal Key As String) As Variant
   ' devuelve el valor asociado a un option button,
   ' checkbox o frame (el valor del frame es la
   ' key del option button cuyo valor es true)
   On Error Resume Next
   
   Value = Values.Item(Key)
   
End Function
Function ValueFormatted(ByVal Key As String) As Variant
   ' devuelve el valor asociado a un option button,
   ' checkbox o frame (el valor del frame es la
   ' key del option button cuyo valor es true)
   On Error Resume Next
   
   ValueFormatted = ValuesFormatted.Item(Key)
   
End Function

Private Sub Check1_Click(Index As Integer)
Dim IX               As Integer
Dim ThisCheck        As CheckBox
Dim strKey           As String
Dim aControlsNames() As String
Dim aControls()      As Control
Dim ListControls     As String
Dim vValue           As Variant

   On Error GoTo GestErr

   On Error GoTo GestErr
   
   '--
   '--   Deshabilitación de Controles
   '--
   Set ThisCheck = Check1(Index)
   
   strKey = ThisCheck.Tag
   
   'salvo el nuevo valor
   SaveProperties ThisCheck
   
   RaiseEvent Click(strKey)
   
   'busco la posición del elemento buscado
   On Error Resume Next
   ListControls = colControlsDisabled.Item(strKey)
   If Err.Number = 0 Then
      If Len(ListControls) > 0 Then
         aControlsNames = Split(ListControls, ";")
         
         'completo el arreglo aControls con los controles que entran en juego
         ReDim aControls(UBound(aControlsNames))
         For IX = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(IX) = GetControl(aControlsNames(IX))
            aControls(IX).Enabled = Not (ThisCheck.Value = 1)
            
            If TypeOf aControls(IX) Is TextBox Or _
               TypeOf aControls(IX) Is RichTextBox Or _
               TypeOf aControls(IX) Is ComboBox Then
               Set aControls(IX) = GetControl("Label" & aControlsNames(IX))
               aControls(IX).Enabled = Not (ThisCheck.Value = vbChecked)
            End If
            
         Next IX
   
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
         For IX = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(IX) = GetControl(aControlsNames(IX))
            aControls(IX).Enabled = (ThisCheck.Value = vbChecked)
         
            If TypeOf aControls(IX) Is TextBox Or _
               TypeOf aControls(IX) Is RichTextBox Or _
               TypeOf aControls(IX) Is ComboBox Then
               Set aControls(IX) = GetControl("Label" & aControlsNames(IX))
               aControls(IX).Enabled = (ThisCheck.Value = vbChecked)
            End If
            
         Next IX
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
         For IX = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(IX) = GetControl(aControlsNames(IX))
            
            vValue = aControls(IX).Value
            aControls(IX).Value = Abs(Not (vValue))
            aControls(IX).Value = vValue
            
            If TypeOf aControls(IX) Is TextBox Or _
               TypeOf aControls(IX) Is RichTextBox Or _
               TypeOf aControls(IX) Is ComboBox Then
               
               Set aControls(IX) = GetControl("Label" & aControlsNames(IX))
               
               vValue = aControls(IX).Value
               aControls(IX).Value = Abs(Not (vValue))
               aControls(IX).Value = vValue
            End If
            
         Next IX
      End If
   End If
   
   Exit Sub

GestErr:
   LoadError ErrorLog, "Check1_Click"
   ShowErrMsg ErrorLog
End Sub

Private Sub Check1_LostFocus(Index As Integer)
   RaiseEvent LostFocus(Check1(Index).Tag)
End Sub

Private Sub cmdAvanzado_Click()
   
   RaiseEvent BeforeButtonClick(Advanced)
    
   'salvo las propiedades
   SaveProperties
   
   RaiseEvent ButtonClick(Advanced)
   RaiseEvent AfterButtonClick(Advanced)
   
End Sub

Private Sub cmdFiltros_Click()
   
   RaiseEvent BeforeButtonClick(Filter)
   
   'salvo las propiedades
   SaveProperties
   
   'informo al objeto llamador que ha sido llamado el form de filtros
   RaiseEvent ButtonClick(Filter)

   mvarfrmFiltros.Show vbModal
   
   'salvo las propiedades
   SaveProperties
   
   RaiseEvent AfterButtonClick(Filter)
   
End Sub

Private Sub cmdGuardar_Click()
Dim cntl          As Control
Dim strKey        As String
Dim strValor      As String
Dim aKeysToSave() As String
Dim IX            As Integer

   On Error GoTo GestErr
   
   If colControlsCanSave.Count = 0 Then Exit Sub
   
   Me.MousePointer = vbHourglass
   
   ReDim aKeysToSave(colControlsCanSave.Count - 1, 1)

   IX = 0
   For Each cntl In colControlsCanSave
      
      Select Case TypeName(cntl)
         Case "TextBox", "RichTextBox", "ComboBox"
            strValor = cntl.Text
         Case "CheckBox"
            strValor = cntl.Value
         Case "OptionButton"
            strValor = cntl.Value
      End Select
         
      strKey = "Dialogs\Defaults\" & CUsuario.Usuario & "\" & Me.Caption & "\" & cntl.Tag
      aKeysToSave(IX, 0) = strKey
      aKeysToSave(IX, 1) = strValor
      
      IX = IX + 1
   Next cntl
   
   objTabla.ControlData = mvarControlData
   objTabla.UpdateGlobal NullString, aKeysToSave

   Me.MousePointer = vbDefault
   
   Exit Sub

GestErr:
   Me.MousePointer = vbDefault
   LoadError ErrorLog, "frmDialog [cmdGuardar_Click]"
   ShowErrMsg ErrorLog
End Sub

Private Sub cmdOK_Click()
Dim Response As String
   
   RaiseEvent BeforeButtonClick(Accept)
   
   'salvo las propiedades
   SaveProperties
      
   'dejo que el objeto llamador valide los datos del form
   RaiseEvent ValidateDialog(Response)
   If Len(Response) > 0 Then
   
      MsgBox Response, vbExclamation, Me.Caption
      
      Set Values = Nothing
      Set ValuesFormatted = Nothing
      
      Exit Sub
   End If
   
   ButtonPressed = Accept
   
   
   'informo al objeto llamado que los datos han sido aceptados
   RaiseEvent ButtonClick(Accept)
   
   If mvarAcceptContinueEnabled = False Then
      Unload Me
   Else
      Dim IX As Integer
      For IX = Frame1.LBound To Frame1.UBound
         SetFrameControls Me, Frame1(IX), True, True, True
      Next IX
   End If
   
   RaiseEvent AfterButtonClick(Accept)
   
End Sub

Private Sub cmdCancel_Click()

   RaiseEvent BeforeButtonClick(Cancel)

   ButtonPressed = Cancel
   
   RaiseEvent ButtonClick(Cancel)
   Unload Me
    
   RaiseEvent AfterButtonClick(Cancel)
    
End Sub

Private Sub cmdPreview_Click()
Dim Response As String
    
   On Error GoTo GestErr

   RaiseEvent BeforeButtonClick(Preview)
    
   ButtonPressed = Preview
    
   'salvo las propiedades
   SaveProperties
   
   'dejo que el objeto llamador valide los datos del form
   RaiseEvent ValidateDialog(Response)
   If Len(Response) > 0 Then
      MsgBox Response, vbExclamation, Me.Caption
      Set Values = Nothing
      Set ValuesFormatted = Nothing
      Exit Sub
   End If
    
   'informo al objeto llamador el Preview de los datos
   RaiseEvent ButtonClick(Preview)
   
   RaiseEvent AfterButtonClick(Preview)

   Exit Sub

GestErr:
   LoadError ErrorLog, "cmdPreview_Click"
   ShowErrMsg ErrorLog
   
End Sub

Private Sub Combo1_Change(Index As Integer)
   RaiseEvent Change(Index, Combo1(Index).Tag)
End Sub

Private Sub Combo1_Click(Index As Integer)
   RaiseEvent Click(Combo1(Index).Tag)
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
   RaiseEvent LostFocus(Combo1(Index).Tag)
End Sub

Private Sub Command1_Click(Index As Integer)
   RaiseEvent CommandButtonClick(Index, Command1(Index).Tag)
End Sub

Private Sub Command1_LostFocus(Index As Integer)
   RaiseEvent LostFocus(Command1(Index).Tag)
End Sub

Private Sub DataCombo1_Change(Index As Integer)
Dim thisDataCombo As DataCombo

   If Me.Visible = False Then Exit Sub

   'selecciono el textbox
   Set thisDataCombo = Me.DataCombo1(Index)
   
   'salvo el nuevo valor
   SaveProperties thisDataCombo

   RaiseEvent Change(Index, DataCombo1(Index).Tag)
End Sub

Private Sub DataCombo1_LostFocus(Index As Integer)
   RaiseEvent LostFocus(DataCombo1(Index).Tag)
End Sub

Public Sub ShowDialog(Optional ByVal ShowModal As FormShowConstants = vbModal)
Dim ctrl       As Control
Dim IX         As Integer
Dim iOptionTrue As Integer

  'alinea los botones al primer contenedor
    
   On Error Resume Next
   
   For IX = 0 To Option1.UBound
      If Option1(IX).Value = True Then
         iOptionTrue = IX
      End If
      Option1_Click (IX)
   Next IX
   Option1_Click (iOptionTrue)
   
   ReDim aKeys(colControlsCanSave.Count - 1, 1)
   IX = 0
   For Each ctrl In colControlsCanSave
      aKeys(IX, 0) = "Dialogs\Defaults\" & CUsuario.Usuario & "\" & Me.Caption & "\" & ctrl.Tag
      IX = IX + 1
   Next ctrl
   
   'leo las claves
   GetKeyValues mvarControlData.Empresa, aKeys
   
   IX = 0
   For Each ctrl In colControlsCanSave
      vValue = aKeys(IX, 1)
      'me fijo si tiene ya definido un valor default
      If Not IsNull(vValue) And Not IsEmpty(vValue) Then
         Select Case TypeName(ctrl)
            Case "TextBox", "RichTextBox", "ComboBox", "PowerMask"
               ctrl.Text = vValue
            Case "CheckBox"
               ctrl.Value = IIf(vValue = "1", vbChecked, vbUnchecked)
            Case "OptionButton"
               ctrl.Value = CBool(vValue)
         End Select
      End If
      IX = IX + 1
   Next ctrl

   cmdPreview.Visible = Not mvarShowPrintButton = False
   cmdGuardar.Visible = Not mvarShowSaveButton = False
   cmdFiltros.Visible = Not mvarShowFilterButton = False
   cmdAvanzado.Visible = Not mvarShowAdvancedButton = False
   cmdCancel.Visible = Not mvarShowCancelButton = False
   cmdOK.Visible = Not mvarShowOkButton = False
   
   'Antes de mostrar los controles, dimensiono adecuadamente el form y lo centro
   Paint
   
   'el dialogo es siempre modal (y no es MDIChild)
   Me.Show vbModal
   
End Sub

Private Sub Form_Activate()
   
   SaveProperties

   RaiseEvent Activate
   
   'Bug# 5958 LGA - usamos de bandera  la propiedad SetDefaultEditData para no hacer setDefault si es edición
   'y no un alta en un frmDialog y evitamos que proponga los valores personalizados de objetos
      
   ' bSetDefault = SetDefaultEditData
    If mvarSetDefaultEditData Then Exit Sub  'por inc. SDP:82870 TP:6901
                                             'en vez de pisar la bandera del SetDefault lo hago salir antes en caso de que sea EDICION y no quiera hacer el SetDefaul
                                             'por Bug# 5958 LGA se setea en TRUE esta bandera cuando edita. En los casos que no EDITE sigue funcionando con la bandera bSetDefault como antes
    If bSetDefault Then Exit Sub
       
    If Not objControls Is Nothing Then
       objControls.SetDefaults
    End If
               
   bSetDefault = True

End Sub

Private Sub Form_Initialize()
   
   On Error GoTo GestErr

   Set mvarfrmFiltros = New frmFilter
   
   Set m_TextControls = New clsControlItems
   
   mvarShowPrintButton = False
   mvarShowSaveButton = False
   mvarShowFilterButton = False
   mvarShowAdvancedButton = False
   mvarShowCancelButton = True
   mvarShowOkButton = True
   
   Set frmHook = New AlgStdFunc.MsgHook
   Set tmr1 = New AlgStdFunc.clsTimer
   
   Set objControls = New clsControls

   bControlsPlaced = False

   Exit Sub

GestErr:
   LoadError ErrorLog, "Form_Initialize"
   ShowErrMsg ErrorLog
   
End Sub

Private Sub Form_Load()

   ' reseteo todos los valores
   On Error GoTo GestErr

   Set Values = Nothing
   Set ValuesFormatted = Nothing
   
   Me.Icon = LoadPicture(Icons & "Forms.ico")
   DoEvents
   
   Set objControls.Form = Me
   Set objControls.FormFind = frmFind
   Set objControls.Usuario = CUsuario
   
   DoEvents

   Exit Sub

GestErr:
   LoadError ErrorLog, "Form_Load"
   ShowErrMsg ErrorLog
End Sub

Private Sub Form_Terminate()

   Set mvarfrmFiltros = Nothing
   Set objControls = Nothing

   Set Values = Nothing
   Set ValuesFormatted = Nothing
   
   Set colFormat = Nothing
   Set colKeyPress = Nothing
   Set colSearch = Nothing
   Set colProperties = Nothing
   Set colDataType = Nothing
   Set colControlsDisabled = Nothing
   Set colControlsEnabled = Nothing
   Set colControlsCanSave = Nothing
   Set colFrames = Nothing
   
   Set m_TextControls = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If Not mvarfrmFiltros Is Nothing Then
      Unload mvarfrmFiltros
   End If
   
   Set mvarfrmFiltros = Nothing
   
   Set objControls = Nothing
   
   Unload frmFind
   
   '  Importante:
   '
   '  colControlsCanSave mantiene referencias a controles. Es necesario
   '  terminar esta colección para que el dialogo pueda descargarse correctamente.
   '  (caso contrario, el 29 de febrero del año del mongo nos daremos cuenta porque
   '  al finalizar la apliación compilada, se produce un Crash !)
   '
   Set colControlsCanSave = Nothing
      
   Set frmDialog = Nothing
   
End Sub

Private Sub frmHook_AfterMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, retValue As Long)
      
   '-- si el form administrar no es MRU, al cerrarlo le llega el mensaje WM_DESTROY
   '-- si el form administrar es MRU, al cerrarlo le llega el mensaje WM_SHOWWINDOW con parametro false
   
   If uMsg = WM_DESTROY Or (uMsg = WM_SHOWWINDOW And wParam = False) Then
      frmHook.StopSubclass hWnd
      tmr1.StartTimer 100
   End If

End Sub

Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComCtlLib.ColumnHeader)
      
'   Dim iOrdenActual  As Integer
'
'   iOrdenActual = ListView1(Index).SortOrder
'   RaiseEvent ColumnClick(ListView1(Index).Tag, ColumnHeader)
'
'   If iOrdenActual = ListView1(Index).SortOrder Then
'      ListView1(Index).SortOrder = IIf(ListView1(Index).SortOrder = lvwDescending, lvwAscending, lvwDescending)
'      ListView1(Index).SortKey = ColumnHeader.Index - 1
'
'      ListView1(Index).Sorted = True
'   End If
      
'Inc.SDP 91683
'Este metodo permite ordenar fechas, numeros, caracteres en un listView al hacer click en el encabezado.
'Nota: El tipo de dato que contiene la columna debe estar definido en la propiedad Tag del ColumnHeaders.
   Dim i As Long
   Dim Formato As String
   Dim strData() As String
   Dim columna As Long
        
   On Error Resume Next
   
   Const WM_SETREDRAW As Long = &HB&
    
   With Me.ListView1(Index)
    
        Call SendMessage(Me.hWnd, WM_SETREDRAW, 0&, 0&)
        
        columna = ColumnHeader.Index - 1
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Tipo de dato a ordenar
        ''''''''''''''''''''''''''''''''''''''''''''''
        Select Case UCase$(ColumnHeader.Tag)
        ' Fecha
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "DATE"
        
            Formato = "YYYYMMDDHhMmSs"
        
            ' Ordena alfabéticamente la columna con Fechas _
              ( es la columna que tiene en el tag el valor DATE )
        
            With .ListItems
              '  If (Columna > 0) Then ' no existe nada en el text paso a ser column_numero (siempre va a ser mayor que 0)
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(columna)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    Formato)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i

            End With
            
            ' Ordena alfabéticamente
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            With .ListItems
                'If (Columna > 0) Then ' no existe nada en el text paso a ser column_numero por lo tanto (siempre va a ser mayor que 0)
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(columna)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i

            End With
            
        ' Datos de numéricos
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "NUMBER" 'Or ""
        
            ' Ordena alfabéticamente la columna con números _
              ( es la columna que tiene en el tag el valor NUMBER )
        
            Formato = String(30, "0") & "." & String(30, "0")
                
            With .ListItems
                'If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(columna)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        Formato)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        Formato))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i

            End With
            
            ' Ordena alfabéticamente
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            With .ListItems
                'If (Columna > 0) Then
                    For i = 1 To .Count
                        With .Item(i).ListSubItems(columna)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i

            End With
        
        Case Else

            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
        End Select
      
    End With
    
    Call SendMessage(Me.hWnd, WM_SETREDRAW, 1&, 0&)
    Me.ListView1(Index).Refresh
   
End Sub
Private Function InvNumber(ByVal Number As String) As String
  Static i As Integer
  
  For i = 1 To Len(Number)
      Select Case Mid$(Number, i, 1)
      Case "-": Mid$(Number, i, 1) = " "
      Case "0": Mid$(Number, i, 1) = "9"
      Case "1": Mid$(Number, i, 1) = "8"
      Case "2": Mid$(Number, i, 1) = "7"
      Case "3": Mid$(Number, i, 1) = "6"
      Case "4": Mid$(Number, i, 1) = "5"
      Case "5": Mid$(Number, i, 1) = "4"
      Case "6": Mid$(Number, i, 1) = "3"
      Case "7": Mid$(Number, i, 1) = "2"
      Case "8": Mid$(Number, i, 1) = "1"
      Case "9": Mid$(Number, i, 1) = "0"
      End Select
  Next
  InvNumber = Number
End Function

Private Sub ListView1_DblClick(Index As Integer)
   RaiseEvent ItemDblClick(ListView1(Index).Tag)
End Sub

Private Sub ListView1_ItemCheck(Index As Integer, ByVal Item As MSComCtlLib.ListItem)
   RaiseEvent ItemCheck(ListView1(Index).Tag, Item)
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComCtlLib.ListItem)
   RaiseEvent ItemClick(ListView1(Index).Tag, Item)
End Sub

Private Sub ListView1_LostFocus(Index As Integer)
   RaiseEvent LostFocus(ListView1(Index).Tag)
End Sub

Private Sub ListView1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseDown(Index, Button, Shift, x, y)
End Sub

Private Sub ListView1_OLEDragDrop(Index As Integer, data As MSComCtlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent OLEDragDrop(Index, data, Effect, Button, Shift, x, y, ListView1(Index))
End Sub

Private Sub ListView1_OLEDragOver(Index As Integer, data As MSComCtlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
   RaiseEvent OLEDragOver(Index, data, Effect, Button, Shift, x, y, State, ListView1(Index))
End Sub

Private Sub ListView1_OLEStartDrag(Index As Integer, data As MSComCtlLib.DataObject, AllowedEffects As Long)
   RaiseEvent OLEStartDrag(Index, data, AllowedEffects, ListView1(Index))

End Sub

Private Sub m_TextControls_ItemChange(ByVal nIndex As Long, ByVal sKey As String)
Dim thisText As Control
   
   If Me.Visible = False Then Exit Sub
   
   ' obtengo la referencia al control que generó el evento
   Set thisText = m_TextControls.Item(nIndex).MyControl

   RaiseEvent Change(nIndex, thisText.Tag)

End Sub

Private Sub m_TextControls_ItemLostFocus(ByVal nIndex As Long, ByVal sKey As String)

Dim thisText         As Control
Dim aProp()          As Variant
Dim aArray()         As Variant
'Dim str              As String
Dim strExpresion     As String
Dim aControlsNames() As String
Dim aControls()      As Control
Dim aBoundField()    As String
Dim strFieldName     As String
Dim strTableName     As String
Dim IX               As Integer
Dim strCaption       As Variant
Dim dLookUpItem      As String
Dim sql              As String
Dim Rst              As ADODB.Recordset
Dim Sc               As AlgStdFunc.clsSQLScomposer
Dim aTableProperties As Variant
Dim strField         As String
Dim Valor1           As String


   'selecciono el textbox
   Set thisText = m_TextControls.Item(nIndex).MyControl

'   If Not objControls.IsChanged(thisText.hWnd) Then
'      'no cambio
'      RaiseEvent LostFocus(thisText.Tag)
'      Exit Sub
'   Else
      'salvo el nuevo valor
      SaveProperties thisText
'   End If

   On Error Resume Next
   aProp = colProperties(thisText.Tag)
   If Err.Number <> 0 Then
      ' el control es unbound
      Select Case FieldProperty(aProp, thisText.DataField, dsTipoDato)
         Case adNumeric, adDBTimeStamp
            thisText.Text = objControls.FormatControl(thisText)
      End Select
      RaiseEvent LostFocus(thisText.Tag)
      Exit Sub
   End If

   strCaption = NullString

   'si el arreglo esta lleno significa que el control es bound
   If Not IsArrayEmpty(aProp) Then

      'busco los datos para la búsqueda
      aArray = colSearch.Item(thisText.Tag)


      Select Case FieldProperty(aProp, thisText.DataField, dsTipoDato)
         Case adNumeric, adDBTimeStamp
            If objControls.GetControlInfo(thisText.hWnd).Formatear = FormatNo Then
            Else
               thisText.Text = objControls.FormatControl(thisText)   'me lo formateaba de prepo ?
            End If
      End Select


     If UBound(aArray) <= 3 Then    'no tiene una etiqueta asociada
        RaiseEvent LostFocus(thisText.Tag)
        Exit Sub
     End If


      IX = InStr(aArray(TABLE_FIELD), ".")

      strTableName = FieldProperty(aProp, aArray(TABLE_FIELD), dsTablaReferencia)
      strFieldName = FieldProperty(aProp, aArray(TABLE_FIELD), dsCampoReferencia)

      If Len(strTableName) = 0 Then
         strTableName = Left(aArray(TABLE_FIELD), IX - 1)
      End If

      If Len(strFieldName) = 0 Then
         strFieldName = Mid(aArray(TABLE_FIELD), IX + 1)
      End If


      aBoundField = Split(aArray(FIELD_LIST), ";")
      aControlsNames = Split(aArray(CONTROL_LIST), ";")

      'completo el arreglo aControls con los controles que entran en juego
      ReDim aControls(UBound(aControlsNames))
      For IX = LBound(aControlsNames) To UBound(aControlsNames)
         Set aControls(IX) = GetControl(aControlsNames(IX))
      Next IX

      'verifico si la búsqueda del dato que se va a mostrar en la etiqueta debe hacerce usando
      'dLookUp o bien la query del campo del diccionario. Si fue definido el parámetro strWhereDlookUp
      'significa que para obtener el valor de la etiqueta debo usar la query definida en el
      'diccionario y no la función dLookUp.

      dLookUpItem = colDlookUp.Item(thisText.Tag)
      If Len(dLookUpItem) > 0 Then

         'Uso la query del diccionario

         Set Sc = New AlgStdFunc.clsSQLScomposer
         Dim a1() As String
         Dim a2() As String
         Dim strWhere As String

         a1 = Split(dLookUpItem, ";")
         For IX = LBound(a1) To UBound(a1)
            a2 = Split(a1(IX), ",")

            strField = a2(0)
            aTableProperties = GetFieldInformation(strField)

            If a2(1) = "%1" Then
               Valor1 = thisText.Text
            Else
               Valor1 = a2(1)
            End If

            Select Case FieldProperty(aTableProperties, a2(0), dsTipoDato)
               Case adInteger, adDecimal, adNumeric, adSingle, adDouble
                  strWhere = strWhere & a2(0) & " = " & Valor1 & " AND "
               Case adDBTimeStamp
                  strWhere = strWhere & a2(0) & " = '" & Valor1 & "'AND "
               Case Else
                  ' si el tipo es alfanumerico es necesario agregar apostrofes
                  strWhere = strWhere & a2(0) & " = " & " '" & Valor1 & "' AND "
            End Select

         Next IX

         If Right(strWhere, 4) = "AND " Then
            strWhere = Left(strWhere, Len(strWhere) - 4)
         End If

         sql = FieldProperty(aProp, thisText.DataField, dsQuery)
         Sc.SQLInputString = sql
         Sc.SQLWhere = Sc.SQLWhere & " AND " & strWhere

         Set Rst = Fetch(mvarControlData.Empresa, Left(Sc.SQLOutputString, Len(Sc.SQLOutputString) - 1))

         If Not Rst.EOF Then
            strCaption = Rst(aArray(LABEL_FIELDNAME))
         End If

      Else

         'Uso la función dLookUp


         'genero la expresion para la funcion dLookUp()
         strExpresion = NullString
         For IX = LBound(aBoundField) To UBound(aBoundField)

            If Len(Trim(aControls(IX))) > 0 Then
               Select Case FieldProperty(aProp, thisText.DataField, dsTipoDato)
                  Case adNumeric
                     If Len(strExpresion) = 0 Then
'                        strExpresion = strExpresion & aBoundField(ix) & " = " & aControls(ix).Text
                        strExpresion = strExpresion & aBoundField(IX) & " = " & objControls.ControlValue(aControls(IX))
                     Else
'                        strExpresion = strExpresion & " AND " & aBoundField(ix) & " = " & aControls(ix).Text
                        strExpresion = strExpresion & " AND " & aBoundField(IX) & " = " & objControls.ControlValue(aControls(IX))
                     End If

                  Case adChar, adVarChar
                     If Len(strExpresion) = 0 Then
                        strExpresion = strExpresion & aBoundField(IX) & " = '" & aControls(IX).Text & "'"
                     Else
                        strExpresion = strExpresion & " AND " & aBoundField(IX) & " = '" & aControls(IX).Text & "'"
                     End If

               End Select
            End If
         Next IX

         If Len(strExpresion) > 0 Then
            strCaption = DLookUp(mvarControlData.Empresa, aArray(LABEL_FIELDNAME), strTableName, strExpresion)
         End If

      End If

      If IsNull(strCaption) Then
         aArray(LABEL_CONTROL).Caption = NullString
      Else
         aArray(LABEL_CONTROL).Caption = strCaption
      End If
      aArray(LABEL_CONTROL).ZOrder 0
   End If

   RaiseEvent LostFocus(thisText.Tag)

End Sub

Private Sub mnuContextItem_Click(Index As Integer)
   objControls.mnuContextItem_Click Index
End Sub

Private Sub objControls_Messages(ByVal lngMessage As Long, Info As Variant)
Dim hWndAdmin As Long
Dim ctrl As Control
Dim Key As String

   On Error GoTo GestErr
   
   Select Case lngMessage
      Case CTL_CALL_ADMIN
      
      
         '-- escondo la ventana modal, de esta manera puedo presentar otro form
         Me.Hide
         
         hWndAdmin = CallAdmin(Info, mvarControlData)
      
         If hWndAdmin = 0 Then
            '-- la apertura del form admin fallo
            Me.Show vbModal
            Exit Sub
         End If
            
         If Not frmHook.IsSubClassed(hWndAdmin) Then
            frmHook.StartSubclass hWndAdmin
         End If
      
   End Select
   
   'raiseo todos los mensajes de la clase objControls
   '(agrego al evento la clave del control que generó el mensaje)
   
   On Error Resume Next
   
   For Each ctrl In Me.Controls
      If ctrl.Tag = Info.ctrl.Tag Then
         If Err.Number = 0 Then
            Key = ctrl.Tag
            RaiseEvent Messages(Key, lngMessage, Info)
            Exit For
         End If
      End If
   Next ctrl
   
   Exit Sub
   
GestErr:
   LoadError ErrorLog, "frmDialog [objControls_Messages]"
   ShowErrMsg ErrorLog
End Sub

Private Sub Option1_Click(Index As Integer)
Dim IX               As Integer
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
         For IX = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(IX) = GetControl(aControlsNames(IX))
            aControls(IX).Enabled = Not (oControl.Value = True)
            
            If TypeOf aControls(IX) Is TextBox Or _
               TypeOf aControls(IX) Is RichTextBox Or _
               TypeOf aControls(IX) Is ComboBox Then
               Set aControls(IX) = GetControl("Label" & aControlsNames(IX))
               aControls(IX).Enabled = Not (oControl.Value = True)
            End If
            
         Next IX
   
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
         For IX = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(IX) = GetControl(aControlsNames(IX))
            aControls(IX).Enabled = (oControl.Value = True)
            
            If TypeOf aControls(IX) Is TextBox Or _
               TypeOf aControls(IX) Is RichTextBox Or _
               TypeOf aControls(IX) Is ComboBox Then
               Set aControls(IX) = GetControl("Label" & aControlsNames(IX))
               aControls(IX).Enabled = (oControl.Value = True)
            End If
            
         Next IX
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
         For IX = LBound(aControlsNames) To UBound(aControlsNames)
            Set aControls(IX) = GetControl(aControlsNames(IX))
            
            vValue = aControls(IX).Value
            aControls(IX).Value = Abs(Not (vValue))
            aControls(IX).Value = vValue
            
            If TypeOf aControls(IX) Is TextBox Or _
               TypeOf aControls(IX) Is RichTextBox Or _
               TypeOf aControls(IX) Is ComboBox Then
               Set aControls(IX) = GetControl("Label" & aControlsNames(IX))
               vValue = aControls(IX).Value
               aControls(IX).Value = Abs(Not (vValue))
               aControls(IX).Value = vValue
            End If
            
         Next IX
      End If
   End If
   
   Exit Sub

GestErr:
   LoadError ErrorLog, "Option1 (Click)"
   ShowErrMsg ErrorLog

End Sub

Private Sub Option1_LostFocus(Index As Integer)
   RaiseEvent LostFocus(Option1(Index).Tag)
End Sub

Private Sub RichTextBox1_Change(Index As Integer)
   If Me.Visible = False Then Exit Sub
   RaiseEvent Change(Index, RichTextBox1(Index).Tag)
End Sub

Private Sub RichTextBox1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyReturn And RichTextBox1(Index).LOCKED = True Then
       cmdOK_Click
   Else
       RaiseEvent KeyDown(RichTextBox1(Index).Tag, KeyCode, Shift)
   End If
   
End Sub

Private Sub RichTextBox1_LostFocus(Index As Integer)
   RaiseEvent LostFocus(RichTextBox1(Index).Tag)
End Sub

Public Function GetControl(ByVal Key As String) As Control
Dim cntl As Control

   On Error GoTo GestErr

   For Each cntl In Me.Controls
      If UCase(cntl.Tag) = UCase(Key) Then
         Set GetControl = cntl
         Exit For
      End If
   Next cntl

   Exit Function

GestErr:
   LoadError ErrorLog, "GetControl"
   ShowErrMsg ErrorLog

End Function
Public Property Let AcceptContinueEnabled(ByVal vData As Boolean)
    mvarAcceptContinueEnabled = vData
End Property
Public Property Get AcceptContinueEnabled() As Boolean
    AcceptContinueEnabled = mvarAcceptContinueEnabled
End Property
Public Property Let DialogName(ByVal vData As String)
    mvarDialogName = vData
End Property
Public Property Get DialogName() As String
    DialogName = mvarDialogName
End Property

Public Property Let CallerObjName(ByVal vData As String)
    mvarCallerObjName = vData
    
    'seteo el Caller Obj Name de la clase clsControls
    objControls.CallerObjName = vData
    
End Property
Public Property Get CallerObjName() As String
    CallerObjName = mvarCallerObjName
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
Public Property Let ShowCancelButton(ByVal vData As Boolean)
    mvarShowCancelButton = vData
    
    cmdCancel.Visible = mvarShowCancelButton
    
End Property
Public Property Let ShowOkButton(ByVal vData As Boolean)
    mvarShowOkButton = vData
    
    cmdOK.Visible = mvarShowOkButton
    If cmdOK.Visible = False Then
      cmdCancel.Move cmdOK.Left, cmdOK.Top
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
Public Property Let ShowAdvancedButton(ByVal vData As Boolean)
    mvarShowAdvancedButton = vData
    
    cmdAvanzado.Visible = mvarShowAdvancedButton
End Property

Public Sub AddFormulaToFilter(ByVal strCaption As String, Optional ByVal strConditionList As String, Optional ByVal strValueList As String)

   '-- permite agregar formulas al filtro
   
   On Error GoTo GestErr

   Select Case True
      Case Len(strConditionList) = 0 And Len(strValueList) = 0
         mvarfrmFiltros.Filter1.AddFormula strCaption
      Case Len(strConditionList) > 0 And Len(strValueList) = 0
         mvarfrmFiltros.Filter1.AddFormula strCaption, strConditionList
      Case Len(strConditionList) > 0 And Len(strValueList) > 0
         mvarfrmFiltros.Filter1.AddFormula strCaption, strConditionList, strValueList
   End Select

   Exit Sub

GestErr:
   LoadError ErrorLog, "AddFormulaToFilter"
   ShowErrMsg ErrorLog
   
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

Private Sub SaveProperties(Optional ctrlToUpdate As Control)
Dim ctrl As Control, ctrl2 As Control
Dim IX      As Integer
Dim j       As Integer
Dim itmx    As ListItem
Dim subI    As ListSubItem

   ' salvo las propiedades
    
   On Error Resume Next
   
   If ctrlToUpdate Is Nothing Then
   
      Set Values = New Collection
      Set ValuesFormatted = New Collection
      
      For Each ctrl In Controls
          If ctrl.Visible = True Then
          
            If ctrl.Tag <> "" Then
            
               ' si es un option button o check box
               ' almaceno el valor actual
               If TypeOf ctrl Is OptionButton Then
                   Values.Add ctrl.Value, ctrl.Tag
               ElseIf TypeOf ctrl Is CheckBox Then
                   Values.Add ctrl.Value, ctrl.Tag
               ElseIf TypeOf ctrl Is TextBox Then
                   Values.Add objControls.ControlValue(ctrl), ctrl.Tag
                   ValuesFormatted.Add ctrl.Text, ctrl.Tag
               ElseIf TypeOf ctrl Is PowerMask Then
                   Values.Add ctrl.Text, ctrl.Tag
                   ValuesFormatted.Add ctrl.Text, ctrl.Tag
               ElseIf TypeOf ctrl Is ComboBox Then
                   Values.Add Trim(ctrl.Text), ctrl.Tag
               ElseIf TypeOf ctrl Is DataCombo Then
                   Values.Add objControls.ControlValue(ctrl), ctrl.Tag
                   ValuesFormatted.Add ctrl.Text, ctrl.Tag
               ElseIf TypeOf ctrl Is RichTextBox Then
                   Values.Add Trim(ctrl.Text), ctrl.Tag
               ElseIf TypeOf ctrl Is ListView Then
                   IX = 0
                   For Each itmx In ctrl.ListItems
                   
                      IX = IX + 1
                      Values.Add Trim(itmx.Text), ctrl.Tag & ";Item" & IX & ";Text"
                      Values.Add itmx.Checked, ctrl.Tag & ";Item" & IX & ";Checked"
                      If itmx.Selected Then
                        Values.Add Trim(itmx.Text), ctrl.Tag & ";SelectedItem"
                      End If
                      
                      j = 0
                      For Each subI In itmx.ListSubItems
                         j = j + 1
                         Values.Add Trim(subI.Text), ctrl.Tag & ";Item" & IX & ";SubItem" & j
                      Next subI
                      
                   Next itmx
                    
               ElseIf TypeOf ctrl Is Frame Then
                   ' si es un frame,almaceno la key solo del
                   ' option button cuyo valor es true
                   For Each ctrl2 In Controls
                       If TypeOf ctrl2 Is OptionButton Then
                           If (ctrl2.Container Is ctrl) And ctrl2.Value = True Then
                               Values.Add ctrl2.Tag, ctrl.Tag
                               Exit For
                           End If
                       End If
                   Next
               End If
            End If
         End If
      Next
      
      Values.Add mvarfrmFiltros.Filter1.SQLWhere, "SQLWhere"
      Values.Add mvarfrmFiltros.Filter1.FilterWhere, "FilterWhere"
      Values.Add mvarfrmFiltros.Filter1.ListConditions, "ListConditions"
      Values.Add mvarfrmFiltros.Filter1.ArrayFilterList, "ArrayFilterList"
      
   Else
   
      For Each ctrl In Controls
          
         If ctrlToUpdate.Tag = ctrl.Tag Then
          
             Values.Remove ctrlToUpdate.Tag
          
             If TypeOf ctrl Is OptionButton Then
                 Values.Add ctrl.Value, ctrl.Tag
             ElseIf TypeOf ctrl Is CheckBox Then
                 Values.Add ctrl.Value, ctrl.Tag
             ElseIf TypeOf ctrl Is TextBox Then
                 Values.Add objControls.ControlValue(ctrl), ctrl.Tag
                 ValuesFormatted.Add ctrl.Text, ctrl.Tag
             ElseIf TypeOf ctrl Is PowerMask Then
                 Values.Add ctrl.Text, ctrl.Tag
             ElseIf TypeOf ctrl Is ComboBox Then
                 Values.Add Trim(ctrl.Text), ctrl.Tag
             ElseIf TypeOf ctrl Is DataCombo Then
                 Values.Add objControls.ControlValue(ctrl), ctrl.Tag
                 ValuesFormatted.Add ctrl.Text, ctrl.Tag
             ElseIf TypeOf ctrl Is RichTextBox Then
                 Values.Add Trim(ctrl.Text), ctrl.Tag
             ElseIf TypeOf ctrl Is ListView Then
                 IX = 0
                 For Each itmx In ctrl.ListItems
                 
                    IX = IX + 1
                    Values.Add Trim(itmx.Text), ctrl.Tag & ";Item" & IX & ";Text"
                    Values.Add itmx.Checked, ctrl.Tag & ";Item" & IX & ";Checked"
                    
                    j = 0
                    For Each subI In itmx.ListSubItems
                       j = j + 1
                       Values.Add Trim(subI.Text), ctrl.Tag & ";Item" & IX & ";SubItem" & j
                    Next subI
                    
                 Next itmx
                  
             ElseIf TypeOf ctrl Is Frame Then
                 ' si es un frame,almaceno la key solo del
                 ' option button cuyo valor es true
                 For Each ctrl2 In Controls
                     If TypeOf ctrl2 Is OptionButton Then
                         If (ctrl2.Container Is ctrl) And ctrl2.Value = True Then
                             Values.Add ctrl2.Tag, ctrl.Tag
                             Exit For
                         End If
                     End If
                 Next
             End If
             
         End If
         
      Next
   
   End If
   

End Sub

Public Function IsChanged(ByVal Key As String) As Boolean
Dim ctrl As Control

   Set ctrl = GetControl(Key)
   
   IsChanged = objControls.IsChanged(ctrl.hWnd)
   
End Function
Public Property Get ControlsClass() As clsControls
   Set ControlsClass = objControls
End Property

Public Property Let ControlData(ByVal vData As Variant)
    mvarControlData = vData
    
    MenuKey = vData.MenuKey 'tp 8367
    
    mvarfrmFiltros.ControlData = vData
    
   With ErrorLog
      .Form = Me.Name
      .Empresa = mvarControlData.Empresa
   End With
    
End Property

Public Property Get ControlData() As Variant
    ControlData = mvarControlData
End Property

Private Sub Paint()
Dim MinTop     As Integer
Dim ctrl       As Control
Dim thisFrame  As Frame
Dim frameIndex As Integer
Dim LeftFreButtons As Integer

   '/
   ' dimensiono adecuadamente el form, reubico el freButtons y centro el form
   '/
   
   On Error Resume Next
   
   MinTop = 32700
   LeftFreButtons = 0
   
   ' busco el primer control para darle el foco
   For Each ctrl In Controls
   
      If TypeName(ctrl) <> "Menu" And TypeName(ctrl) <> "MDIExtend" Then
          If ctrl.Index > 0 Then
            'descarto los Index = 0
            If ctrl.Container.Name = Me.Name Then
               If ctrl.Name <> "freButtons" Then
                  If ctrl.Top < MinTop Then
                     MinTop = ctrl.Top
                  End If
               End If
            End If
          End If
         
         If ctrl.TabIndex = 0 Then ctrl.SetFocus
      
      End If
      
      If ctrl.Name = "Frame1" Then
         'busco el frame externo mas ancho
         If ctrl.Container.Name = Me.Name Then
            If ctrl.Left + ctrl.Width > LeftFreButtons Then
               LeftFreButtons = ctrl.Left + ctrl.Width
            End If
         End If
      End If
   Next ctrl
   
   'posiciono el freButtons a la derecha del Frame mayor
   freButtons.Left = LeftFreButtons + 50
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
      If thisFrame.Container.Name <> "frmDynamicChild" Then
         Set thisFrame = thisFrame.Container
         thisFrame.Height = thisFrame.Height + FRAME_DISTANCE
      End If
   End If
   
   'si existe el boton guardar, lo muevo a la posicion Sup-DX
   If mvarShowSaveButton Then
      cmdGuardar.Move thisFrame.Width - cmdGuardar.Width - 100, 200
   End If
   
   'ajusto el ancho del form '***
   Width = freButtons.Left + freButtons.Width + 150
   
   'ajusto el alto del form
   Height = 600 + thisFrame.Top + thisFrame.Height
   
   If freButtons.Top + freButtons.Height > Height Then
      Height = freButtons.Top + freButtons.Height + 800
   End If

   If mvarShowOkButton = False And mvarShowCancelButton = False Then
      Me.Width = Me.Width - 1100
   End If

'   CenterForm Me
   
End Sub

Public Property Let MenuKey(ByVal vData As String)
   mvarMenuKey = EliminaAcentos(Left(vData, 50)) 'tp 7815
   'Para que me reinicie toda las propiedades
   Set objControls.Form = Me
End Property

Public Property Get MenuKey() As String
   MenuKey = mvarMenuKey
End Property
Public Sub RemoveControl(ByRef Key As String, _
                         Optional ByVal vTag As Variant = NullString)
'   Dim cntl As Control
'
'   For Each cntl In Me.Controls
'      If UCase(cntl.Tag) = UCase(Key) Then Exit For
'   Next cntl
'
'   Exit Sub
End Sub

Private Function EliminaAcentos(ByVal strWord As String) As String
Dim strAcentos As String
Dim Caracter   As String
Dim IX         As Integer

   '-- elimina las vocales acentuadas
   
   strAcentos = "áéíóúàèìòùÁÉÍÓÚÀÈÌÒÙ"
   
   For IX = 1 To Len(strWord)
      Caracter = Mid(strWord, IX, 1)
      If InStr(strAcentos, Caracter) > 0 Then
         Select Case Caracter
            Case "á", "à"
               strWord = Replace(strWord, Caracter, "a")
            Case "Á", "À"
               strWord = Replace(strWord, Caracter, "A")
            Case "é", "è"
               strWord = Replace(strWord, Caracter, "e")
            Case "É", "È"
               strWord = Replace(strWord, Caracter, "E")
            Case "í", "ì"
               strWord = Replace(strWord, Caracter, "i")
            Case "Í", "Ì"
               strWord = Replace(strWord, Caracter, "I")
            Case "ó", "ò"
               strWord = Replace(strWord, Caracter, "o")
            Case "Ó", "Ò"
               strWord = Replace(strWord, Caracter, "O")
            Case "ú", "ù"
               strWord = Replace(strWord, Caracter, "u")
            Case "Ú", "Ù"
               strWord = Replace(strWord, Caracter, "U")
         End Select
      End If
   Next IX
   
   EliminaAcentos = strWord
   
End Function
