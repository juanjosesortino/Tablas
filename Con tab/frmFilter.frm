VERSION 5.00
Object = "{B97E3E11-CC61-11D3-95C0-00C0F0161F05}#172.0#0"; "ALGControls.ocx"
Begin VB.Form frmFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtros"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7455
   Begin ALGControls.Filter Filter1 
      Height          =   405
      Left            =   450
      TabIndex        =   0
      Top             =   165
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmHook As AlgStdFunc.MsgHook
Attribute frmHook.VB_VarHelpID = -1
Private WithEvents tmr1    As AlgStdFunc.clsTimer
Attribute tmr1.VB_VarHelpID = -1

Private Const WM_DESTROY = &H2                          'mensaje que todas las windows reciben antes de ser cerradas
Private Const WM_SHOWWINDOW = &H18

Private mvarEmpresa        As String                    'código de la empresa

'-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
'  Siempre que el usuario elija la opciòn Administrar (del menu contextual) dentro del
'  control Filter, este enviara al frmFilter un mensaje a traves del evento Filter1_Messages,
'  informando cual es el form que debe ser abierto para administar. Una vez cerrado dicho
'  form, se vuelve a presentar el frmFilter que previamente fue ocultado para permitir que
'  el form administrar se abierto (frmFilter se presenta siempre de manera modal).
'-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --

Private Sub Form_Initialize()
   Set frmHook = New AlgStdFunc.MsgHook
   Set tmr1 = New AlgStdFunc.clsTimer
End Sub

Private Sub frmHook_AfterMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, retValue As Long)
      
   '-- si el form administrar no es MRU, al cerrarlo le llega el mensaje WM_DESTROY
   '-- si el form administrar es MRU, al cerrarlo le llega el mensaje WM_SHOWWINDOW con parametro false
   
   If uMsg = WM_DESTROY Or (uMsg = WM_SHOWWINDOW And wParam = False) Then
      frmHook.StopSubclass hWnd
      tmr1.StartTimer 100
   End If

End Sub

Private Sub Filter1_Messages(ByVal lngMessage As Long, ByRef Info As Variant)
'Dim hWndAdmin As Long
'
'   On Error GoTo GestErr
'
'   Select Case lngMessage
'      Case FILTER_CALL_ADMIN
'
'         Dim sFormAdmin As String
'         Dim sModuloAdmin As String
'
'         sFormAdmin = Left(Info, InStr(Info, ";") - 1)
'         sModuloAdmin = Mid(Info, InStr(Info, ";") + 1)
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
'            hWndAdmin = ShowForm(sFormAdmin, Filter1.Empresa)
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
'            objEXE.OpenForm sFormAdmin, Filter1.Empresa
'
'         End If
'
'
'      Case FILTER_QUERY_USER
'
'         Set Info = CUsuario
'
'      Case FILTER_QUERY_EMPRESA
'         Info = Empresa
'
'   End Select
'
'   Exit Sub
'
'GestErr:
''   LoadError "frmFilter"
''   ShowErrMsg
End Sub

Public Property Let Empresa(ByVal vData As String)
    mvarEmpresa = vData
End Property

Public Property Get Empresa() As String
    Empresa = mvarEmpresa
End Property

Private Sub tmr1_Timer()
   '-- vuelvo a presentar el frmFilter modal
   tmr1.StopTimer
   Me.Show vbModal
End Sub

Private Sub Filter1_Unload()
   Me.Hide
End Sub

Private Sub Form_Load()
Const distoriz = 0
Const distvert = 0
  
   Me.Top = 1000
   Me.Left = 1000
   
   With Filter1
      .Left = distoriz
      .Top = distvert
   End With
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   If UnloadMode = vbFormControlMenu Then
      Cancel = True
      Me.Hide
   End If
   
End Sub

Private Sub Form_Resize()
Static bResizing As Boolean

   If bResizing Then Exit Sub

   bResizing = True
   
   If Me.Width < 10000 Then Me.Width = 10000
   If Me.Height < 5000 Then Me.Height = 5000
   
   bResizing = False
   
   With Filter1
     .Width = Me.ScaleWidth
     .Height = Me.ScaleHeight
   End With

   
End Sub

