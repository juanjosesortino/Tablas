VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : clsControlItem
' DateTime  : 28/11/2002 11:40
' Author    : Algoritmo
' Purpose   : Esta clase, junto con la clase clsControlItems permiten implementar
'             una colección de controles con eventos. En nuestro caso, solo meteremos
'             en dicha colección controles de tipo TextBox y PowerMask de manera tal
'             de detectar algunos de sus eventos cuyo subclassing genera inconvenientes
'             durante el debug.
'---------------------------------------------------------------------------------------
Option Explicit

Private WithEvents objText As TextBox
Attribute objText.VB_VarHelpID = -1
Private WithEvents objPmk As PowerMask
Attribute objPmk.VB_VarHelpID = -1

Private m_sKey As String
Private m_bCancel As Boolean

Private m_lpControlObject As Long         ' puntero ilegal al objeto pariente

'---------------------------------------------------------------------------------------
' Procedure : MyControl
' DateTime  : 28/11/2002 11:50
' Author    : Algoritmo
' Purpose   : referencia al control de la coleccíon (TextBox / PowerMask)
'---------------------------------------------------------------------------------------
'
Public Property Set MyControl(ByVal vData As Object)
   If TypeName(vData) = "TextBox" Then
      Set objText = vData
   Else
      Set objPmk = vData
   End If
End Property

Public Property Get MyControl()
   If Not objText Is Nothing Then
      Set MyControl = objText
   Else
      Set MyControl = objPmk
   End If
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' DateTime  : 28/11/2002 11:51
' Author    : Algoritmo
' Purpose   : propiedad Name (R/O )
'---------------------------------------------------------------------------------------
'
Public Property Get Name() As String
   If objText Is Nothing Then
      Name = objPmk.Name
   Else
      Name = objText.Name
   End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : hWnd
' DateTime  : 28/11/2002 11:51
' Author    : Algoritmo
' Purpose   : Hwnd del control
'---------------------------------------------------------------------------------------
'
Public Property Get hWnd() As Long
   If objText Is Nothing Then
      hWnd = objPmk.hWnd
   Else
      hWnd = objText.hWnd
   End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : Key
' DateTime  : 28/11/2002 11:52
' Author    : Algoritmo
' Purpose   : Clave
'---------------------------------------------------------------------------------------
'
Public Property Get Key() As String
  Key = m_sKey
End Property
Public Property Let Key(ByVal sNewVal$)
  m_sKey = sNewVal
End Property

'---------------------------------------------------------------------------------------
' Procedure : Cancel
' DateTime  : 28/11/2002 11:52
' Author    : Algoritmo
' Purpose   : propiedad que permite de cancelar el evento Validate
'---------------------------------------------------------------------------------------
'
Public Property Let Cancel(ByVal Cancel As Boolean)
  m_bCancel = Cancel
End Property

'---------------------------------------------------------------------------------------
' Procedure : CollectionPointer
' DateTime  : 28/11/2002 11:54
' Author    : Algoritmo
' Purpose   : propiedad W/O usada para pasar el puntero a la colección
'             cuando este objeto es creado y agregado a la colección
'---------------------------------------------------------------------------------------
'
Friend Property Let CollectionPointer(ByVal nNewVal&)
  If m_lpControlObject Then
    ' el puntero debería ser seteado solo una vez. Si esto no ocurre raiseo un error
    ' si m_lpControlObject ya tiene ese puntero
    Err.Raise &H80000700, , "El puntero ha sido inicilizado"
  Else
    m_lpControlObject = nNewVal
  End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : objPmk_GotFocus
' DateTime  : 28/11/2002 11:57
' Author    : Algoritmo
' Purpose   : evento de los controles que quiero raisear
'---------------------------------------------------------------------------------------
'
Private Sub objPmk_GotFocus()
  CallEvent cf_ItemGotFocus
End Sub

Private Sub objPmk_LostFocus()
  CallEvent cf_ItemLostFocus
End Sub

Private Sub objPmk_SetText()
  CallEvent cf_ItemSetText
End Sub

Private Sub objPmk_Validate(Cancel As Boolean)
  
  CallEvent cf_ItemValidate
  
  Cancel = m_bCancel
  
End Sub

Private Sub objPmk_Change()
   CallEvent cf_ItemChange
End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : objText_GotFocus
' DateTime  : 28/11/2002 11:58
' Author    : Algoritmo
' Purpose   : paso el texto del evento al objeto colección
'---------------------------------------------------------------------------------------
'
Private Sub objText_GotFocus()
  CallEvent cf_ItemGotFocus
End Sub

Private Sub objText_LostFocus()
  CallEvent cf_ItemLostFocus
End Sub
Private Sub objText_Change()
   CallEvent cf_ItemChange
End Sub
Private Sub objText_Validate(Cancel As Boolean)
Dim bCancel As Boolean

  bCancel = Cancel
  
  CallEvent cf_ItemValidate
  
  If m_bCancel Then
     Cancel = m_bCancel
  Else
     Cancel = bCancel
  End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CallEvent
' DateTime  : 28/11/2002 11:59
' Author    : Algoritmo
' Purpose   : compruebo para asegurarme que el puntero sigue siendo válido
'             si es válido, utilizamos la función ResolvePointer para convertir
'             la referencia ilegal en una referencia legal al objeto de la coleccíon.
'             y llamo al método FireEvent pasandole cual es el evento que deseo raisear
'---------------------------------------------------------------------------------------
'
Private Sub CallEvent(ByVal nEvent As EVENT_INDEXES)

  If MValidatePointer.IsPointerValid(m_lpControlObject) > (-1) Then
    ResolvePointer(m_lpControlObject).FireEvent nEvent, m_sKey, ObjPtr(Me)
  End If

End Sub
' **************************************************

'---------------------------------------------------------------------------------------
' Procedure : ResolvePointer
' DateTime  : 28/11/2002 12:03
' Author    : Algoritmo
' Purpose   : transforma el puntero en una referencia valida.  No terminar el projecto
'             jamas en este punto a menos que desee generar un hermoso GPF
'---------------------------------------------------------------------------------------
'
Private Function ResolvePointer(ByVal lpObj&) As clsControlItems
  
  Dim oCF As clsControlItems
  
  CopyMemory oCF, lpObj, 4&
  Set ResolvePointer = oCF
  CopyMemory oCF, 0&, 4&

End Function

