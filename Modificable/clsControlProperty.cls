VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarUser As String
Private mvarControlType As String
Private mvarCurrentDefaultValue As Variant
Private mvarCurrentEdicion As String
Private mvarCurrentForeColor As Long
Private mvarCurrentBackColor As Long
Private mvarCurrentFont As String
Private mvarCurrentFontName As String
Private mvarCurrentFontSize As Single
Private mvarCurrentFontBold As String
Private mvarCurrentFontItalic As String
Private mvarCurrentFontStrikethrough As String
Private mvarCurrentFontUnderline As String
Private mvarUseDefaultValues As String

Public Property Let UseDefaultValues(ByVal vData As String)
    mvarUseDefaultValues = vData
    
    If mvarUseDefaultValues = si Then
    
      mvarCurrentForeColor = BlackColor
      Select Case mvarControlType
         Case "OptionButton", "CheckBox"
            mvarCurrentBackColor = &H8000000F   'gris
         Case Else
            mvarCurrentBackColor = &H80000005   'blanco
      End Select
      mvarCurrentFontName = "MS Sans Serif"
      mvarCurrentFontSize = 8
      mvarCurrentFontBold = No
      mvarCurrentFontItalic = No
      mvarCurrentFontStrikethrough = No
      mvarCurrentFontUnderline = No
    
      mvarCurrentFont = mvarCurrentFontName & ";" & _
                        IIf(mvarCurrentFontBold = si, "Bold", NullString) & ";" & _
                        mvarCurrentFontSize & ";" & _
                        IIf(mvarCurrentFontItalic = si, "Italic", NullString) & ";" & _
                        IIf(mvarCurrentFontStrikethrough = si, "Tachado", NullString) & ";" & _
                        IIf(mvarCurrentFontUnderline = si, "Subrayado", NullString)
    
    End If
    
End Property
Public Property Get UseDefaultValues() As String
    UseDefaultValues = mvarUseDefaultValues
End Property

Public Property Let User(ByVal vData As String)
    mvarUser = vData
End Property
Public Property Get User() As String
    User = mvarUser
End Property

Public Property Let ControlType(ByVal vData As String)
    mvarControlType = vData
End Property
Public Property Get ControlType() As String
    ControlType = mvarControlType
End Property

Public Property Let CurrentFontUnderline(ByVal vData As String)
    mvarCurrentFontUnderline = vData
End Property
Public Property Get CurrentFontUnderline() As String
    CurrentFontUnderline = mvarCurrentFontUnderline
End Property
Public Property Let CurrentFontStrikethrough(ByVal vData As String)
    mvarCurrentFontStrikethrough = vData
End Property
Public Property Get CurrentFontStrikethrough() As String
    CurrentFontStrikethrough = mvarCurrentFontStrikethrough
End Property
Public Property Let CurrentFontItalic(ByVal vData As String)
    mvarCurrentFontItalic = vData
End Property
Public Property Get CurrentFontItalic() As String
    CurrentFontItalic = mvarCurrentFontItalic
End Property
Public Property Let CurrentFontBold(ByVal vData As String)
    mvarCurrentFontBold = vData
End Property
Public Property Get CurrentFontBold() As String
    CurrentFontBold = mvarCurrentFontBold
End Property
Public Property Let CurrentFontSize(ByVal vData As Single)
    mvarCurrentFontSize = vData
End Property
Public Property Get CurrentFontSize() As Single
    CurrentFontSize = mvarCurrentFontSize
End Property
Public Property Let CurrentFont(ByVal vData As String)
Dim aFont() As String

   mvarCurrentFont = vData
    
   aFont = Split(mvarCurrentFont, ";")
   mvarCurrentFontName = aFont(0)
   mvarCurrentFontSize = aFont(1)
   mvarCurrentFontBold = IIf(aFont(2) = si, True, False)
   mvarCurrentFontItalic = IIf(aFont(3) = si, True, False)
   mvarCurrentFontStrikethrough = IIf(aFont(4) = si, True, False)
   mvarCurrentFontUnderline = IIf(aFont(5) = si, True, False)
    
    
End Property
Public Property Get CurrentFont() As String
    CurrentFont = mvarCurrentFont
End Property
Public Property Let CurrentFontName(ByVal vData As String)
    mvarCurrentFontName = vData
End Property
Public Property Get CurrentFontName() As String
    CurrentFontName = mvarCurrentFontName
End Property
Public Property Let CurrentBackColor(ByVal vData As Long)
    mvarCurrentBackColor = vData
End Property
Public Property Get CurrentBackColor() As Long
    CurrentBackColor = mvarCurrentBackColor
End Property
Public Property Let CurrentForeColor(ByVal vData As Long)
    mvarCurrentForeColor = vData
End Property
Public Property Get CurrentForeColor() As Long
    CurrentForeColor = mvarCurrentForeColor
End Property
Public Property Let CurrentEdicion(ByVal vData As String)
    mvarCurrentEdicion = vData
End Property
Public Property Get CurrentEdicion() As String
    CurrentEdicion = mvarCurrentEdicion
End Property
Public Property Let CurrentDefaultValue(ByVal vData As Variant)
    mvarCurrentDefaultValue = vData
End Property
Public Property Get CurrentDefaultValue() As Variant
   CurrentDefaultValue = mvarCurrentDefaultValue
End Property

