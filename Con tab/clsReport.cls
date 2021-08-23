VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Event DblClickOnReport(ByVal X As Long, ByVal Y As Long, EventInfo As Variant, UseDefault As Boolean)
Event OnContextMenu(ByVal ObjectDescription As Variant, ByVal X As Long, ByVal Y As Long, UseDefault As Boolean)
Event Messages(ByVal lngMessage As Long, ByVal Info As Variant)
Event CantPaginasImpresas(ByVal lngCantPaginas As Long)

Private Type SubReport
   Report                  As CRAXDRT.Report
   Name                    As String
   Recordset               As ADODB.Recordset
   PrimaryLinks            As String
   RelatedLinks            As String
End Type

Private Type udtRecordsets
   adoRecordset            As ADODB.Recordset
   TableNumber             As Integer
End Type
Private ItemRecordsets     As udtRecordsets

Private mvarControlData    As DataShare.udtControlData
'Private WithEvents Reporte As frmReport

Private mvarComments       As Boolean
Private mvarReportFile     As CRAXDRT.Report
Private mvarRecordset      As ADODB.Recordset
Private mvarParameters     As New Collection
Private mvarPreviewMode    As Boolean
Private mvarCopies         As Integer
Private mvarCodigoEstructura  As String

Private ItemSubReport      As SubReport
Private mcolSubReport      As Collection
Private mcolRecordsets     As Collection

Private mvarBitMap         As String
Private mvarPreview        As Boolean
Private mvarCommentText    As String
Private mvarTitle          As String
Private mvarShowUserInfo   As Boolean
Private mvarZoom           As Integer
Private mvarSuprimeSubReport As Boolean

'Public Sub AddSubReport(ByVal SubReportName As String, ByVal SubReportRecordset As ADODB.Recordset)
'Dim aSplit() As String
'Dim aSplitFS() As String
'Dim aSplitOrd() As String
'Dim ix As Integer
'Dim iY As Integer
'
'   If Len(SubReportName) = 0 Then Exit Sub
'
'   With ItemSubReport
'      .Name = SubReportName
'      Set .Recordset = SubReportRecordset
'      Set .Report = mvarReportFile.OpenSubreport(SubReportName)
'
'      If InStr(.Report.RecordSelectionFormula, "ado") > 0 Then
'         aSplit = Split(Replace(Replace(Replace(Replace(.Report.RecordSelectionFormula, "?Pm-ado.", ""), "ado.", ""), "{", ""), "}", ""), "and")
'         .PrimaryLinks = NullString
'         .RelatedLinks = NullString
'
'         If mcolSubReport.Count > 0 Then
'            ' La propiedad  RecordSelectionFormula a veces no devuelve la where en el
'            ' orden correcto, por eso si ya hay un SubReport en la coleccion,
'            ' hago que el orden de los PrimaryLinks Coincida, por lo menos
'            ' que queden todos en el mismo orden que el primero
'
'            aSplitFS = Split(mcolSubReport(1).PrimaryLinks, ";")
'
'            ReDim aSplitOrd(UBound(aSplitFS))
'
'            For ix = 0 To UBound(aSplitFS)
'               For iY = 0 To UBound(aSplit)
'                  If InStr(aSplit(iY), aSplitFS(ix)) > 0 Then
'                     aSplitOrd(ix) = aSplit(iY)
'                     Exit For
'                  End If
'               Next iY
'            Next ix
'
'            aSplit = aSplitOrd
'
'         End If
'
'         For ix = 0 To UBound(aSplit)
'           .PrimaryLinks = .PrimaryLinks & Trim(Mid(aSplit(ix), InStr(aSplit(ix), "=") + 1)) & ";"
'           .RelatedLinks = .RelatedLinks & Trim(Left(aSplit(ix), InStr(aSplit(ix), "=") - 1)) & ";"
'         Next ix
'         If Right(.PrimaryLinks, 1) = ";" Then .PrimaryLinks = Left(.PrimaryLinks, Len(.PrimaryLinks) - 1)
'         If Right(.RelatedLinks, 1) = ";" Then .RelatedLinks = Left(.RelatedLinks, Len(.RelatedLinks) - 1)
'
'         .PrimaryLinks = Replace(.PrimaryLinks, vbCrLf, "")
'         .RelatedLinks = Replace(.RelatedLinks, vbCrLf, "")
'
'         .PrimaryLinks = Trim(.PrimaryLinks)
'         .RelatedLinks = Trim(.RelatedLinks)
'
'      End If
'   End With
'
'   mcolSubReport.Add ItemSubReport, SubReportName
'
'End Sub
'Public Sub AddRecordset(ByVal rst As ADODB.Recordset, ByVal iTableNumber As Integer)
'
'   If iTableNumber < 2 Then Exit Sub
'
'   With ItemRecordsets
'      Set .adoRecordset = rst
'      .TableNumber = iTableNumber
'   End With
'
'   mcolRecordsets.Add ItemRecordsets, CStr(iTableNumber)
'
'End Sub
Public Property Set ReportForm(ByVal vData As Object)
    Set Reporte = vData
End Property
Public Property Get ReportForm() As Object
'    Set ReportForm = Reporte
End Property

Public Property Set SubReports(ByVal vData As Collection)
    Set mcolSubReport = vData
End Property
Public Property Get SubReports() As Collection
    Set SubReports = mcolSubReport
End Property
Public Property Set colRecordsets(ByVal vData As Collection)
    Set mcolRecordsets = vData
End Property
Public Property Get colRecordsets() As Collection
    Set colRecordsets = mcolRecordsets
End Property
Public Property Let CommentText(ByVal vData As String)
    mvarCommentText = vData
End Property

Public Property Get CommentText() As String
    CommentText = mvarCommentText
End Property

Public Property Let Copies(ByVal vData As Integer)
    mvarCopies = vData
End Property

Public Property Get Copies() As Integer
    Copies = mvarCopies
End Property

Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
End Property

Public Property Get Title() As String
    Title = mvarTitle
End Property

Public Property Let Preview(ByVal vData As Boolean)
    mvarPreview = vData
End Property

Public Property Get Preview() As Boolean
    Preview = mvarPreview
End Property

Public Property Let PreviewMode(ByVal vData As Boolean)
    mvarPreviewMode = vData
End Property

Public Property Get PreviewMode() As Boolean
    PreviewMode = mvarPreviewMode
End Property

Public Property Let ShowUserInfo(ByVal vData As Boolean)
    mvarShowUserInfo = vData
End Property

Public Property Get ShowUserInfo() As Boolean
    ShowUserInfo = mvarShowUserInfo
End Property
Public Property Let SuprimeSubReport(ByVal vData As Boolean)
    mvarSuprimeSubReport = vData
End Property
Public Property Get SuprimeSubReport() As Boolean
    SuprimeSubReport = mvarSuprimeSubReport
End Property
Public Property Let BitMap(ByVal vData As String)
   '
   '  estaba asterisquiada por eso no imprimia (estar atento a posibles consecuencias !)
   '
    mvarBitMap = vData
End Property
Public Property Get BitMap() As String
    BitMap = mvarBitMap
End Property

Public Property Set Recordset(ByVal vData As ADODB.Recordset)
   Set mvarRecordset = vData
End Property

Public Property Get Recordset() As ADODB.Recordset
   Set Recordset = mvarRecordset
End Property

Public Property Set Report(ByVal vData As CRAXDRT.Report)
    Set mvarReportFile = vData
End Property

Public Property Get Report() As CRAXDRT.Report
   Set Report = mvarReportFile
End Property

Public Property Let Comments(ByVal vData As Boolean)
    mvarComments = vData
End Property

Public Property Get Comments() As Boolean
    Comments = mvarComments
End Property
Public Property Let CodigoEstructura(ByVal vData As String)
    mvarCodigoEstructura = vData
End Property

Public Property Get CodigoEstructura() As String
    CodigoEstructura = mvarCodigoEstructura
End Property
Public Sub ShowReport(Optional ByVal objImpresora As BOGeneral.clsImpresoras, Optional ByVal bSoloExporta As Boolean = False)
Dim ix As Integer

   On Error Resume Next

   If Not mvarRecordset Is Nothing Then
      'el report posee datos
      mvarReportFile.DiscardSavedData
      mvarReportFile.Database.SetDataSource mvarRecordset, 3, 1
      Set Reporte.RecordsetPrint = mvarRecordset
   Else
      'Continuo. El report puede estar compuesto solo por subreports --> Falla la exportacion Excell
   End If
   
   For ix = 1 To mcolRecordsets.Count
      mvarReportFile.Database.SetDataSource mcolRecordsets(ix).adoRecordset, 3, mcolRecordsets(ix).TableNumber
   Next ix
   
   For ix = 1 To mcolSubReport.Count
      mcolSubReport(ix).Report.DiscardSavedData
      mcolSubReport(ix).Report.Database.SetDataSource mcolSubReport(ix).Recordset, 3, 1
   Next ix
   
   'comentarios
   mvarReportFile.txtComments.Suppress = Not mvarComments
   Reporte.CommentText = mvarCommentText
   Reporte.CodigoEstructura = mvarCodigoEstructura
   
   'Usuario
   mvarReportFile.Sections("UserInfo").Suppress = Not mvarShowUserInfo
   If mvarShowUserInfo = True Then
      mvarReportFile.txtUsuario.SetText CUsuario.NombreUsuario
   Else
      mvarReportFile.txtUsuario.SetText NullString
   End If
   
   'referencia logo
   If Len(mvarBitMap) > 0 Then
      Reporte.BitMap = BitMaps & mvarBitMap
   End If
   
   'titulo de la vista previa
   Reporte.Caption = IIf(Len(mvarTitle) = 0, "Vista Previa", "Vista Previa de " & mvarTitle)
   
   Reporte.Zoom = mvarZoom
   
   Reporte.ControlData = mvarControlData
   
   Set Reporte.Report = mvarReportFile
   Set Reporte.Parameters = mvarParameters
   
   Reporte.PreviewMode = mvarPreviewMode
   
   If mvarCopies <> 0 Then Reporte.txtCopias = mvarCopies
   
   ' le paso al frmReport una referencia a esta instancia
   Set Reporte.MyClass = Me
   
   If bSoloExporta Then
      'Esto se tiene que ejecutar si quiero exportarPDF y no imprimir (comp. retencion LGA)
      mvarReportFile.PictureReference.SetText Reporte.BitMap
      mvarReportFile.txtComments.SetText mvarCommentText
   Else
      Reporte.ShowReport mvarPreview, objImpresora
   End If
   
End Sub

Public Sub ExportToPDF(ByVal PathFile As String)
    Reporte.ExportToPDF PathFile
End Sub
Public Sub PapelPersonalizado(ByVal strPapel As String)
    Reporte.strPapel = strPapel ' Implementar hoja legal Inc. 43102 (debe haber algo mas facil pero no lo encuentro a esta hora)
End Sub

Private Sub Class_Initialize()
   mvarShowUserInfo = True
   mvarSuprimeSubReport = False
      
'   Set Reporte = New frmReport
   
   Set mcolSubReport = New Collection
   Set mcolRecordsets = New Collection
   
'   mvarZoom = SystemOptions.iZoom
   
End Sub

Private Sub Class_Terminate()
   
   If Not ItemSubReport.Report Is Nothing Then
      Set ItemSubReport.Report = Nothing
   End If
   
'   Unload Reporte
'   Set Reporte = Nothing
   
   Set mcolSubReport = Nothing
   Set mvarReportFile = Nothing
   Set mvarParameters = Nothing
   Set mcolSubReport = Nothing
   
End Sub

Public Sub AddParameter(ByVal strName As String, ByVal vValue As Variant)
Dim ix   As Integer
Dim i    As Integer

   If Len(strName) = 0 Then
      Err.Raise vbObjectError + 1, "clsReport", "El nombre del parámetro no puede ser nulo"
      Exit Sub
   End If
   
   i = 0
   For ix = 1 To mvarReportFile.ParameterFields.Count
      If UCase(strName) = UCase(mvarReportFile.ParameterFields(ix).ParameterFieldName) Then
         i = ix
      End If
   Next
   
   If i = 0 Then
      Err.Raise vbObjectError + 1, "clsReport", "No existe el parámetro " & strName & " en el Report"
      Exit Sub
   End If
   
   mvarParameters.Add vValue, strName
   
End Sub

Public Property Let ControlData(ByRef vData As DataShare.udtControlData)
   mvarControlData = vData
   
   Dim aKeysToRead()

   ReDim aKeysToRead(0, 1)

   aKeysToRead(0, 0) = "Opciones\Logo Empresa;Logo.bmp"

   GetKeyValues mvarControlData.Empresa, aKeysToRead

   mvarBitMap = IIf(IsNull(aKeysToRead(0, 1)), "Logo.bmp", aKeysToRead(0, 1))
   
End Property

Public Property Get ControlData() As DataShare.udtControlData
   ControlData = mvarControlData
End Property

Public Property Let Zoom(ByVal vData As Integer)
    mvarZoom = vData
End Property

Public Property Get Zoom() As Integer
    Zoom = mvarZoom
End Property

Private Sub Reporte_DblClickOnReport(ByVal X As Long, ByVal Y As Long, EventInfo As Variant, UseDefault As Boolean)
   RaiseEvent DblClickOnReport(X, Y, EventInfo, UseDefault)
End Sub

Private Sub Reporte_Messages(ByVal lngMessage As Long, ByVal Info As Variant)
   RaiseEvent Messages(lngMessage, Info)
End Sub

Private Sub Reporte_OnContextMenu(ByVal ObjectDescription As Variant, ByVal X As Long, ByVal Y As Long, UseDefault As Boolean)
   RaiseEvent OnContextMenu(ObjectDescription, X, Y, UseDefault)
End Sub

Private Sub Reporte_CantPaginasImpresas(ByVal lngCantPaginas As Long)
   RaiseEvent CantPaginasImpresas(lngCantPaginas)
End Sub
