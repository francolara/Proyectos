Attribute VB_Name = "mdlReportes"

Public Sub mostrarReporte(ByVal strNombreRpt As String, ByVal strNombreParametros As String, ByVal strValoresParametros As String, ByVal strTituloReporte As String, ByRef StrMsgError As String)
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim pReport         As ParameterFieldDefinition
Dim crxTable        As CRAXDRT.DatabaseTable
Dim tables As CRAXDRT.DatabaseTables
Dim arrNomParam()   As String
Dim arrValParam()   As String
Dim p As Integer
Dim CuentaTabla     As Integer
Dim rsReporte       As New ADODB.Recordset
Dim procedureName   As String
Dim database As CRAXDRT.database
On Error GoTo Err

Set reporte = aplicacion.OpenReport(gStrRutaRpts & strNombreRpt)

'    Set reporte = aplicacion.OpenReport(gStrRutaRpts & strNombreRpt)
'    arrNomParam = Split(strNombreParametros, "|")
'    arrValParam = Split(strValoresParametros, "|")
    
'    For Each crxTable In reporte.database.tables
'        crxTable.ConnectionProperties.item("Data Source") = gbservidor
'        crxTable.ConnectionProperties.item("Initial Catalog") = gbDatabase
'        crxTable.ConnectionProperties.item("User ID") = gbusuario
'        crxTable.ConnectionProperties.item("Password") = gbPassword
'        crxTable.ConnectionProperties.item("Integrated Security") = False
'    Next
'
'    For p = 0 To reporte.ParameterFields.Count - 1
'        Set pReport = reporte.ParameterFields.GetItemByName("@" & arrNomParam(p))
'        If pReport.ValueType = crNumberField Then
'            pReport.AddCurrentValue (Val("" & arrValParam(p)))
'        Else
'            pReport.AddCurrentValue ("" & arrValParam(p) & "")
'        End If
'    Next

Set database = reporte.database
Set tables = database.tables

For Each crxTable In tables
    procedureName = crxTable.Location
Next

procedureName = left(procedureName, (InStr(procedureName, ";")) - 1)
    
If procedureName = "" Then Screen.MousePointer = 0: Exit Sub
Set rsReporte = DataProcedimiento_Rpt(procedureName, StrMsgError, strValoresParametros)
If StrMsgError <> "" Then GoTo Err

If Not rsReporte.EOF And Not rsReporte.BOF Then
     reporte.database.SetDataSource rsReporte, 3
     vistaPrevia.CRViewer91.ReportSource = reporte
     vistaPrevia.Caption = strTituloReporte
     vistaPrevia.CRViewer91.ViewReport
     vistaPrevia.CRViewer91.DisplayGroupTree = True
     Screen.MousePointer = 0
     vistaPrevia.WindowState = 2
     vistaPrevia.Show
Else
    Screen.MousePointer = 0
    MsgBox "No existen Registros", vbInformation, App.Title
End If
Screen.MousePointer = 0
Set rsReporte = Nothing
Set vistaPrevia = Nothing
Set aplicacion = Nothing
Set reporte = Nothing

'vistaPrevia.CRViewer91.ReportSource = reporte
'vistaPrevia.Caption = strTituloReporte
'vistaPrevia.CRViewer91.ViewReport
'vistaPrevia.CRViewer91.DisplayGroupTree = True
'Screen.MousePointer = 0
'vistaPrevia.WindowState = 2
'
'vistaPrevia.Show
'
'Set vistaPrevia = Nothing
'Set aplicacion = Nothing
'Set pReport = Nothing
'Set reporte = Nothing
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set pReport = Nothing
    Set reporte = Nothing
End Sub
