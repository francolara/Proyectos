Attribute VB_Name = "mdlReportes"

Public Sub mostrarReporte(ByVal strNombreRpt As String, ByVal strNombreParametros As String, ByVal strValoresParametros As String, ByVal strTituloReporte As String, ByRef StrMsgError As String)
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim pReport         As ParameterFieldDefinition
Dim crxTable        As CRAXDRT.DatabaseTable
Dim arrNomParam() As String
Dim arrValParam() As String
Dim p As Integer
Dim CuentaTabla As Integer
On Error GoTo Err

'        Set pReport = reporte.ParameterFields.GetItemByName("@" & arrNomParam(p))
'        If pReport.ValueType = crNumberField Then
'            reporte.ParameterFields.GetItemByName("@" & arrNomParam(p)).AddCurrentValue ((Val("" & arrValParam(p))))
'        Else
'            reporte.ParameterFields.GetItemByName("@" & arrNomParam(p)).AddCurrentValue ("" & arrValParam(p) & "")
'        End If

    Set reporte = aplicacion.OpenReport(gStrRutaRpts & strNombreRpt)
    arrNomParam = Split(strNombreParametros, "|")
    arrValParam = Split(strValoresParametros, "|")
    

    '****** Cambia Servidor ********
'    For CuentaTabla = 1 To reporte.Database.Tables.Count
'        reporte.Database.Tables(CuentaTabla).ConnectionProperties.item("Data Source") = gbservidor
'        reporte.Database.Tables(CuentaTabla).ConnectionProperties.item("Initial Catalog") = gbDatabase
'        reporte.Database.Tables(CuentaTabla).ConnectionProperties("User ID") = gbusuario
'        reporte.Database.Tables(CuentaTabla).ConnectionProperties("Password") = gbPassword
'        reporte.Database.Tables(CuentaTabla).ConnectionProperties.item("Integrated Security") = True
'    Next CuentaTabla
    
    For Each crxTable In reporte.Database.Tables
        crxTable.ConnectionProperties.item("Data Source") = gbservidor
        crxTable.ConnectionProperties.item("Initial Catalog") = gbDatabase
        crxTable.ConnectionProperties.item("User ID") = gbusuario
        crxTable.ConnectionProperties.item("Password") = gbPassword
        crxTable.ConnectionProperties.item("Integrated Security") = False
    Next

    '*******************************
    
    For p = 0 To reporte.ParameterFields.Count - 1
        Set pReport = reporte.ParameterFields.GetItemByName("@" & arrNomParam(p))
        If pReport.ValueType = crNumberField Then
            pReport.AddCurrentValue (Val("" & arrValParam(p)))
        Else
            pReport.AddCurrentValue ("" & arrValParam(p) & "")
        End If
    Next
    
    vistaPrevia.CRViewer91.ReportSource = reporte
    vistaPrevia.Caption = strTituloReporte
    vistaPrevia.CRViewer91.ViewReport
    vistaPrevia.CRViewer91.DisplayGroupTree = True
    Screen.MousePointer = 0
    vistaPrevia.WindowState = 2
    
    vistaPrevia.Show
    
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set pReport = Nothing
    Set reporte = Nothing
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set pReport = Nothing
    Set reporte = Nothing
End Sub
