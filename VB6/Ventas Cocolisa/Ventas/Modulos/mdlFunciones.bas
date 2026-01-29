Attribute VB_Name = "mdlFunciones"
Public Declare Function ShellExecute _
                           Lib "shell32.dll" _
                           Alias "ShellExecuteA" ( _
                           ByVal hwnd As Long, _
                           ByVal lpOperation As String, _
                           ByVal lpFile As String, _
                           ByVal lpParameters As String, _
                           ByVal lpDirectory As String, _
                           ByVal nShowCmd As Long) _
                           As Long
 Public Sub Graba_Logico_Vales(ByRef Ope As String, ByRef StrMsgError As String, PIdSucursal As String, PIdValesCab As String, PTipoVale As String)
On Error GoTo Err
Dim Cadmysql                            As String
Dim RsConsulta                          As New ADODB.Recordset
Dim NItem_Log                           As Long
Dim CamposOpe                           As String

    If Ope = "1" Then
        CamposOpe = ",IdUsuarioModificacion, FechaModificacion, HoraModificacion"
    End If
    
    If Ope = "2" Then
        CamposOpe = ",IdUsuarioAnulacion, FechaAnulacion, HoraAnulacion"
    End If
    
    If Ope = "3" Then
        CamposOpe = ",IdUsuarioEliminacion, FechaEliminacion, HoraEliminacion"
    End If
    
    Cadmysql = "Select Item_Log " & _
               "From ValesCab_Log " & _
               "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & PIdSucursal & "' And idValescab = '" & PIdValesCab & "' And TipoVale = '" & PTipoVale & "' " & _
               "Order By Item_Log Desc "
    RsConsulta.Open Cadmysql, Cn, adOpenStatic, adLockReadOnly
    If Not RsConsulta.EOF Then
        NItem_Log = Val("" & RsConsulta.Fields("Item_Log")) + 1
    Else
        NItem_Log = 1
    End If
    If RsConsulta.State = 1 Then RsConsulta.Close
    Set RsConsulta = Nothing
    
    Cadmysql = "Insert into valescab_Log (" & _
               "idValesCab, tipoVale, fechaEmision, valorTotal, igvTotal, precioTotal, idProvCliente, idConcepto, idAlmacen, " & _
               "obsValesCab, idMoneda, GlsDocReferencia, TipoCambio, idEmpresa, idSucursal, estValeCab, idPeriodoInv, idCentroCosto, " & _
               "codanula , obsAnulacion, fecAnulacion, usuAnula, IdValeTemp, TipoValeRef, IdValesCabRef " & CamposOpe & " ,Item_Log,GlsPCLog,GlsPCUsuarioLog) " & _
               "SELECT idValesCab, tipoVale, fechaEmision, valorTotal, igvTotal, precioTotal, idProvCliente, idConcepto, idAlmacen, " & _
               "obsValesCab, idMoneda, GlsDocReferencia, TipoCambio, idEmpresa, idSucursal, estValeCab, idPeriodoInv, idCentroCosto, " & _
               "codanula , obsAnulacion, fecAnulacion, usuAnula, IdValeTemp, TipoValeRef, IdValesCabRef,'" & glsUser & "',CAST(GETDATE() AS DATE),CAST(GETDATE() AS TIME)," & _
               "" & NItem_Log & ",'" & ComputerName & "','" & fpUsuarioActual & "' " & _
               "FROM valescab " & _
               "Where idvalescab = '" & PIdValesCab & "' " & _
               "and tipovale = '" & PTipoVale & "' " & _
               "and idempresa = '" & glsEmpresa & "' " & _
               "and idsucursal = '" & PIdSucursal & "' "
    Cn.Execute (Cadmysql)
    
    Cadmysql = "Insert into valesDet_Log (" & _
               "idValesCab, Item, idProducto, GlsProducto, idUM, Factor, Afecto, Cantidad, VVUnit, IGVUnit, PVUnit, TotalVVNeto, TotalIGVNeto, " & _
               "TotalPVNeto, idMoneda, idEmpresa, idSucursal, NumLote, FecVencProd, Cantidad2, idDocumentoImp, idSerieImp, " & _
               " idDocVentasImp, idsucursalOrigen, idLote, tipoVale,Item_Log) " & _
               "SELECT idValesCab, Item, idProducto, GlsProducto, idUM, Factor, Afecto, Cantidad, VVUnit, IGVUnit, PVUnit, TotalVVNeto, TotalIGVNeto, " & _
               "TotalPVNeto, idMoneda, idEmpresa, idSucursal, NumLote, FecVencProd, Cantidad2, idDocumentoImp, idSerieImp, " & _
               " idDocVentasImp, idsucursalOrigen, idLote, tipoVale," & NItem_Log & " " & _
               "FROM valesDet " & _
               "Where idvalescab = '" & PIdValesCab & "' " & _
               "and tipovale = '" & PTipoVale & "' " & _
               "and idempresa = '" & glsEmpresa & "' " & _
               "and idsucursal = '" & PIdSucursal & "' "
    Cn.Execute (Cadmysql)
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub
Public Sub Actualiza_Stock_Nuevo(StrMsgError As String, PAccion As String, PGlsSucursal As String, PTipoVale As String, PIdValesCab As String, PIdAlmacen As String)
On Error GoTo Err
Dim CSqlC                           As String

    If PAccion = "E" Then 'Eliminar
        
        If PTipoVale = "I" Then
        
            CSqlC = "Update B Set B.Stock = B.Stock - A.Cantidad FROM ValesDet A " & _
                    "Inner Join ProductosStock B " & _
                        "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                    " " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                    "And A.IdValesCab = '" & PIdValesCab & "'"
            
            Cn.Execute CSqlC
            
            If Val("" & traerCampo("Almacenes", "IndStockDisponible", "IdAlmacen", PIdAlmacen, True)) = 1 Then
            
                CSqlC = "Update B Set B.Disponible = (B.Stock - A.Cantidad) - B.Separacion,B.Stock = B.Stock - A.Cantidad FROM ValesDet A " & _
                        "Inner Join ProductosStockDisponible B " & _
                            "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                        " " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                        "And A.IdValesCab = '" & PIdValesCab & "'"
                
                Cn.Execute CSqlC
            
            End If
            
            CSqlC = "Update B Set B.CantidadStock = B.CantidadStock - A.Cantidad FROM ValesDet A " & _
                    "Inner Join ProductosAlmacen B " & _
                        "On A.IdEmpresa = B.IdEmpresa And '" & PIdAlmacen & "' = B.IdAlmacen And A.IdProducto = B.IdProducto " & _
                    " " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                    "And A.IdValesCab = '" & PIdValesCab & "'"
                    
            Cn.Execute CSqlC
            
        Else
            
            CSqlC = "Update B Set B.Stock = B.Stock + A.Cantidad FROM ValesDet A " & _
                    "Inner Join ProductosStock B " & _
                        "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                    " " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                    "And A.IdValesCab = '" & PIdValesCab & "'"
            
            Cn.Execute CSqlC
            
            If Val("" & traerCampo("Almacenes", "IndStockDisponible", "IdAlmacen", PIdAlmacen, True)) = 1 Then
            
                CSqlC = "Update B Set B.Disponible = (B.Stock + A.Cantidad) - B.Separacion,B.Stock = B.Stock + A.Cantidad FROM ValesDet A " & _
                        "Inner Join ProductosStockDisponible B " & _
                            "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                        " " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                        "And A.IdValesCab = '" & PIdValesCab & "'"
                
                Cn.Execute CSqlC
            
            End If
            
            CSqlC = "Update B Set B.CantidadStock = B.CantidadStock + A.Cantidad FROM ValesDet A " & _
                    "Inner Join ProductosAlmacen B " & _
                        "On A.IdEmpresa = B.IdEmpresa And '" & PIdAlmacen & "' = B.IdAlmacen And A.IdProducto = B.IdProducto " & _
                    " " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                    "And A.IdValesCab = '" & PIdValesCab & "'"
                    
            Cn.Execute CSqlC
            
        End If
        
    Else 'I Insertar
    
        If PTipoVale = "I" Then
            
            CSqlC = "Update B Set B.Stock = B.Stock + A.Cantidad FROM ValesDet A " & _
                    "Inner Join ProductosStock B " & _
                        "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                    " " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                    "And A.IdValesCab = '" & PIdValesCab & "'"
            
            Cn.Execute CSqlC
            
            If Val("" & traerCampo("Almacenes", "IndStockDisponible", "IdAlmacen", PIdAlmacen, True)) = 1 Then
            
                CSqlC = "Update B Set B.Disponible = (B.Stock + A.Cantidad) - B.Separacion,B.Stock = B.Stock + A.Cantidad FROM ValesDet A " & _
                        "Inner Join ProductosStockDisponible B " & _
                            "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                        " " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                        "And A.IdValesCab = '" & PIdValesCab & "'"
                
                Cn.Execute CSqlC
            
            End If
            
            CSqlC = "Update B Set B.CantidadStock = B.CantidadStock + A.Cantidad FROM ValesDet A " & _
                    "Inner Join ProductosAlmacen B " & _
                        "On A.IdEmpresa = B.IdEmpresa And '" & PIdAlmacen & "' = B.IdAlmacen And A.IdProducto = B.IdProducto " & _
                    " " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                    "And A.IdValesCab = '" & PIdValesCab & "'"
                    
            Cn.Execute CSqlC
            
        Else
            
            CSqlC = "Update B Set B.Stock = B.Stock - A.Cantidad FROM ValesDet A " & _
                    "Inner Join ProductosStock B " & _
                        "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                    " " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                    "And A.IdValesCab = '" & PIdValesCab & "'"
            
            Cn.Execute CSqlC
            
            If Val("" & traerCampo("Almacenes", "IndStockDisponible", "IdAlmacen", PIdAlmacen, True)) = 1 Then
            
                CSqlC = "Update B Set B.Disponible = (B.Stock - A.Cantidad) - B.Separacion,B.Stock = B.Stock - A.Cantidad FROM ValesDet A " & _
                        "Inner Join ProductosStockDisponible B " & _
                            "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                        " " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                        "And A.IdValesCab = '" & PIdValesCab & "'"
                
                Cn.Execute CSqlC
            
            End If
            
            CSqlC = "Update B Set B.CantidadStock = B.CantidadStock - A.Cantidad FROM ValesDet A " & _
                    "Inner Join ProductosAlmacen B " & _
                        "On A.IdEmpresa = B.IdEmpresa And '" & PIdAlmacen & "' = B.IdAlmacen And A.IdProducto = B.IdProducto " & _
                    " " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & PGlsSucursal & "' And A.TipoVale = '" & PTipoVale & "' " & _
                    "And A.IdValesCab = '" & PIdValesCab & "'"
                    
            Cn.Execute CSqlC
            
        End If
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
End Sub

Public Sub ExportarReporte(ByVal strNombreRpt As String, ByVal strNombreParametros As String, ByVal strValoresParametros As String, ByVal strTituloReporte As String, ByRef strDescripcionDoc As String, ByRef StrMsgError As String)
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim pReport         As ParameterFieldDefinition
Dim arrNomParam()   As String
Dim arrValParam()   As String
Dim strPdfName      As String
Dim p               As Integer

On Error GoTo Err

Set reporte = aplicacion.OpenReport(gStrRutaRpts & strNombreRpt)

arrNomParam = Split(strNombreParametros, "|")
arrValParam = Split(strValoresParametros, "|")

For p = 0 To reporte.ParameterFields.Count - 1
    Set pReport = reporte.ParameterFields.GetItemByName(arrNomParam(p))
    
    If pReport.ValueType = crNumberField Then
        pReport.AddCurrentValue (Val("" & arrValParam(p)))
    Else
        pReport.AddCurrentValue ("'" & arrValParam(p) & "'")
    End If
Next

reporte.ExportOptions.DestinationType = crEDTDiskFile
reporte.ExportOptions.PDFExportAllPages = True
reporte.ExportOptions.FormatType = crEFTPortableDocFormat
reporte.ExportOptions.DiskFileName = App.Path & "\Temporales\" & strDescripcionDoc & ".pdf"
reporte.Export False

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
Public Function generaCorrelativoAnoMesFecha(strTabla As String, strCod As String, strFecha As String, Optional indEmpresa As Boolean = True)
Dim rst As New ADODB.Recordset
Dim dateSys As Date
Dim strCond As String
Dim csql As String

    dateSys = strFecha
    strCond = right(CStr(Year(dateSys)), 2) & Format(Month(dateSys), "00")
    csql = "SELECT " & strCod & " FROM " & strTabla & " WHERE left(" & strCod & ",4) = '" & strCond & "' "
    
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    csql = csql & " ORDER BY 1 DESC"
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        generaCorrelativoAnoMesFecha = strCond & Format((Val(right("" & rst.Fields(0), 4)) + 1), "0000")
    Else
        generaCorrelativoAnoMesFecha = strCond & "0001"
    End If
    rst.Close: Set rst = Nothing

End Function
Public Function generaCorrelativoAnoMes_ValeFecha(strTabla As String, strCod As String, StrTipVale As String, strFecha As String, Optional indEmpresa As Boolean = True)
Dim rst As New ADODB.Recordset
Dim dateSys As Date
Dim strCond As String
Dim csql As String

    dateSys = strFecha
    strCond = right(CStr(Year(dateSys)), 2) & Format(Month(dateSys), "00")
    csql = "SELECT " & strCod & " FROM " & strTabla & " WHERE left(" & strCod & ",4) = '" & strCond & "' " & _
           "and tipoVale = '" & StrTipVale & "' "

    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    csql = csql & " ORDER BY 1 DESC"
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        generaCorrelativoAnoMes_ValeFecha = strCond & Format((Val(right("" & rst.Fields(0), 4)) + 1), "0000")
    Else
        generaCorrelativoAnoMes_ValeFecha = strCond & "0001"
    End If
    rst.Close: Set rst = Nothing

End Function

Public Sub llenaCombo(C As ComboBox, tabla As String, campoDesc As String, mostrarPrimero As Boolean, Optional campoCod As String, Optional condicion As String)
Dim rs              As New ADODB.Recordset
Dim csql            As String
Dim texto           As String
    
    C.Clear
    
    csql = "Select " & campoDesc
    If campoCod <> "" Then csql = csql + "," & campoCod
    
    csql = csql + " From " + tabla
    If condicion <> "" Then csql = csql + " Where " & condicion
    
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rs.EOF
        texto = ""
        If Not IsNull(rs.Fields(campoDesc)) Then
            texto = Trim(rs.Fields(campoDesc))
            If campoCod <> "" Then
                If Not IsNull(rs.Fields(campoCod)) Then
                    texto = texto + Space(100) + Trim(rs.Fields(campoCod))
                End If
            End If
        End If
        If texto <> "" Then
            C.AddItem texto
        End If
        
        rs.MoveNext
    Loop
    If mostrarPrimero Then
        If C.ListCount > 0 Then C.ListIndex = 0
    End If
    rs.Close: Set rs = Nothing
    
End Sub

Public Sub traerCampos2(cone As Connection, tabla As String, campoTraer As String, campoComp As String, valor As String, nCantArray As Integer, ArrObject() As String, ByVal indEmpresa As Boolean, Optional condicion As String)
Dim rs              As New ADODB.Recordset
Dim csql            As String
Dim X               As Integer
    
    csql = "Select " & campoTraer & " From " & tabla & " where " & campoComp & " = '" & valor & "'"
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "' "
    End If
    If condicion <> "" Then csql = csql & " and " & condicion
    
    rs.Open csql, cone, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        For X = 0 To nCantArray - 1
            ArrObject(X) = IIf(Not IsNull(rs.Fields(X)), (rs.Fields(X)), "")
        Next X
    End If
    rs.Close: Set rs = Nothing

End Sub

Public Sub mostrarAyudaTextoProducto(strBus As String, ByRef strCod As String, ByRef strDes As String, Optional strAdic As String)
    
    FrmBusquedaProducto.ExecuteReturnText strBus, strCod, strDes, strAdic

End Sub

Public Sub traerCampos(tabla As String, campoTraer As String, campoComp As String, valor As String, nCantArray As Integer, ArrObject() As String, ByVal indEmpresa As Boolean, Optional condicion As String)
Dim rs              As New ADODB.Recordset
Dim csql            As String
Dim X               As Integer
    
    csql = "Select " & campoTraer & " From " & tabla & " where " & campoComp & " = '" & valor & "'"
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "' "
    End If
    
    If condicion <> "" Then csql = csql & " and " & condicion
    
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        For X = 0 To nCantArray - 1
            ArrObject(X) = IIf(Not IsNull(rs.Fields(X)), (rs.Fields(X)), "")
        Next X
    End If
    rs.Close: Set rs = Nothing

End Sub

Public Function traerCampo(tabla As String, campoTraer As String, campoComp As String, valor As String, ByVal indEmpresa As Boolean, Optional condicion As String) As String
Dim rs              As New ADODB.Recordset
Dim csql            As String
    
    csql = "Select " & campoTraer & " From " & tabla & " where " & campoComp & " = '" & valor & "'"
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "' "
    End If
    If condicion <> "" Then csql = csql & " and " & condicion
    
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    traerCampo = ""
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then traerCampo = (rs.Fields(0))
    End If
    rs.Close: Set rs = Nothing

End Function

Public Function traerCampoConta(tabla As String, campoTraer As String, campoComp As String, valor As String, ByVal indEmpresa As Boolean, Optional condicion As String) As String
Dim rs              As New ADODB.Recordset
Dim csql            As String
    
    csql = "Select " & campoTraer & " From " & tabla & " where " & campoComp & " = '" & valor & "'"
    
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "' "
    End If
    If condicion <> "" Then csql = csql & " and " & condicion
    
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    traerCampoConta = ""
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then traerCampoConta = (rs.Fields(0))
    End If
    rs.Close: Set rs = Nothing

End Function

Public Sub validaHomonimia(ByVal strTabla As String, ByVal strGlsCampoEvaluar As String, ByVal strIdCampoEvaluar As String, ByVal strValorEvaluar As String, ByVal strIdValorEvaluado As String, ByVal indEmpresa As Boolean, ByRef StrMsgError As String, Optional strCondicion As String, Optional strMensaje As String)
On Error GoTo Err
Dim rs              As New ADODB.Recordset
Dim csql            As String
    
    csql = "Select " & strIdCampoEvaluar & " From " & strTabla & " where (" & strGlsCampoEvaluar & " = '" & strValorEvaluar & "' "
    '--- Si existe alguna condicion extra
    If strCondicion <> "" Then
        If UCase(left(Trim(strCondicion), 2)) <> "OR" Then
            csql = csql & " and " & strCondicion & ")"
        Else
            csql = csql & strCondicion & ")"
        End If
    Else
        csql = csql & ")"
    End If
      
    '--- Si la tabla es por empresa
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "' "
    End If
    
    '--- Si es modificacion para q no tome el mismo resgistro
    If strIdValorEvaluado <> "" Then csql = csql + " and " & strIdCampoEvaluar & " <> '" & strIdValorEvaluado & "'"
    
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        If strMensaje = "" Then
            StrMsgError = "El registro o descripcion ya existe"
        Else
            StrMsgError = strMensaje
        End If
        GoTo Err
    End If
    If rs.State = 1 Then rs.Close: Set rs = Nothing

    Exit Sub
    
Err:
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub



Public Sub abrirConexion(ByRef StrMsgError As String, Optional cnInterna As Boolean = False)
On Error GoTo Err

   n = FreeFile()
   Open (App.Path & "\" & "Conexion.ini") For Input As n
    Do While Not EOF(n)
        Line Input #1, linea
      
        If Trim(Mid(linea, 1, Len("Servidor"))) = "Servidor" Then gbservidor = Trim(right(linea, Len(linea) - Len("Servidor =")))
        If Trim(Mid(linea, 1, Len("Database"))) = "Database" Then gbDatabase = Trim(right(linea, Len(linea) - Len("Database =")))
        If Trim(Mid(linea, 1, Len("Usuarios"))) = "usuarios" Then gbusuario = Trim(right(linea, Len(linea) - Len("Usuarios =")))
        If Trim(Mid(linea, 1, Len("Password"))) = "password" Then gbPassword = Trim(right(linea, Len(linea) - Len("Password =")))
        If Trim(Mid(linea, 1, Len("Rutasku"))) = "Rutasku" Then gbRutaProductos = Trim(right(linea, Len(linea) - Len("Rutasku =")))
        
    Loop
   Close n
   
    If Cn.State = 1 Then Cn.Close
    Cn.CursorLocation = adUseClient
    Cn.CommandTimeout = 0
    Cn.ConnectionString = "driver={SQL Server};server=" & gbservidor & ";uid=" & gbusuario & ";pwd=" & gbPassword & ";database=" & gbDatabase
    Cn.Open

Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub cerrarConexion(Optional cnInterna As Boolean = False)
    
    If cnInterna Then
        If Cnf.State = 1 Then Cnf.Close
        Set Cnf = Nothing
    Else
        If Cn.State = 1 Then Cn.Close
        Set Cn = Nothing
    End If

End Sub

Public Function funcLeeConfiguracion(pAppName As String, _
                                     ByRef gStrServidor As String, _
                                     ByRef gStrMotorBD As Integer, _
                                     ByRef gStrBD As String, _
                                     ByRef gStrTC As Integer, _
                                     ByRef gStrUsuario As String, _
                                     ByRef gStrClave As String, _
                                     ByRef gStrRutaRpts As String) As Long
On Error GoTo ErrMng
    
    gStrServidor = GetSetting(pAppName, CONF_CONEXION, CONF_SERVIDOR)
    gStrMotorBD = Val(GetSetting(pAppName, CONF_CONEXION, CONF_MOTORBD))
    gStrBD = GetSetting(pAppName, CONF_CONEXION, CONF_BD)
    gStrTC = Val(GetSetting(pAppName, CONF_CONEXION, CONF_TC))
    gStrUsuario = GetSetting(pAppName, CONF_CONEXION, CONF_USUARIO)
    gStrClave = GetSetting(pAppName, CONF_CONEXION, CONF_CLAVE)
    gStrRutaRpts = GetSetting(pAppName, CONF_OTROS, CONF_RUTA_RPTS)
        
    If (gStrServidor) <> "" And (gStrBD) <> "" Then
        funcLeeConfiguracion = 0
    Else
        funcLeeConfiguracion = -1
    End If
    
    Exit Function
    
ErrMng:
    funcLeeConfiguracion = -2
End Function

Public Sub funcGuardaConfiguracion(pAppName As String, _
                                   pServidor As String, _
                                   pMotorBD As Integer, _
                                   pBase As String, _
                                   pAutentificacion As Integer, _
                                   pUsuario As String, _
                                   pPassword As String, _
                                   Optional pRutaReporte As String)
               
    SaveSetting pAppName, CONF_CONEXION, CONF_MOTORBD, pMotorBD
    SaveSetting pAppName, CONF_CONEXION, CONF_SERVIDOR, pServidor
    SaveSetting pAppName, CONF_CONEXION, CONF_BD, pBase
    SaveSetting pAppName, CONF_CONEXION, CONF_USUARIO, pUsuario
    SaveSetting pAppName, CONF_CONEXION, CONF_CLAVE, pPassword
    SaveSetting pAppName, CONF_OTROS, CONF_RUTA_RPTS, pRutaReporte
    SaveSetting pAppName, CONF_CONEXION, CONF_TC, pAutentificacion
  
End Sub

Public Function generaCorrelativo(tabla As String, campo As String, caracteres As Integer, Optional textoIni As String, Optional indEmpresa As Boolean = True, Optional strCondicion As String = "") As String
Dim rs As New ADODB.Recordset
Dim csql As String
Dim valor As String
Dim i As Integer
Dim StrMsgError As String

    For i = 0 To caracteres - 1
        strFormato = strFormato + "0"
    Next
    
    csql = "Select max(" + campo + ") From " + tabla
    
    csql = csql + " where 1 = 1 "
    If textoIni <> "" Then
        csql = csql + " AND " + campo + " like '" + textoIni + "%'"
    End If
    
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    If strCondicion <> "" Then
        csql = csql & " AND " & strCondicion
    End If
    
    valor = 0
    If textoIni <> "" Then
        valor = strFormato
        For i = 0 To Len(textoIni) - 1
            valor = valor + "0"
        Next
    End If
    
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF And Not IsNull(rs.Fields(0)) Then valor = rs.Fields(0)
    If textoIni = "" Then
        generaCorrelativo = Format(Val(valor) + 1, strFormato)
    Else
        generaCorrelativo = textoIni + Format(Val(right(valor, Len(valor) - Len(textoIni))) + 1, strFormato)
    End If

End Function

Public Sub limpiaForm(m As Form)
Dim C As Object

    For Each C In m.Controls
        If TypeOf C Is TextBox Or TypeOf C Is CATTextBox Then
            C.Text = ""
        End If
        
        If TypeOf C Is DTPicker Then
            C.Value = getFechaSistema
        End If
    Next

End Sub

Public Sub ubicaDatoCombo(C As ComboBox, valor As String, caracteres As Integer)
Dim i As Integer

    For i = 0 To C.ListCount - 1
        If right(C.List(i), caracteres) = valor Then
            C.ListIndex = i
            Exit Sub
        End If
    Next
    C.ListIndex = -1

End Sub

'--- Ayuda Clasica
Public Sub mostrarAyuda(strBus As String, ByRef TextBox1 As Object, ByRef TextBox2 As Object, Optional strAdic As String)
    
    FrmBusqueda.Execute strBus, TextBox1, TextBox2, strAdic

End Sub

Public Sub mostrarAyudaTexto(strBus As String, ByRef strCod As String, ByRef strDes As String, Optional strAdic As String)
    
    FrmBusqueda.ExecuteReturnText strBus, strCod, strDes, strAdic

End Sub

Public Sub mostrarAyudaKeyascii(KeyAscii As Integer, strBus As String, ByRef TextBox1 As Object, ByRef TextBox2 As Object, Optional strAdic As String)
    
    FrmBusqueda.ExecuteKeyascii KeyAscii, strBus, TextBox1, TextBox2, strAdic

End Sub

Public Sub mostrarAyudaKeyasciiTexto(KeyAscii As Integer, strBus As String, ByRef strCod As String, ByRef strDes As String, Optional strAdic As String)
    
    FrmBusqueda.ExecuteKeyasciiReturnText KeyAscii, strBus, strCod, strDes, strAdic

End Sub

'--- Ayuda de productos
Public Sub mostrarAyudaTextoProdAlm(strAlm As String, ByRef rspa As ADODB.Recordset, ByRef strCod As String, ByRef strDes As String, ByRef strCodUM As String, ByVal indValidaStock As Boolean, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef strTipoDoc As String, ByRef StrMsgError As String, Optional strAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1, Optional PIdCliente As String, Optional PCodMotivo As String)
    
    FrmAyudaProductos.ExecuteReturnTextAlm strAlm, rspa, strCod, strDes, strCodUM, indValidaStock, strVarCodLista, indVarUMVenta, indVarMostrarPresentaciones, indVarPedido, strTipoDoc, StrMsgError, strAdic, indMostrarTP, TipoProd, PIdCliente, PCodMotivo

End Sub

Public Sub mostrarAyudaKeyasciiTextoProdAlm(KeyAscii As Integer, strAlm As String, ByRef strCod As String, ByRef strDes As String, ByRef strCodUM As String, ByVal indValidaStock As Boolean, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef StrMsgError As String, Optional strAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
    
    FrmAyudaProductos.ExecuteKeyasciiReturnTextAlm KeyAscii, strAlm, strCod, strDes, strCodUM, indValidaStock, strVarCodLista, indVarUMVenta, indVarMostrarPresentaciones, indVarPedido, StrMsgError, strAdic, indMostrarTP, TipoProd

End Sub

'--- Ayuda de Precios
Public Sub mostrarAyudaTextoPrecios(strCodProd As String, strCodLista As String, ByRef strCod As String, ByRef strDes As String, ByRef dblFactor As Double, Optional strAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
    
    FrmAyudaPrecios.ExecuteReturnText strCodProd, strCodLista, strCod, strDes, strAdic, dblFactor

End Sub

Public Sub mostrarAyudaKeyasciiTextoPrecios(KeyAscii As Integer, strCodProd As String, strCodLista As String, ByRef strCod As String, ByRef strDes As String, Optional strAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
    
    FrmAyudaPrecios.ExecuteKeyasciiReturnText KeyAscii, strCodProd, strCodLista, strCod, strDes, strAdic

End Sub

'--- Ayuda de Clientes
Public Sub mostrarAyudaClientes(ByRef TextBox1 As Object, ByRef TextBox2 As Object, ByRef TextRUC As Object, ByRef TextDireccion As Object, ByRef TextCodtienda As Object, Optional strAdic As String, Optional inddirll As Boolean)
    
    FrmBusquedaClientes.Execute TextBox1, TextBox2, TextRUC, TextDireccion, TextCodtienda, strAdic, inddirll

End Sub

Public Sub mostrarAyudaClientesKeyascii(KeyAscii As Integer, ByRef TextBox1 As Object, ByRef TextBox2 As Object, ByRef TextRUC As Object, ByRef TextDireccion As Object, Optional strAdic As String)
    
    FrmBusquedaClientes.ExecuteKeyascii KeyAscii, TextBox1, TextBox2, TextRUC, TextDireccion, strAdic

End Sub

Public Sub mostrarAyudaTextoPlanCuentas(cone As String, strBus As String, ByRef strCod As String, ByRef strDes As String, Optional strAdic As String, Optional strAnno As String)
    
   FrmAyudaPlanCuenta.ExecuteReturnText cone, strBus, strCod, strDes, strAdic, strAnno

End Sub

Public Function getFechaSistema() As String
Dim rst As New ADODB.Recordset
    
    rst.Open "select Getdate() ", Cn, adOpenStatic, adLockReadOnly
    getFechaSistema = "" & Format(rst.Fields(0), "dd/mm/yyyy")
    rst.Close: Set rst = Nothing

End Function

Public Function getFechaHoraSistema() As String
Dim rst As New ADODB.Recordset
    
    rst.Open "select getdate() ", Cn, adOpenStatic, adLockReadOnly
    getFechaHoraSistema = CStr("" & rst.Fields(0))
    rst.Close: Set rst = Nothing
    
End Function

Public Function GeneraCorrelativoAnoMes(strTabla As String, strCod As String, Optional indEmpresa As Boolean = True)
Dim rst As New ADODB.Recordset
Dim dateSys As Date
Dim strCond As String
Dim csql As String

    dateSys = getFechaSistema
    strCond = right(CStr(Year(dateSys)), 2) & Format(Month(dateSys), "00")
    csql = "SELECT " & strCod & " FROM " & strTabla & " WHERE left(" & strCod & ",4) = '" & strCond & "' "
    
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    csql = csql & " ORDER BY 1 DESC"
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        GeneraCorrelativoAnoMes = strCond & Format((Val(right("" & rst.Fields(0), 4)) + 1), "0000")
    Else
        GeneraCorrelativoAnoMes = strCond & "0001"
    End If
    rst.Close: Set rst = Nothing
    
End Function


Public Sub mostrarAyudaTextoProdAlm2(strAlm As String, ByRef strCod As String, ByRef strDes As String, ByRef strCodUM As String, ByVal indValidaStock As Boolean, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef StrMsgError As String, Optional strAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
    
    FrmAyudaProductos_2.ExecuteReturnTextAlm strAlm, strCod, strDes, strCodUM, indValidaStock, strVarCodLista, indVarUMVenta, indVarMostrarPresentaciones, indVarPedido, StrMsgError, strAdic, indMostrarTP, TipoProd

End Sub

Public Function GeneraCorrelativoAnoMesAdd(strTabla As String, strCod As String, Optional indEmpresa As Boolean = True, Optional PSqlAdd As String)
Dim rst As New ADODB.Recordset
Dim dateSys As Date
Dim strCond As String
Dim csql As String
    
    dateSys = getFechaSistema
    strCond = right(CStr(Year(dateSys)), 2) & Format(Month(dateSys), "00")
    csql = "SELECT " & strCod & " FROM " & strTabla & " WHERE left(" & strCod & ",4) = '" & strCond & "' "
    
    If PSqlAdd <> "" Then
        csql = csql & "And " & PSqlAdd & " "
    End If
    
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    csql = csql & " ORDER BY 1 DESC"
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        GeneraCorrelativoAnoMesAdd = strCond & Format((Val(right("" & rst.Fields(0), 4)) + 1), "0000")
    Else
        GeneraCorrelativoAnoMesAdd = strCond & "0001"
    End If
    rst.Close: Set rst = Nothing
    
End Function

Public Sub EjecutaSQLForm_1(F As Form, tipoOperacion As Integer, indEmpresa As Boolean, strTabla As String, ByRef StrMsgError As String, Optional strCampoCod As String, Optional g As dxDBGrid, Optional strTablaDet As String, Optional strCampoDet As String, Optional strDataCampo As String, Optional indFechaRegistro As Boolean = False)
On Error GoTo Err
Dim C As Object
Dim csql As String
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strValCod As String
Dim strCampoEmpresa As String
Dim strValorEmpresa As String
Dim strCondEmpresa As String
Dim strCampoFecReg As String
Dim strValorFecReg As String
Dim strCondSistema As String
Dim indTrans As Boolean

    If indEmpresa Then
        strCampoEmpresa = ",idEmpresa"
        strValorEmpresa = ",'" & glsEmpresa & "'"
        strCondEmpresa = " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    strCondSistema = "  And CodSistema = '" & StrcodSistema & "' "
    
    If indFechaRegistro Then
        strCampoFecReg = ",FecRegistro"
        strValorFecReg = ",sysdate()"
    End If
    
    indTrans = False
    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & C.Value & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    Case 1
                        If UCase(strCampoCod) <> UCase(strCampo) Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = C.Value
                                Case "T"
                                    strValores = "'" & C.Value & "'"
                                Case "F"
                                    strValores = "'" & Format(C.Value, "yyyy-mm-dd") & "'"
                            End Select
                            strCampos = strCampos & strCampo & "=" & strValores & ","
                        Else
                            strValCod = C.Value
                        End If
                End Select
            End If
        End If
    Next
    
    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
    
    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            csql = "INSERT INTO " & strTabla & "(" & strCampos & strCampoEmpresa & strCampoFecReg & " ,codSistema) VALUES(" & strValores & strValorEmpresa & strValorFecReg & " ,'" & StrcodSistema & "' )"
        Case 1
            csql = "UPDATE " & strTabla & " SET " & strCampos & "  " & _
                   "WHERE " & strCampoCod & " = '" & strValCod & "'" & strCondEmpresa & strCondSistema
    End Select
        
    indTrans = True
    Cn.BeginTrans
    
    '--- Graba controles
    If strCampos <> "" Then
        Cn.Execute csql
    End If
    
    '--- Grabando Grilla
    If TypeName(g) <> "Nothing" Then
        Cn.Execute "DELETE FROM " & strTablaDet & " WHERE " & strCampoDet & " = '" & strDataCampo & "'" & strCondEmpresa
        
        g.Dataset.First
        Do While Not g.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To g.Columns.Count - 1
                If UCase(left(g.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(g.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(g.Columns(i).ObjectName, 3)
                    
                    strCampos = strCampos & strCampo & ","
                    
                    Select Case strTipoDato
                        Case "N"
                            strValores = strValores & g.Columns(i).Value & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO " & strTablaDet & "(" & strCampos & "," & strCampoDet & strCampoEmpresa & ") VALUES(" & strValores & ",'" & strDataCampo & "'" & strValorEmpresa & ")"
            Cn.Execute csql
            
            g.Dataset.Next
        Loop
    End If
    
    Cn.CommitTrans
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
End Sub

'--- TipoOperacion : 0 = insert     1 = Update
Public Sub EjecutaSQLForm(F As Form, tipoOperacion As Integer, indEmpresa As Boolean, strTabla As String, ByRef StrMsgError As String, Optional strCampoCod As String, Optional g As dxDBGrid, Optional strTablaDet As String, Optional strCampoDet As String, Optional strDataCampo As String, Optional indFechaRegistro As Boolean = False)
On Error GoTo Err
Dim C As Object
Dim csql As String
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strValCod As String
Dim strCampoEmpresa As String
Dim strValorEmpresa As String
Dim strCondEmpresa As String
Dim strCampoFecReg As String
Dim strValorFecReg As String
Dim GlsObsCli   As String
Dim indTrans As Boolean

    If indEmpresa Then
        strCampoEmpresa = ",idEmpresa"
        strValorEmpresa = ",'" & glsEmpresa & "'"
        strCondEmpresa = " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    If indFechaRegistro Then
        strCampoFecReg = ",FecRegistro"
        strValorFecReg = ",sysdate()"
    End If

    indTrans = False
    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & C.Value & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    Case 1
                        If UCase(strCampoCod) <> UCase(strCampo) Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = C.Value
                                Case "T"
                                    strValores = "'" & C.Value & "'"
                                Case "F"
                                    strValores = "'" & Format(C.Value, "yyyy-mm-dd") & "'"
                            End Select
                            strCampos = strCampos & strCampo & "=" & strValores & ","
                        Else
                            strValCod = C.Value
                        End If
                End Select
            End If
        End If
    Next
    
    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
    
    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            csql = "INSERT INTO " & strTabla & "(" & strCampos & strCampoEmpresa & strCampoFecReg & ") VALUES(" & strValores & strValorEmpresa & strValorFecReg & ")"
        Case 1
            csql = "UPDATE " & strTabla & " SET " & strCampos & " WHERE " & strCampoCod & " = '" & strValCod & "'" & strCondEmpresa
    End Select

    indTrans = True
    Cn.BeginTrans
    
    '--- Graba controles
    If strCampos <> "" Then
        Cn.Execute csql
    End If
    
    GlsObsCli = "" & glsobservacioncliente
    
    If UCase(strTabla) = "CLIENTES" Then
        csql = "Update Clientes Set GlsObservacion = '" & GlsObsCli & "' Where Idcliente ='" & strDataCampo & "' And idEmpresa= '" & glsEmpresa & "' "
               Cn.Execute (csql)
    ElseIf UCase(strTabla) = "PROVEEDORES" Then
        csql = "Update proveedores Set GlsObservacion = '" & GlsObsCli & "' Where Idproveedor ='" & strDataCampo & "' And idEmpresa= '" & glsEmpresa & "' "
               Cn.Execute (csql)
    End If
    
    '--- Grabando Grilla
    If TypeName(g) <> "Nothing" Then
        Cn.Execute "DELETE FROM " & strTablaDet & " WHERE " & strCampoDet & " = '" & strDataCampo & "'" & strCondEmpresa
        
        g.Dataset.First
        Do While Not g.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To g.Columns.Count - 1
                If UCase(left(g.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(g.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(g.Columns(i).ObjectName, 3)
                    
                    strCampos = strCampos & strCampo & ","
                    
                    Select Case strTipoDato
                        Case "N"
                            strValores = strValores & g.Columns(i).Value & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO " & strTablaDet & "(" & strCampos & "," & strCampoDet & strCampoEmpresa & ") VALUES(" & strValores & ",'" & strDataCampo & "'" & strValorEmpresa & ")"
            Cn.Execute csql
            
            g.Dataset.Next
        Loop
    End If
    
    Cn.CommitTrans
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
End Sub

Public Sub ConfGrid_Inv(g As dxDBGrid, indMod As Boolean, Optional mostrarFooter As Boolean, Optional mostrarGroupPanel As Boolean, Optional mostrarBandas As Boolean, Optional POrden As Boolean)
     
     With g.Options
        If indMod Then
            .Set (egoEditing)
            .Set (egoCanDelete)
            .Set (egoCanInsert)
        End If
        If mostrarBandas Then .Set (egoShowBands)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        If POrden Then .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoNameCaseInsensitive)
        If mostrarFooter Then .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        
        If mostrarGroupPanel Then .Set (egoShowGroupPanel)
        .Set (egoEnableNodeDragging)
        .Set (egoDragCollapse)
        .Set (egoDragExpand)
        .Set (egoDragScroll)
        .Set (egoEnableNodeDragging)
    End With
    g.Filter.DropDownCount = 30
    
End Sub

Public Sub ConfGrid(g As dxDBGrid, indMod As Boolean, Optional mostrarFooter As Boolean, Optional mostrarGroupPanel As Boolean, Optional mostrarBandas As Boolean)
        
    With g.Options
        If indMod Then
            .Set (egoEditing)
            .Set (egoCanInsert)
        End If
        
        If mostrarBandas Then .Set (egoShowBands)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoNameCaseInsensitive)
        If mostrarFooter Then .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        
        If mostrarGroupPanel Then .Set (egoShowGroupPanel)
        .Set (egoEnableNodeDragging)
        .Set (egoDragCollapse)
        .Set (egoDragExpand)
        .Set (egoDragScroll)
        .Set (egoEnableNodeDragging)
    End With
    g.Filter.DropDownCount = 30
    
End Sub

Public Sub validaFormSQL(F As Form, ByRef StrMsgError As String)
Dim C As Object

    For Each C In F.Controls
        If TypeOf C Is CATTextBox Then
            If C.Vacio = False And C.Visible = True Then
                If C.Estilo >= 3 Then '--- Numerico
                    If Val(C.Value) = 0 Then
                        C.OnError = True
                        StrMsgError = "Debe ingresar un monto"
                        Exit Sub
                    End If
                Else
                    If Trim(C.Value) = "" Then
                        C.OnError = True
                        StrMsgError = "Faltan Datos"
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next

End Sub

Public Sub mostrarDatosFormSQL(F As Form, ByVal r As Recordset, ByRef StrMsgError As String, Optional ByVal objContenedor As Object)
On Error GoTo Err
   
    If TypeName(objContenedor) = "Nothing" Then
        For Each Ctrl In F
            If Ctrl.Tag <> "" Then
                Ctrl.DataField = right(Ctrl.Tag, Len(Ctrl.Tag) - 1)
                Set Ctrl.DataSource = r
                Set Ctrl.DataSource = Nothing
            End If
        Next
    Else
        For Each Ctrl In F
            If Ctrl.Tag <> "" Then
                If Ctrl.Container.Name = objContenedor.Name Then
                    Ctrl.DataField = right(Ctrl.Tag, Len(Ctrl.Tag) - 1)
                    Set Ctrl.DataSource = r
                    Set Ctrl.DataSource = Nothing
                End If
            End If
       Next
    End If
    r.Close: Set r = Nothing
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub mostrarDatosGridSQL(g As dxDBGrid, r As Recordset, ByRef StrMsgError As String)
On Error GoTo Err
    
    With g
        .DefaultFields = False
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        Set .DataSource = r
        .Dataset.Active = True
        .KeyField = "Item"
        .Dataset.Edit
        .Dataset.Post
    End With

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Function leeParametro(strParametro As String) As String
Dim rs              As New ADODB.Recordset
Dim csql            As String
    
    csql = "Select ValParametro From parametros where GlsParametro = '" & strParametro & "' AND idEmpresa = '" & glsEmpresa & "'"
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    leeParametro = ""
    If Not rs.EOF Then
        If Not IsNull(rs.Fields("ValParametro")) Then leeParametro = (rs.Fields("ValParametro"))
    End If
    rs.Close: Set rs = Nothing
    
End Function

Public Function LeeParametroConta(strParametro As String) As String
Dim rs              As New ADODB.Recordset
Dim csql            As String
    
    csql = "Select ValParametro From paramConta where GlsParametro = '" & strParametro & "' AND idEmpresa = '" & glsEmpresa & "'"
    rs.Open csql, CnConta, adOpenForwardOnly, adLockReadOnly
    LeeParametroConta = ""
    If Not rs.EOF Then
        If Not IsNull(rs.Fields("ValParametro")) Then LeeParametroConta = (rs.Fields("ValParametro"))
    End If
    rs.Close: Set rs = Nothing
    
End Function

Public Function existeEnGrilla(g As dxDBGrid, strNomCampo As String, strValor As String) As Boolean
Dim intItem As Integer
Dim fil As Integer
 
    existeEnGrilla = False
    intItem = g.Columns.ColumnByFieldName("item").Value
    fil = g.Dataset.RecNo
    fil = g.Dataset.RecNo
    g.Dataset.First
    Do While Not g.Dataset.EOF
        If g.Columns.ColumnByFieldName("item").Value <> intItem Then
            If g.Columns.ColumnByFieldName(strNomCampo).Value = strValor Then
                existeEnGrilla = True
            End If
        End If
        g.Dataset.Next
    Loop
    g.Dataset.RecNo = fil
    g.Dataset.RecNo = fil

End Function

Function Cadenanum(PNum As Double, pmon As String) As String
Dim WDECIMAL    As String * 2
Dim WENTERO     As String
Dim wMoneda     As String
Dim WCADENA     As String
Dim wcont       As Integer
Dim WSUBENT     As String

    WDECIMAL = right(Format$(PNum, "#0.00"), 2)
    WENTERO = left(Format$(PNum, "#0.00"), Len(Format$(PNum, "#0.00")) - 3)
    wMoneda = pmon
    
    WCADENA = ""
    wcont = 0
    WSUBENT = WENTERO
    Do While wcont < Len(WENTERO)
        WSUBENT = right(WENTERO, Len(WENTERO) - wcont)
        Select Case Len(WSUBENT)
        Case Is = 3, 6, 9: WCADENA = WCADENA & FCENTENA(Mid(WSUBENT, 1, 3))
        Case Is = 2, 5, 8
            If Val(Mid(WSUBENT, 1, 2)) > 15 Then
                WCADENA = WCADENA & FDECENA(Mid(WSUBENT, 1, 2))
            Else
                WCADENA = WCADENA & FUNIDAD(Mid(WSUBENT, 1, 2), Len(WSUBENT), Val(WENTERO))
                wcont = wcont + 1
            End If
        Case Is = 1, 4, 7: WCADENA = WCADENA & FUNIDAD(Mid(WSUBENT, 1, 1), Len(WSUBENT), Val(WENTERO))
        End Select
        wcont = wcont + 1
    Loop
    Cadenanum = WCADENA & " Y " & WDECIMAL & "/100 " & wMoneda

End Function

Function FCENTENA(PCAD As String)
ReDim WUNI(10) As String
Dim WSUBCAD     As String
    
    WUNI(0) = " "
    WUNI(1) = "CIENTO "
    WUNI(2) = "DOSCIENTOS "
    WUNI(3) = "TRESCIENTOS "
    WUNI(4) = "CUATROCIENTOS "
    WUNI(5) = "QUINIENTOS "
    WUNI(6) = "SEISCIENTOS "
    WUNI(7) = "SETECIENTOS "
    WUNI(8) = "OCHOCIENTOS "
    WUNI(9) = "NOVECIENTOS "

    If PCAD = "100" Then
        WSUBCAD = "CIEN"
    Else
        WSUBCAD = WUNI(Val(left(PCAD, 1)))
    End If
    FCENTENA = WSUBCAD

End Function

Function FDECENA(PCAD As String) As String
ReDim WUNI(10) As String
Dim WSUBCAD     As String
    
    WCAD = left(PCAD, 2)
    WUNI(0) = " "
    WUNI(1) = "DIEZ "
    WUNI(2) = "VEINTE "
    WUNI(3) = "TREINTA "
    WUNI(4) = "CUARENTA "
    WUNI(5) = "CINCUENTA "
    WUNI(6) = "SESENTA "
    WUNI(7) = "SETENTA "
    WUNI(8) = "OCHENTA "
    WUNI(9) = "NOVENTA "

    If right(PCAD, 1) = 0 Then
        WSUBCAD = WUNI(Val(left(PCAD, 1)))
    Else
        WSUBCAD = WUNI(Val(left(PCAD, 1))) & "Y "
    End If
    FDECENA = WSUBCAD

End Function

Function FUNIDAD(PCAD As String, PLEN As Integer, PNum As Double) As String
ReDim WUNI(16) As String
Dim WSUBCAD     As String
    
    WUNI(0) = " "
    WUNI(1) = "UN "
    WUNI(2) = "DOS "
    WUNI(3) = "TRES "
    WUNI(4) = "CUATRO "
    WUNI(5) = "CINCO "
    WUNI(6) = "SEIS "
    WUNI(7) = "SIETE "
    WUNI(8) = "OCHO "
    WUNI(9) = "NUEVE "
    WUNI(10) = "DIEZ "
    WUNI(11) = "ONCE "
    WUNI(12) = "DOCE "
    WUNI(13) = "TRECE "
    WUNI(14) = "CATORCE "
    WUNI(15) = "QUINCE "
           
    Select Case PLEN
        Case Is = 1, 2: WSUBCAD = WUNI(Val(PCAD))
        
        Case Is = 4, 5:
            
            If Val(left(right(PNum, 6), 3)) = 0 Then
                WSUBCAD = WSUBCAD
            Else
                WSUBCAD = WUNI(Val(PCAD)) & "MIL "
            End If
        
        Case Is = 7, 8: WSUBCAD = WUNI(Val(PCAD)) & IIf(PCAD = "1", "MILLON ", "MILLONES ")
    End Select
    FUNIDAD = WSUBCAD

End Function

Public Function EnLetras(numero As String, strMoneda As String) As String
Dim B, paso
Dim expresion, entero, deci, flag
        
    flag = "N"
    For paso = 1 To Len(numero)
        If Mid(numero, paso, 1) = "." Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero
            Else
                deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero
            End If
        End If
    Next
    
    If Len(deci) = 1 Then
        deci = deci & "0"
    End If
    
    flag = "N"
    If Int(numero) >= -999999999 And Int(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
        For paso = Len(entero) To 1 Step -1
            B = Len(entero) - (paso - 1)
            Select Case paso
            Case 3, 6, 9
                Select Case Mid(entero, B, 1)
                    Case "1"
                        If Mid(entero, B + 1, 1) = "0" And Mid(entero, B + 2, 1) = "0" Then
                            expresion = expresion & "CIEN "
                        Else
                            expresion = expresion & "CIENTO "
                        End If
                    Case "2"
                        expresion = expresion & "DOSCIENTOS "
                    Case "3"
                        expresion = expresion & "TRESCIENTOS "
                    Case "4"
                        expresion = expresion & "CUATROCIENTOS "
                    Case "5"
                        expresion = expresion & "QUINIENTOS "
                    Case "6"
                        expresion = expresion & "SEISCIENTOS "
                    Case "7"
                        expresion = expresion & "SETECIENTOS "
                    Case "8"
                        expresion = expresion & "OCHOCIENTOS "
                    Case "9"
                        expresion = expresion & "NOVECIENTOS "
                End Select
                
            Case 2, 5, 8
                Select Case Mid(entero, B, 1)
                    Case "1"
                        If Mid(entero, B + 1, 1) = "0" Then
                            flag = "S"
                            expresion = expresion & "DIEZ "
                        End If
                        If Mid(entero, B + 1, 1) = "1" Then
                            flag = "S"
                            expresion = expresion & "ONCE "
                        End If
                        If Mid(entero, B + 1, 1) = "2" Then
                            flag = "S"
                            expresion = expresion & "DOCE "
                        End If
                        If Mid(entero, B + 1, 1) = "3" Then
                            flag = "S"
                            expresion = expresion & "TRECE "
                        End If
                        If Mid(entero, B + 1, 1) = "4" Then
                            flag = "S"
                            expresion = expresion & "CATORCE "
                        End If
                        If Mid(entero, B + 1, 1) = "5" Then
                            flag = "S"
                            expresion = expresion & "QUINCE "
                        End If
                        If Mid(entero, B + 1, 1) > "5" Then
                            flag = "N"
                            expresion = expresion & "DIECI"
                        End If
                
                    Case "2"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "VEINTE "
                            flag = "S"
                        Else
                            expresion = expresion & "VEINTE"
                            flag = "N"
                        End If
                    
                    Case "3"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "TREINTA "
                            flag = "S"
                        Else
                            expresion = expresion & "TREINTA Y "
                            flag = "N"
                        End If
                
                    Case "4"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "CUARENTA "
                            flag = "S"
                        Else
                            expresion = expresion & "CUARENTA Y "
                            flag = "N"
                        End If
                
                    Case "5"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "CINCUENTA "
                            flag = "S"
                        Else
                            expresion = expresion & "CINCUENTA Y "
                            flag = "N"
                        End If
                
                    Case "6"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "SESENTA "
                            flag = "S"
                        Else
                            expresion = expresion & "SESENTA Y "
                            flag = "N"
                        End If
                
                    Case "7"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "SETENTA "
                            flag = "S"
                        Else
                            expresion = expresion & "SETENTA Y "
                            flag = "N"
                        End If
                
                    Case "8"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "OCHENTA "
                            flag = "S"
                        Else
                            expresion = expresion & "OCHENTA Y "
                            flag = "N"
                        End If
                
                    Case "9"
                        If Mid(entero, B + 1, 1) = "0" Then
                            expresion = expresion & "NOVENTA "
                            flag = "S"
                        Else
                            expresion = expresion & "NOVENTA Y "
                            flag = "N"
                        End If
                End Select
                
            Case 1, 4, 7
                Select Case Mid(entero, B, 1)
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "UNO "
                            Else
                                expresion = expresion & "UN "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            expresion = expresion & "DOS "
                        End If
                    Case "3"
                        If flag = "N" Then
                            expresion = expresion & "TRES "
                        End If
                    Case "4"
                        If flag = "N" Then
                            expresion = expresion & "CUATRO "
                        End If
                    Case "5"
                        If flag = "N" Then
                            expresion = expresion & "CINCO "
                        End If
                    Case "6"
                        If flag = "N" Then
                            expresion = expresion & "SEIS "
                        End If
                    Case "7"
                        If flag = "N" Then
                            expresion = expresion & "SIETE "
                        End If
                    Case "8"
                        If flag = "N" Then
                            expresion = expresion & "OCHO "

                        End If
                    Case "9"
                        If flag = "N" Then
                            expresion = expresion & "NUEVE "
                        End If
                End Select
            End Select
            If paso = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And Len(entero) <= 6) Then
                    expresion = expresion & "MIL "
                End If
            End If
            If paso = 7 Then
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "MILLON "
                Else
                    expresion = expresion & "MILLONES "
                End If
            End If
        Next
        
        If deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion & "con " & deci ' & "/100"
            Else
                EnLetras = expresion & "Y " & deci ' & "/100"
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "MENOS " & expresion
            Else
                EnLetras = expresion
            End If
        End If
        EnLetras = "SON:" & EnLetras & "/100 " & strMoneda
    
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        EnLetras = ""
    End If

End Function

Public Sub validaMenu(m As MDIForm, usu As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsform As New ADODB.Recordset
Dim C As Control
Dim strEstado As String
Dim intIndex As Integer
Dim ContGrupo As Integer

    '--- Deshabilitamos todos las opciones a disable
    For Each C In frmPrincipal.Controls
        If left(C.Name, 3) = "mnu" Then
            If Len(C.Name) <> "5" Then
                If C.Caption <> "-" Then
                    If C.Enabled = True Then
                        C.Enabled = False
                    End If
                End If
            End If
        End If
    Next
            
    '--- Asigna opciones del perfil del usuario
    If rsform.State = 1 Then rsform.Close: Set rsform = Nothing
    'csql = "Select o.opmNum from opcionesperfil o where o.idEmpresa = '" & glsEmpresa & "' and o.CodSistema = '" & StrcodSistema & "' AND o.idPerfil = (select P.idPerfil From perfilesporusuario p WHERE p.idEmpresa = '" & glsEmpresa & "'  AND p.idUsuario = '" & usu & "' and CodSistema = '" & StrcodSistema & "')"
    csql = "EXEC spu_Usuario_OpcPerfil 1,'" & usu & "','" & glsEmpresa & "','" & StrcodSistema & "' "
    rsform.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If rsform.EOF Then
        m.dxSideBar1.Groups(0).Visible = False
        m.dxSideBar1.Groups(1).Visible = False
        
        For Each C In m.Controls
            If left(C.Name, 3) = "mnu" Then
                If C.Enabled = False Then
                    C.Visible = False
                End If
            End If
        Next
        StrMsgError = "Usted no cuenta con permisos en el sistema"
        GoTo Err
    End If
    
    Do While Not rsform.EOF
        For Each C In m.Controls
            If C.Name = rsform.Fields("opmNum") Then
                C.Enabled = True
                Exit For
            End If
        Next
        rsform.MoveNext
    Loop
    
    '--- Coloca invisible las opciones deshabilitadas
    If rsform.State = 1 Then rsform.Close
    csql = "EXEC spu_Usuario_OpcPerfil 2,'" & usu & "','" & glsEmpresa & "','" & StrcodSistema & "' " '"Select opmNum from opcionesmenu where opmEstado = 'N' and CodSistema = '" & StrcodSistema & "' "
    rsform.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rsform.EOF
        For Each C In m.Controls
            If C.Name = rsform.Fields("opmNum") Then
                C.Visible = False
                Exit For
            End If
        Next
        rsform.MoveNext
    Loop
    
    For Each C In m.Controls
        If left(C.Name, 3) = "mnu" Then
            If C.Enabled = False Then
                C.Visible = False
            End If
        End If
    Next
    
    'VENTAS
    ContGrupo = -1
    For Each C In m.Controls
        If left(C.Name, 5) = "mnu07" And Len(C.Name) = 7 Then
            If C.Enabled = False Or C.Visible = False Then
                intIndex = 0
                For intIndex = 0 To m.dxSideBar1.Groups(0).Links.Count - 1
                    If m.dxSideBar1.Groups(0).Links(intIndex).ObjectName = C.Name Then
                        m.dxSideBar1.Groups(0).Links.Remove (intIndex)
                        Exit For
                    End If
                Next
            End If
        End If
    Next

    'INVENTARIO
    ContGrupo = -1
    For Each C In m.Controls
        If left(C.Name, 5) = "mnu10" And Len(C.Name) = 7 Then
            If C.Enabled = False Or C.Visible = False Then
                intIndex = 0
                For intIndex = 0 To m.dxSideBar1.Groups(1).Links.Count - 1
                    If m.dxSideBar1.Groups(1).Links(intIndex).ObjectName = C.Name Then
                        m.dxSideBar1.Groups(1).Links.Remove (intIndex)
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    
    If rsform.State = 1 Then rsform.Close: Set rsform = Nothing
Exit Sub
Err:
    If rsform.State = 1 Then rsform.Close: Set rsform = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Function funcLeeConfiguracionEmpresa(pAppName As String) As Integer
On Error GoTo ErrMng

    funcLeeConfiguracionEmpresa = GetSetting(pAppName, CONF_CONEXION, CONF_INDEX_EMPRESA)

    Exit Function
    
ErrMng:
End Function

Public Sub funcGuardaConfiguracionEmpresa(pAppName As String, _
                                   indiceEmpresa As String)
               
    SaveSetting pAppName, CONF_CONEXION, CONF_INDEX_EMPRESA, indiceEmpresa
  
End Sub

Public Function EstadoCajaUsuario(ByVal strCaja As String, ByVal strFecCaja As String, ByRef StrMsgError As String) As String
On Error GoTo Err
Dim rst As New ADODB.Recordset
    
    csql = "SELECT m.indEstado " & _
            "FROM movcajas m " & _
            "WHERE m.idCaja = '" & strCaja & "' " & _
             "AND m.idUsuario = '" & glsUser & "' " & _
             "AND m.idEmpresa = '" & glsEmpresa & "' " & _
             "AND m.idSucursal = '" & glsSucursal & "' " & _
             "AND m.FecCaja = '" & Format(strFecCaja, "yyyy-mm-dd") & "' " & _
             "ORDER BY m.indEstado DESC"
             
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    If Not rst.EOF Then
        EstadoCajaUsuario = "" & rst.Fields("indEstado")
    Else
        EstadoCajaUsuario = ""
    End If

    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Function
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError <> "" Then Exit Function
    StrMsgError = Err.Description
End Function

'-- Devuelve si la caja q fue usada por el documento sigue abierta o activa
Public Function EstadoCajaDocVentas(ByVal strTD As String, ByVal strNumDoc As String, ByVal strSerie As String, ByRef StrMsgError As String) As String
On Error GoTo Err
Dim strIdMovCaja As String
Dim strEstMovCaja As String
                     
    strIdMovCaja = traerCampo("docventas", "idMovCaja", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'")
    strEstMovCaja = traerCampo("movcajas", "indEstado", "idMovCaja", strIdMovCaja, True, " idSucursal = '" & glsSucursal & "'")

    EstadoCajaDocVentas = strEstMovCaja

    Exit Function
    
Err:
    If StrMsgError <> "" Then Exit Function
    StrMsgError = Err.Description
End Function

'--- Si indTipoRpta = 0 devuelve idMovCaja
'--- Si indTipoRpta = 1 devuelve idCaja
Public Function CajaAperturadaUsuario(ByVal indTipoRpta As String, ByRef StrMsgError As String) As String
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim strTipoRpta As String
    
    strTipoRpta = "idMovCaja"
    If indTipoRpta = 1 Then strTipoRpta = "idCaja"
    
    csql = "SELECT m." & strTipoRpta & " as Resultado,m.indEstado " & _
            "FROM movcajas m " & _
            "WHERE m.idUsuario = '" & glsUser & "' " & _
             "AND m.idEmpresa = '" & glsEmpresa & "' " & _
             "AND m.idSucursal = '" & glsSucursal & "' AND indEstado = 'A' " & _
            " ORDER BY m.FecCaja DESC"
             
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    If Not rst.EOF Then
        If rst.Fields("indEstado") = "A" Then
            CajaAperturadaUsuario = "" & rst.Fields("Resultado")
        Else
            StrMsgError = "No hay caja Aperturada"
            GoTo Err
        End If
    Else
        CajaAperturadaUsuario = ""
        StrMsgError = "No hay caja Aperturada"
        GoTo Err
    End If
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Function
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Function

Public Sub cargarParametrosSistema(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
    
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    rst.Open "EXEC spu_Parametros '" & glsEmpresa & "'", Cn, adOpenStatic, adLockReadOnly

    Do While Not rst.EOF
        Select Case UCase(rst.Fields("GlsParametro"))
            Case "IGV": glsIGV = Val(rst.Fields("ValParametro")) / 100
            Case "MONEDAVENTAS": glsMonVentas = "" & rst.Fields("ValParametro")
            Case "LISTAVENTAS": glsListaVentas = "" & rst.Fields("ValParametro")
            Case "DECIMALESCAJA": glsDecimalesCaja = "" & rst.Fields("ValParametro")
            Case "DECIMALESTIPOCAMBIO": glsDecimalesTC = "" & rst.Fields("ValParametro")
            Case "VALIDASTOCK": glsValidaStock = IIf("" & rst.Fields("ValParametro") = "S", True, False)
            Case "ENTIDADSYSTEM": glsSystem = "" & rst.Fields("ValParametro")
            Case "CLIENTEVENTAS": glsClienteVentas = "" & rst.Fields("ValParametro")
            Case "FORMAPAGOVENTAS": glsFormaPagoVentas = "" & rst.Fields("ValParametro")
            Case "RUTAIMAGENPROD": glsRutaImagenProd = Replace("" & rst.Fields("ValParametro"), "*", "\")
            Case "DECIMALESPRECIOS": glsDecimalesPrecios = "" & rst.Fields("ValParametro")
            Case "INDMODVENDEDORCAMPO": glsModVendCampo = IIf("" & rst.Fields("ValParametro") = "S", True, False)
            Case "SOLOGUIAMUEVESTOCK": glsSoloGuiaMueveStock = IIf("" & rst.Fields("ValParametro") = "S", True, False)
            Case "RECEPCIONAUTO": glsRecepcionAuto = IIf("" & rst.Fields("ValParametro") = "S", True, False)
            Case "DCTOMINIMOVALIDACION": glsDctoMinValidacion = Val(rst.Fields("ValParametro"))
            Case "RUTA_ACCESS": glsRuta_Access = Trim("" & rst.Fields("ValParametro"))
            Case "ORIGEN_CONTABLE": glsOrigen_Contable = Trim("" & rst.Fields("ValParametro"))
            Case "CUENTA_IGV_VENTAS": glsCuenta_Igv_Ventas = Trim("" & rst.Fields("ValParametro"))
            Case "RUTA_ACCESS_CONTABILIDAD": glsRuta_Access_Conta = Trim("" & rst.Fields("ValParametro"))
            Case "GRABA_CONTADO": glsGraba_Contado = Trim("" & rst.Fields("ValParametro"))
            Case "TIPO_CAMBIO": glsTipoCambio = Trim("" & rst.Fields("ValParametro"))
            Case "GRABA_TODO": glsGrabaTodo = Trim("" & rst.Fields("ValParametro"))
            Case "LEE_CODIGO_BARRAS": glsLeeCodigoBarras = Trim("" & rst.Fields("ValParametro"))
            Case "DCTOMINIMO_MONTO": glsDctoMinMonto = Val(rst.Fields("ValParametro"))
            Case "VISUALIZA_CODFAB": glsVisualizaCodFab = Trim(rst.Fields("ValParametro"))
            Case "ENTERAYUDACLIENTE": glsEnterAyudaClientes = IIf("" & rst.Fields("ValParametro") = "S", True, False)
            Case "ENTERAYUDAPRODUCTOS": glsEnterAyudaProductos = IIf("" & rst.Fields("ValParametro") = "S", True, False)
            Case "GRABAR_EN_SISTEMA_ACCESS": glsSistemaAccess = Trim(rst.Fields("ValParametro"))
            Case "PORCENTAJE_RETENCION": glsPorcentajeRetencion = Val(Trim(rst.Fields("ValParametro")))
            Case "DSCTO_CON_CLAVE": glsDsctoConClave = Trim(rst.Fields("ValParametro"))
            Case "MODIFICA_PRECIO_PRODUCTO": glsModificarPrecio = Trim(rst.Fields("ValParametro"))
            Case "FORMATO_IMP_LETRA": glsFormatoImpLetra = Trim(rst.Fields("ValParametro"))
            Case "MOTIVOSALIDA":  glsMotivoSalida = Trim(rst.Fields("ValParametro") & "")
            Case "GRABA_GUIA_FACTURA":  glsGrabaGuiaFactura = Trim(rst.Fields("ValParametro") & "")
            Case "VISUALIZA_FILTRO_DOCUMENTO":  GlsVisualiza_Filtro_Documento = Trim(rst.Fields("ValParametro") & "")
            
            Case "VISUALIZA_OC_FACTURA_PEDIDO":  STR_VISUALIZA_OC_FACTURA_PEDIDO = Trim(rst.Fields("ValParametro") & "")
            Case "IMPORTAR_DOCUMENTOS_ENTRE_EMPRESAS":  STR_IMPORTAR_DOCUMENTOS_ENTRE_EMPRESAS = Trim(rst.Fields("ValParametro") & "")
            Case "EMPRESA_SUCURSAL_DOCUMENTOS_ENTRE_EMPRESAS":  STR_EMPRESA_SUCURSAL_DOCUMENTOS_ENTRE_EMPRESAS = Trim(rst.Fields("ValParametro") & "")
            Case "BLOQUEA_CHK_AFECTO":  STR_BLOQUEA_CHK_AFECTO = Trim(rst.Fields("ValParametro") & "")
            Case "VISUALIZAR_AYUDA_REQUERIMIENTO_COMPRA":  STR_VISUALIZAR_AYUDA_REQUERIMIENTO_COMPRA = Trim(rst.Fields("ValParametro") & "")
            Case "LONGITUD_IMPRESION_DETALLE_PRODUCTO":  STR_LONGITUD_IMPRESION_DETALLE_PRODUCTO = Trim(rst.Fields("ValParametro") & "")
            Case "DESCRIPCION_AREA_O_UPP":  STR_DESCRIPCION_AREA_O_UPP = Trim(rst.Fields("ValParametro") & "")
            Case "IMPORTA_ATENCIONES":  STR_IMPORTA_ATENCIONES = Trim(rst.Fields("ValParametro") & "")
            Case "APRUEBA_PEDIDO_AUTOMATICO":  STR_APRUEBA_PEDIDO_AUTOMATICO = Trim(rst.Fields("ValParametro") & "")
            Case "VENTA_ELECTRONICA":  STR_VENTA_ELECTRONICA = Trim(rst.Fields("ValParametro") & "")
            Case "STOCK_POR_LOTE":  STR_STOCK_POR_LOTE = Trim(rst.Fields("ValParametro") & "")
            Case "LIQUIDACIONES":  STR_LIQUIDACIONES = Trim(rst.Fields("ValParametro") & "")
            Case "VALIDA_SEPARACION":  STR_VALIDA_SEPARACION = Trim(rst.Fields("ValParametro") & "")
            Case "IGV":  STR_IGV = Trim(rst.Fields("ValParametro") & "")
            Case "IGV_ANT":  STR_IGV_ANT = Trim(rst.Fields("ValParametro") & "")
            Case "PERIODO_CAMBIO_IGV":  STR_PERIODO_CAMBIO_IGV = Trim(rst.Fields("ValParametro") & "")
            Case "CLIENTE_ANULA":  STR_CLIENTE_ANULA = Trim(rst.Fields("ValParametro") & "")
            
        End Select
        
        rst.MoveNext
    Loop

    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Function DataProcedimiento(ByVal lstrNombreSP As String, ByRef StrMsgError As String, ParamArray varValores() As Variant) As ADODB.Recordset
Dim cm As ADODB.Command
Dim textoSQL As String
Dim parametros As String
Dim intI As Integer
Dim RsDatos As New ADODB.Recordset
10    On Error GoTo EjecutaProcedimiento_Error

'20    Set cm = New ADODB.Command
'
'30    cm.CommandTimeout = 0
'40    cm.ActiveConnection = strcn
'50    cm.CommandType = adCmdText

60    parametros = ""

70    For intI = 0 To UBound(varValores)
          If parametros = "" Then
            parametros = "'" & varValores(intI) & "'"
          Else
            parametros = parametros & ",'" & varValores(intI) & "'"
          End If
90    Next

      textoSQL = "EXECUTE " & lstrNombreSP & parametros
      
      
     If RsDatos.State = 1 Then RsDatos.Close: Set RsDatos = Nothing
     RsDatos.Open textoSQL, Cn, adOpenStatic, adLockOptimistic
        
100   Set DataProcedimiento = RsDatos

120   On Error GoTo 0
130   Exit Function
EjecutaProcedimiento_Error:
Err:
140   If StrMsgError = "" Then StrMsgError = Err.Description
      Set cm.ActiveConnection = Nothing
150   Set cm = Nothing
End Function

Public Function traerDireccionSucursal() As String
Dim rst As New ADODB.Recordset

    csql = "SELECT concat(p.direccion,' ',u.glsUbigeo,' ',d.glsUbigeo) as direccion " & _
            "FROM personas p,ubigeo u,ubigeo d " & _
            "Where P.idDistrito = u.idDistrito " & _
            "AND left(u.idDistrito,2) = d.idDpto " & _
            "AND d.idProv = '00' " & _
            "AND d.idDist = '00' AND p.idPersona = '" & glsSucursal & "'"
                
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF Then
        traerDireccionSucursal = "" & rst.Fields("direccion")
    End If
    rst.Close: Set rst = Nothing
        
End Function

Public Sub ConfiguracionDecimal()
Dim blnExisteComaDecimal As Boolean

    blnExisteComaDecimal = InStr(CStr(10.5), ",") > 0
    If blnExisteComaDecimal Then
        MsgBox "Configurar el Simbolo Decimal, para seguir trabajando ", vbInformation, App.Title
        If Cn.State = 1 Then Cn.Close
        Set Cn = Nothing
        End
    End If
    
End Sub

Public Sub getEstadoCierreMes(ByVal fec As Date, ByRef StrMsgError As String)
On Error GoTo Err
Dim strEstado As String
Dim strVarAno As String
Dim strVarMes As String

    strVarAno = Format(Year(fec), "0000")
    strVarMes = Format(Month(fec), "00")
    strEstado = traerCampo("cierresmes", "estCierre", "idAno", strVarAno, True, "idMes = '" & strVarMes & "' And IdSistema = '21001'")
    
    If strEstado = "" Or strEstado = "A" Then
        '--- ABIERTO
        StrMsgError = ""
        Exit Sub
    Else
        '--- CERRADO
        StrMsgError = "El mes se encuentra cerrado"
        GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub actualizaStock_Lote(ByVal strCodValesCab As String, ByVal indTipoOpe As Integer, ByRef StrMsgError As String, ByRef StrTipVale As String, Optional ByVal IndUsarTrans As Boolean = True, Optional strOptCodSucursal As String = "")
On Error GoTo Err
Dim rs As New ADODB.Recordset
Dim strTipoVale As String
Dim strSigno As String
Dim strCodAlmacen As String
Dim strCodMoneda As String
Dim indTrans As Boolean
Dim strCodConcepto As String
Dim strCosto As String
Dim rsTempo     As New ADODB.Recordset
Dim strVarCodSucursal As String

    strVarCodSucursal = glsSucursal
    If strOptCodSucursal <> "" Then strVarCodSucursal = strOptCodSucursal
    
    indTrans = False
    csql = ""
    
    If IndUsarTrans Then
        indTrans = True
        Cn.BeginTrans
    End If
    
    '--- Buscamos datos del vale
    csql = "SELECT tipoVale, idAlmacen, idMoneda, idConcepto FROM valescab WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & strVarCodSucursal & "' AND idValesCab = '" & strCodValesCab & "' AND tipoVale = '" & StrTipVale & "' "
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        strTipoVale = "" & rs.Fields("tipoVale")
        strCodAlmacen = "" & rs.Fields("idAlmacen")
        strCodMoneda = "" & rs.Fields("idMoneda")
        strCodConcepto = "" & rs.Fields("idConcepto")
    Else
        StrMsgError = "El Vale no se encuentra registrado"
        GoTo Err
    End If
    
    strSigno = "+"
    If strTipoVale = "S" Then strSigno = "-"
    
    If indTipoOpe = 1 Then 'modifica
        If strSigno = "+" Then
            strSigno = "-"
        Else
            strSigno = "+"
        End If
    End If
    
    '--- Actualizamos stock
    csql = "Select v.idEmpresa,v.idSucursal,v.idProducto,v.idUM,v.idLote " & _
            ",v.Cantidad " & _
            "From productosalmacenporlote s, valesdet v " & _
            "WHERE s.idEmpresa = v.idEmpresa " & _
              "AND s.idSucursal = v.idSucursal " & _
              "AND s.idAlmacen = '" & strCodAlmacen & "' " & _
              "AND s.idProducto = v.idProducto " & _
              "AND s.idUMCompra = v.idUM " & _
              "AND s.idLote = v.idLote " & _
              "AND s.idEmpresa = '" & glsEmpresa & "' " & _
              "AND s.idSucursal = '" & glsSucursal & "' " & _
              "AND v.idValesCab = '" & strCodValesCab & "' " & _
              "AND v.tipoVale = '" & StrTipVale & "' "
              
    If rsTempo.State = 1 Then rsTempo.Close
    Set rsTempo = Nothing
    
    rsTempo.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsTempo.EOF Then
        Do While Not rsTempo.EOF
            csql = "UPDATE productosalmacenporlote  " & _
                    "SET CantidadStock = CantidadStock " & strSigno & " " & rsTempo.Fields("cantidad") & " " & _
                    "WHERE idEmpresa = '" & rsTempo.Fields("idEmpresa") & "' " & _
                      "AND idSucursal = '" & rsTempo.Fields("idSucursal") & "' " & _
                      "AND idAlmacen = '" & strCodAlmacen & "' " & _
                      "AND idProducto = '" & rsTempo.Fields("idProducto") & "' " & _
                      "AND idUMCompra = '" & rsTempo.Fields("idUM") & "' " & _
                      "AND idLote = '" & rsTempo.Fields("idLote") & "' "
    
            Cn.Execute csql
            
            rsTempo.MoveNext
        Loop
    End If
            
    If IndUsarTrans Then Cn.CommitTrans
    
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    
    Exit Sub
    
Err:
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If IndUsarTrans Then
        If indTrans Then Cn.RollbackTrans
    End If
End Sub

Public Sub actualizaStock_LoteTransferencia(ByVal strCodValesCab As String, ByVal indTipoOpe As Integer, ByVal SucursalDestino As String, ByRef StrMsgError As String, ByRef StrTipVale As String, Optional ByVal IndUsarTrans As Boolean = True, Optional strOptCodSucursal As String = "")
On Error GoTo Err
Dim rs As New ADODB.Recordset
Dim strTipoVale As String
Dim strSigno As String
Dim strCodAlmacen As String
Dim strCodMoneda As String
Dim indTrans As Boolean
Dim strCodConcepto As String
Dim strCosto As String
Dim rsTempo     As New ADODB.Recordset
Dim strVarCodSucursal As String

    strVarCodSucursal = SucursalDestino
    If strOptCodSucursal <> "" Then strVarCodSucursal = strOptCodSucursal
    
    indTrans = False
    csql = ""
    
    If IndUsarTrans Then
        indTrans = True
        Cn.BeginTrans
    End If
    
    '--- Buscamos datos del vale
    csql = "SELECT tipoVale, idAlmacen, idMoneda, idConcepto FROM valescab WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & strVarCodSucursal & "' AND idValesCab = '" & strCodValesCab & "' AND tipoVale = '" & StrTipVale & "' "
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        strTipoVale = "" & rs.Fields("tipoVale")
        strCodAlmacen = "" & rs.Fields("idAlmacen")
        strCodMoneda = "" & rs.Fields("idMoneda")
        strCodConcepto = "" & rs.Fields("idConcepto")
    Else
        StrMsgError = "El Vale no se encuentra registrado"
        GoTo Err
    End If
    
    strSigno = "+"
    If strTipoVale = "S" Then strSigno = "-"
    
    If indTipoOpe = 1 Then 'modifica
        If strSigno = "+" Then
            strSigno = "-"
        Else
            strSigno = "+"
        End If
    End If

    '--- Actualizamos stock
    csql = "Select v.idEmpresa,v.idSucursal,v.idProducto,v.idUM,v.idLote " & _
            ",v.Cantidad " & _
            "From productosalmacenporlote s, valesdet v " & _
            "WHERE s.idEmpresa = v.idEmpresa " & _
              "AND s.idSucursal = v.idSucursal " & _
              "AND s.idAlmacen = '" & strCodAlmacen & "' " & _
              "AND s.idProducto = v.idProducto " & _
              "AND s.idUMCompra = v.idUM " & _
              "AND s.idLote = v.idLote " & _
              "AND s.idEmpresa = '" & glsEmpresa & "' " & _
              "AND s.idSucursal = '" & glsSucursal & "' " & _
              "AND v.idValesCab = '" & strCodValesCab & "' " & _
              "AND v.tipoVale = '" & StrTipVale & "' "
              
    If rsTempo.State = 1 Then rsTempo.Close: Set rsTempo = Nothing
    
    rsTempo.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsTempo.EOF Then
        Do While Not rsTempo.EOF
            csql = "UPDATE productosalmacenporlote  " & _
                    "SET CantidadStock = CantidadStock " & strSigno & " " & rsTempo.Fields("cantidad") & " " & _
                    "WHERE idEmpresa = '" & rsTempo.Fields("idEmpresa") & "' " & _
                      "AND idSucursal = '" & rsTempo.Fields("idSucursal") & "' " & _
                      "AND idAlmacen = '" & strCodAlmacen & "' " & _
                      "AND idProducto = '" & rsTempo.Fields("idProducto") & "' " & _
                      "AND idUMCompra = '" & rsTempo.Fields("idUM") & "' " & _
                      "AND idLote = '" & rsTempo.Fields("idLote") & "' "
            Cn.Execute csql
            
            rsTempo.MoveNext
        Loop
    End If
            
    If IndUsarTrans Then Cn.CommitTrans
    
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    
    Exit Sub
    
Err:
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If IndUsarTrans Then
        If indTrans Then Cn.RollbackTrans
    End If
End Sub

'indTipoOpe   0 = Inserta       1 = Modifica
Public Sub actualizaStock(ByVal strCodValesCab As String, ByVal indTipoOpe As Integer, ByRef StrMsgError As String, ByRef StrTipVale As String, Optional ByVal IndUsarTrans As Boolean = True, Optional strOptCodSucursal As String = "")
On Error GoTo Err
Dim rs As New ADODB.Recordset
Dim strTipoVale As String
Dim strSigno As String
Dim strCodAlmacen As String
Dim strCodMoneda As String
Dim indTrans As Boolean
Dim strCodConcepto As String
Dim strCosto As String
Dim strVarCodSucursal As String
Dim RsC As New ADODB.Recordset
Dim strCodProducto As String
Dim strCodUM As String
        
    strVarCodSucursal = glsSucursal
    If strOptCodSucursal <> "" Then strVarCodSucursal = strOptCodSucursal
    
    indTrans = False
    csql = ""
    
    If IndUsarTrans Then
        indTrans = True
        Cn.BeginTrans
    End If
    
    '--- Buscamos datos del vale
    csql = "SELECT tipoVale, idAlmacen, idMoneda, idConcepto FROM valescab WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & strVarCodSucursal & "' AND idValesCab = '" & strCodValesCab & "' AND tipoVale = '" & StrTipVale & "' "
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        strTipoVale = "" & rs.Fields("tipoVale")
        strCodAlmacen = "" & rs.Fields("idAlmacen")
        strCodMoneda = "" & rs.Fields("idMoneda")
        strCodConcepto = "" & rs.Fields("idConcepto")
    Else
        StrMsgError = "El Vale no se encuentra registrado"
        GoTo Err
    End If
    
    strSigno = "+"
    If strTipoVale = "S" Then strSigno = "-"
    
    If indTipoOpe = 1 Then '--- Modifica
        If strSigno = "+" Then
            strSigno = "-"
        Else
            strSigno = "+"
        End If
    End If
        
    csql = "UPDATE P Set p.CantidadStock = (p.CantidadStock " & strSigno & " v.Cantidad) FROM productosalmacen p " & _
             "INNER JOIN (Select SUM(vd.Cantidad) as  Cantidad,vd.idProducto " & _
             "FROM valesDet vd " & _
             "INNER JOIN ValesCab vc ON vd.idEmpresa = vc.idEmpresa AND vd.idSucursal = vc.idSucursal AND vd.TipoVale = vc.TipoVale AND vd.idValesCab = vc.idValesCab " & _
             "WHERE  vc.idValesCab = '" & strCodValesCab & "' AND vc.idEmpresa = '" & glsEmpresa & "' AND vc.idSucursal = '" & strVarCodSucursal & "' AND vc.tipoVale = '" & StrTipVale & "' AND vc.idAlmacen = '" & strCodAlmacen & "' GROUP BY  vd.IdProducto) v " & _
             "ON p.idProducto = v.idProducto " & _
             " " & _
             " " & _
             " " & _
             " " & _
             "WHERE p.idEmpresa = '" & glsEmpresa & "'  " & _
             "AND p.idSucursal = '" & strVarCodSucursal & "'  " & _
             "" & _
             "" & _
             "AND p.idAlmacen = '" & strCodAlmacen & "' "
    Cn.Execute csql
    
    '--- Actualizamos Costo Unitario
    strCosto = ""
    If strCodConcepto = "05" Then '--- Si es compra actualizamos costo
        csql = "SELECT v.idProducto, v.idUM FROM valesdet v " & _
                "WHERE v.idEmpresa = '" & glsEmpresa & "' " & _
                  "AND v.idSucursal = '" & strVarCodSucursal & "' " & _
                  "AND v.idValesCab = '" & strCodValesCab & "' " & _
                  "AND v.tipoVale = '" & StrTipVale & "' "
        RsC.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        
        Do While Not RsC.EOF
            strCodProducto = "" & RsC.Fields("idProducto")
            strCodUM = "" & RsC.Fields("idUM")
            
            '--- Calculo
            costoUnitario strVarCodSucursal, strCodProducto, strCodAlmacen, strCodUM, strCosto, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            '--- Actualizo
            csql = "UPDATE productosalmacen s " & _
                "SET s.CostoUnit = " & strCosto & " " & _
                "WHERE s.idEmpresa = '" & glsEmpresa & "' " & _
                  "AND s.idSucursal = '" & strVarCodSucursal & "' " & _
                  "AND s.idAlmacen = '" & strCodAlmacen & "' " & _
                  "AND s.idProducto = '" & strCodProducto & "' " & _
                  "AND s.idUMCompra = '" & strCodUM & "' "
            Cn.Execute csql
            
            RsC.MoveNext
        Loop
    End If
    If IndUsarTrans Then Cn.CommitTrans
    
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    
    Exit Sub
    
Err:
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If IndUsarTrans Then
        If indTrans Then Cn.RollbackTrans
    End If
    
End Sub

'--- indTipoOpe   0 = Inserta       1 = Modifica
Public Sub actualizaStockTransferencia(ByVal strCodValesCab As String, ByVal indTipoOpe As Integer, ByVal SucursalDest As String, ByRef StrMsgError As String, ByRef StrTipVale As String, Optional ByVal IndUsarTrans As Boolean = True, Optional strOptCodSucursal As String = "")
On Error GoTo Err
Dim rs As New ADODB.Recordset
Dim strTipoVale As String
Dim strSigno As String
Dim strCodAlmacen As String
Dim strCodMoneda As String
Dim indTrans As Boolean
Dim strCodConcepto As String
Dim strCosto As String
Dim strVarCodSucursal As String

    strVarCodSucursal = SucursalDest
    If strOptCodSucursal <> "" Then strVarCodSucursal = strOptCodSucursal
    
    indTrans = False
    csql = ""
    If IndUsarTrans Then
        indTrans = True
        Cn.BeginTrans
    End If
    
    '--- Buscamos datos del vale
    csql = "SELECT tipoVale, idAlmacen, idMoneda, idConcepto FROM valescab WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & strVarCodSucursal & "' AND idValesCab = '" & strCodValesCab & "' AND tipoVale = '" & StrTipVale & "' "
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        strTipoVale = "" & rs.Fields("tipoVale")
        strCodAlmacen = "" & rs.Fields("idAlmacen")
        strCodMoneda = "" & rs.Fields("idMoneda")
        strCodConcepto = "" & rs.Fields("idConcepto")
    Else
        StrMsgError = "El Vale no se encuentra registrado"
        GoTo Err
    End If
    
    strSigno = "+"
    If strTipoVale = "S" Then strSigno = "-"
    
    If indTipoOpe = 1 Then 'modifica
        If strSigno = "+" Then
            strSigno = "-"
        Else
            strSigno = "+"
        End If
    End If
    
    csql = "UPDATE productosalmacen p " & _
             "INNER JOIN (Select SUM(vd.Cantidad) as  Cantidad,vc.idValesCab,vc.tipoVale,vc.idEmpresa,vc.idSucursal,vd.idProducto,vd.idUM " & _
             "FROM valesDet vd " & _
             "INNER JOIN ValesCab vc ON vd.idEmpresa = vc.idEmpresa AND vd.idSucursal = vc.idSucursal AND vd.TipoVale = vc.TipoVale AND vd.idValesCab = vc.idValesCab " & _
             "WHERE  vc.idValesCab = '" & strCodValesCab & "' AND vc.idEmpresa = '" & glsEmpresa & "' AND vc.idSucursal = '" & strVarCodSucursal & "' AND vc.tipoVale = '" & StrTipVale & "' AND vc.idAlmacen = '" & strCodAlmacen & "' GROUP BY  vd.IdProducto) v " & _
             "ON p.idEmpresa = v.idEmpresa " & _
             "AND p.idSucursal = v.idSucursal " & _
             "AND p.idProducto = v.idProducto " & _
             "AND p.idUMCompra = v.idUM " & _
             "Set p.CantidadStock = (p.CantidadStock " & strSigno & " v.Cantidad) " & _
             "WHERE p.idEmpresa = '" & glsEmpresa & "'  " & _
             "AND p.idSucursal = '" & strVarCodSucursal & "'  " & _
             "AND v.idValesCab ='" & strCodValesCab & "' " & _
             "AND v.tipoVale = '" & StrTipVale & "'  " & _
             "AND p.idAlmacen = '" & strCodAlmacen & "' "
    Cn.Execute csql
    
    '--- Actualizamos Costo Unitario
    strCosto = ""
    If strCodConcepto = "05" Then '--- Si es compra actualizamos costo
        Dim RsC As New ADODB.Recordset
        Dim strCodProducto As String
        Dim strCodUM As String
        
        csql = "SELECT v.idProducto, v.idUM FROM valesdet v " & _
                "WHERE v.idEmpresa = '" & glsEmpresa & "' " & _
                  "AND v.idSucursal = '" & strVarCodSucursal & "' " & _
                  "AND v.idValesCab = '" & strCodValesCab & "' " & _
                  "AND v.tipoVale = '" & StrTipVale & "' "
        RsC.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        
        Do While Not RsC.EOF
            strCodProducto = "" & RsC.Fields("idProducto")
            strCodUM = "" & RsC.Fields("idUM")
            
            '--- Calculo
            costoUnitario strVarCodSucursal, strCodProducto, strCodAlmacen, strCodUM, strCosto, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            '--- Actualizo
            csql = "UPDATE productosalmacen s " & _
                "SET s.CostoUnit = " & strCosto & " " & _
                "WHERE s.idEmpresa = '" & glsEmpresa & "' " & _
                  "AND s.idSucursal = '" & strVarCodSucursal & "' " & _
                  "AND s.idAlmacen = '" & strCodAlmacen & "' " & _
                  "AND s.idProducto = '" & strCodProducto & "' " & _
                  "AND s.idUMCompra = '" & strCodUM & "' "
            Cn.Execute csql
            
            RsC.MoveNext
        Loop
    End If
            
    If IndUsarTrans Then Cn.CommitTrans
    
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    
    Exit Sub
    
Err:
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If IndUsarTrans Then
        If indTrans Then Cn.RollbackTrans
    End If
End Sub

Public Sub actualizaStock_Liquidaciones(ByVal strCodValesCab As String, ByVal indTipoOpe As Integer, ByRef StrMsgError As String, ByRef StrTipVale As String, ByRef StrSucursalx As String, Optional ByVal IndUsarTrans As Boolean = True, Optional strOptCodSucursal As String = "")
On Error GoTo Err
Dim rs As New ADODB.Recordset
Dim strTipoVale As String
Dim strSigno As String
Dim strCodAlmacen As String
Dim strCodMoneda As String
Dim indTrans As Boolean
Dim strCodConcepto As String
Dim strCosto As String
Dim strVarCodSucursal As String

    strVarCodSucursal = glsSucursal
    If strOptCodSucursal <> "" Then strVarCodSucursal = strOptCodSucursal
    
    indTrans = False
    csql = ""
    
    If IndUsarTrans Then
        indTrans = True
        Cn.BeginTrans
    End If
    
    '--- Buscamos datos del vale
    csql = "SELECT tipoVale, idAlmacen, idMoneda, idConcepto FROM valescab WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & StrSucursalx & "' AND idValesCab = '" & strCodValesCab & "' AND tipoVale = '" & StrTipVale & "' "
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        strTipoVale = "" & rs.Fields("tipoVale")
        strCodAlmacen = "" & rs.Fields("idAlmacen")
        strCodMoneda = "" & rs.Fields("idMoneda")
        strCodConcepto = "" & rs.Fields("idConcepto")
    Else
        'StrMsgError = "El Vale no se encuentra registrado."
        'GoTo Err
    End If
    
    strSigno = "+"
    If strTipoVale = "S" Then strSigno = "-"
    
    If indTipoOpe = 1 Then '--- Modifica
        If strSigno = "+" Then
            strSigno = "-"
        Else
            strSigno = "+"
        End If
    End If
    
    '--- Actualizamos stock
    csql = "UPDATE productosalmacen s, valesdet v " & _
            "SET s.CantidadStock = s.CantidadStock " & strSigno & " v.Cantidad " & _
            "WHERE s.idEmpresa = v.idEmpresa " & _
              "AND s.idSucursal = v.idSucursal " & _
              "AND s.idAlmacen = '" & strCodAlmacen & "' " & _
              "AND s.idProducto = v.idProducto " & _
              "AND s.idUMCompra = v.idUM " & _
              "AND s.idEmpresa = '" & glsEmpresa & "'" & _
              "AND s.idSucursal = '" & StrSucursalx & "'" & _
              "AND v.idValesCab = '" & strCodValesCab & "' " & _
              "AND v.tipoVale = '" & StrTipVale & "' "
    Cn.Execute csql
            
    If IndUsarTrans Then Cn.CommitTrans
    
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    
    Exit Sub
    
Err:
    If rs.State = 1 Then rs.Close: Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If IndUsarTrans Then
        If indTrans Then Cn.RollbackTrans
    End If
End Sub

Private Sub costoUnitario(ByVal strVarCodSucursal As String, ByVal codproducto As String, ByVal codalmacen As String, ByVal idUM As String, ByRef strCosUni As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsu As New ADODB.Recordset

    csql = "SELECT AVG(valesdet.VVUnit) AS CostoUnit "
    csql = csql & "FROM valescab,valesdet "
    csql = csql & "WHERE valescab.idValesCab = valesdet.idValesCab AND "
    csql = csql & "valescab.idEmpresa = valesdet.idEmpresa AND "
    csql = csql & "valescab.idSucursal = valesdet.idSucursal AND "
    csql = csql & "valescab.tipoVale = valesdet.tipoVale AND "
    csql = csql & "valescab.idEmpresa = '" & glsEmpresa & "' AND "
    csql = csql & "valescab.idSucursal = '" & strVarCodSucursal & "' AND valescab.idConcepto = '05' AND "
    csql = csql & "valescab.idPeriodoInv = '" & glsCodPeriodoINV & "' And valesdet.idProducto = '" & codproducto & "' AND "
    csql = csql & "valescab.idAlmacen = '" & codalmacen & "' AND valesdet.idUM = '" & idUM & "' "
    
    rsu.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsu.EOF Then
       strCosUni = IIf(IsNull(rsu.Fields("CostoUnit")), 0, rsu.Fields("CostoUnit"))
    End If

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub mostrarDatosGridSQL2(g As dxDBGrid, r As Recordset, ByRef ColumnaClave As String, ByRef StrMsgError As String)
On Error GoTo Err

    With g
        .DefaultFields = False
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        Set .DataSource = r
        .Dataset.Active = True
        .KeyField = ColumnaClave
        .Dataset.Edit
        .Dataset.Post
    End With

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Function MonedaTexto(numero As Double, TMoneda As String) As String

'Private ANums:={{" un"," diez"," cien"},{" dos"," veinte"," docientos"},{" tres"," treinta"," trecientos"},{" cuatro"," cuarenta"," cuatrocientos"},{" cinco"," cincuenta"," quinientos"},{" seis"," sesenta", "seiscientos"},{" siete"," setenta"," setecientos"},{" ocho"," ochenta"," ochocientos"},{" nueve"," noventa"," novecientos"}}
'Private ADec :={"diez"," once"," doce"," trece"," catorce"," quince"," diecis‚is"," diecisiete"," dieciocho"," diecinueve"}
'Private AComa:={" mil"," millones"," mil"," billones"," mil"," trillones"," mil"," cuatrillones"," quintillones"," sextillones"}

    ReDim ANums(9, 3)
    ReDim ADec(10)
    ReDim AComa(10)
    
    ANums(1, 1) = " un"
    ANums(1, 2) = " diez"
    ANums(1, 3) = " cien"
    ANums(2, 1) = " dos"
    ANums(2, 2) = " veinte"
    ANums(2, 3) = " docientos"
    ANums(3, 1) = " tres"
    ANums(3, 2) = " treinta"
    ANums(3, 3) = " trecientos"
    ANums(4, 1) = " cuatro"
    ANums(4, 2) = " cuarenta"
    ANums(4, 3) = " cuatrocientos"
    ANums(5, 1) = " cinco"
    ANums(5, 2) = " cincuenta"
    ANums(5, 3) = " quinientos"
    ANums(6, 1) = " seis"
    ANums(6, 2) = " sesenta"
    ANums(6, 3) = " seiscientos"
    ANums(7, 1) = " siete"
    ANums(7, 2) = " setenta"
    ANums(7, 3) = " setecientos"
    ANums(8, 1) = " ocho"
    ANums(8, 2) = " ochenta"
    ANums(8, 3) = " ochocientos"
    ANums(9, 1) = " nueve"
    ANums(9, 2) = " noventa"
    ANums(9, 3) = " novecientos"
    
    ADec(1) = "diez"
    ADec(2) = " once"
    ADec(3) = " doce"
    ADec(4) = " trece"
    ADec(5) = " catorce"
    ADec(6) = " quince"
    ADec(7) = " dieciseis"
    ADec(8) = " diecisiete"
    ADec(9) = " dieciocho"
    ADec(10) = " diecinueve"
    
    AComa(1) = " mil"
    AComa(2) = " millones"
    AComa(3) = " mil"
    AComa(4) = " billones"
    AComa(5) = " mil"
    AComa(6) = " trillones"
    AComa(7) = " mil"
    AComa(8) = " cuatrillones"
    AComa(9) = " quintillones"
    AComa(10) = " sextillones"
    
    xResult = IIf(numero = 0, "cero", "")
    NumC = Format(numero, "#################.00")
    X = Len(NumC)
    NumE = Mid(NumC, 1, X - 3)
    NumD = Mid(NumC, X - 1, 2)
    X = Len(NumE)
    Veces = IIf((X Mod 3) > 0, Int(X / 3) + 1, Int(X / 3))
    
    NumCad = right(String(3 * Veces, "0") & Trim(Val(NumE)), 3 * Veces)
    
    X = Len(NumCad)
    For i = 1 To Veces
        SubCad = Mid(NumCad, (X - (i * 3)) + 1, 3)
        xText = ""
        For j = 1 To 3
         xDig = Val(Mid(SubCad, j, 1))
         xSig = Val(Mid(SubCad, j + 1, 1))
            If j = 2 And xDig = 1 Then
                xText = xText + IIf(xDig > 0, ADec(xSig + 1), "")
                j = 3
            Else
                xText = xText + IIf(xDig > 0, ANums(xDig, 4 - j) + IIf(xSig > 0 And j = 2, " y", ""), "")
                xText = IIf(j = 1 And xDig = 1 And Val(Mid(SubCad, j + 1, 2)) <> 0, xText + "to", xText)
            End If
        Next
            xText = IIf(X > i * 3, AComa(i) + " " + xText, xText)
            xResult = xText + xResult
    Next
    
    strMoneda = ""
    
    Select Case TMoneda
        Case "0": strMoneda = "Nuevos Soles"
        Case "1": strMoneda = "Dolares Americanos"
    End Select
    
    MonedaTexto = xResult + IIf(X > 1 And right(NumCad, 1) = "1" And right(NumCad, 2) <> "11", "o ", "") + " con " + NumD + "/100 " + strMoneda

End Function

Public Function generaCorrelativoAnoMes_Vale(strTabla As String, strCod As String, StrTipVale As String, Optional indEmpresa As Boolean = True)
Dim rst As New ADODB.Recordset
Dim dateSys As Date
Dim strCond As String
Dim csql As String

    dateSys = getFechaSistema
    strCond = right(CStr(Year(dateSys)), 2) & Format(Month(dateSys), "00")
    csql = "SELECT " & strCod & " FROM " & strTabla & " WHERE left(" & strCod & ",4) = '" & strCond & "' " & _
           "and tipoVale = '" & StrTipVale & "' "

    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    csql = csql & " ORDER BY 1 DESC"
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        generaCorrelativoAnoMes_Vale = strCond & Format((Val(right("" & rst.Fields(0), 4)) + 1), "0000")
    Else
        generaCorrelativoAnoMes_Vale = strCond & "0001"
    End If
    rst.Close: Set rst = Nothing

End Function

Public Sub ConfGrid1(g As dxDBGrid, indMod As Boolean, Optional mostrarFooter As Boolean, Optional mostrarGroupPanel As Boolean, Optional mostrarBandas As Boolean)
    
    With g.Options
        If indMod Then
            .Set (egoEditing)
            .Set (egoCanDelete)
            .Set (egoCanInsert)
        End If
        If mostrarBandas Then .Set (egoShowBands)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        '.Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoNameCaseInsensitive)
        If mostrarFooter Then .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
       
        If mostrarGroupPanel Then .Set (egoShowGroupPanel)
        .Set (egoEnableNodeDragging)
        .Set (egoDragCollapse)
        .Set (egoDragExpand)
        .Set (egoDragScroll)
        .Set (egoEnableNodeDragging)
    End With
   
    g.Filter.DropDownCount = 30
   
End Sub

Public Sub EjecutaSQLFormTrans(F As Form, tipoOperacion As Integer, indEmpresa As Boolean, strTabla As String, ByRef StrMsgError As String, IndUsarTrans As Boolean, Optional strCampoCod As String, Optional g As dxDBGrid, Optional strTablaDet As String, Optional strCampoDet As String, Optional strDataCampo As String, Optional indFechaRegistro As Boolean = False)
On Error GoTo Err
Dim C                                               As Object
Dim CSqlC                                           As String
Dim strCampo                                        As String
Dim strTipoDato                                     As String
Dim strCampos                                       As String
Dim strValores                                      As String
Dim strValCod                                       As String
Dim strCampoEmpresa                                 As String
Dim strValorEmpresa                                 As String
Dim strCondEmpresa                                  As String
Dim strCampoFecReg                                  As String
Dim strValorFecReg                                  As String
Dim GlsObsCli                                       As String
Dim indTrans                                        As Boolean

    If indEmpresa Then
        strCampoEmpresa = ",idEmpresa"
        strValorEmpresa = ",'" & glsEmpresa & "'"
        strCondEmpresa = " AND idEmpresa = '" & glsEmpresa & "'"
    End If

    If indFechaRegistro Then
        strCampoFecReg = ",FecRegistro"
        strValorFecReg = ",getdate()"
    End If

    indTrans = False
    CSqlC = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & C.Value & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    Case 1
                        If UCase(strCampoCod) <> UCase(strCampo) Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = C.Value
                                Case "T"
                                    strValores = "'" & C.Value & "'"
                                Case "F"
                                    strValores = "'" & Format(C.Value, "yyyy-mm-dd") & "'"
                            End Select
                            strCampos = strCampos & strCampo & "=" & strValores & ","
                        Else
                            strValCod = C.Value
                        End If
                End Select
            End If
        End If
    Next

    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
    
    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            CSqlC = "Insert Into " & strTabla & "(" & strCampos & strCampoEmpresa & strCampoFecReg & ")Values(" & strValores & strValorEmpresa & strValorFecReg & ")"
        Case 1
            CSqlC = "Update " & strTabla & " Set " & strCampos & " Where " & strCampoCod & " = '" & strValCod & "'" & strCondEmpresa
    End Select
    
    If IndUsarTrans Then
        indTrans = True
        Cn.BeginTrans
    End If

    '--- Graba controles
    If strCampos <> "" Then
        Cn.Execute CSqlC
    End If

    GlsObsCli = "" & glsobservacioncliente
    If UCase(strTabla) = "CLIENTES" Then
        CSqlC = "Update Clientes Set GlsObservacion = '" & GlsObsCli & "' Where Idcliente ='" & strDataCampo & "' And idEmpresa= '" & glsEmpresa & "' "
        Cn.Execute (CSqlC)
               
    ElseIf UCase(strTabla) = "PROVEEDORES" Then
        CSqlC = "Update proveedores Set GlsObservacion = '" & GlsObsCli & "' Where Idproveedor ='" & strDataCampo & "' And idEmpresa= '" & glsEmpresa & "' "
        Cn.Execute (CSqlC)
    End If

    '--- Grabando Grilla
    If TypeName(g) <> "Nothing" Then
    
        Cn.Execute "Delete From " & strTablaDet & " Where " & strCampoDet & " = '" & strDataCampo & "'" & strCondEmpresa
        
        g.Dataset.First
        Do While Not g.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To g.Columns.Count - 1
                If UCase(left(g.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(g.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(g.Columns(i).ObjectName, 3)
                    strCampos = strCampos & strCampo & ","
                    
                    Select Case strTipoDato
                        Case "N"
                            strValores = strValores & g.Columns(i).Value & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            CSqlC = "Insert Into " & strTablaDet & "(" & strCampos & "," & strCampoDet & strCampoEmpresa & ")Values(" & strValores & ",'" & strDataCampo & "'" & strValorEmpresa & ")"
            Cn.Execute CSqlC
            
            g.Dataset.Next
            
        Loop
    End If
    
    If IndUsarTrans Then
        Cn.CommitTrans
    End If
    
    Exit Sub
    
Err:
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub Centrar_Imagen(Objeto As Object, Path_Imagen As String)
On Error GoTo ErrSub
Dim Pos_x As Single 'Posición horizontal de la imagen
Dim Pos_y As Single 'Posición vertical de la imagen
Dim Ancho_IMG As Single 'Ancho de la imagen
Dim Alto_IMG As Single 'Alto de la imagen
Dim Ancho_Obj As Single 'Ancho en Pixeles del objeto contenedor
Dim Alto_Obj As Single 'Alto en Pixeles del objeto contenedor
Dim Old_Scale As Single 'Para lamacenar el ScaleMode del objeto
  
    Static old_Path As String
       
    If old_Path <> Path_Imagen Then
       Set Pic = LoadPicture(Path_Imagen)
    End If
       
    With Objeto
        .AutoRedraw = True
        .Cls
        Old_Scale = .ScaleMode
        .ScaleMode = vbPixels
       
        '--- Pasa el ancho y alto de la imagen a Pixeles
        Ancho_IMG = .ScaleX(Pic.Width, vbHimetric, vbPixels)
        Alto_IMG = .ScaleY(Pic.Height, vbHimetric, vbPixels)
        
        Ancho_Obj = .ScaleWidth
        Alto_Obj = .ScaleHeight
        
        If Ancho_IMG > Ancho_Obj Then
            Alto_IMG = Alto_IMG * Ancho_Obj / Ancho_IMG
            Ancho_IMG = Ancho_Obj
        End If
            If Alto_IMG > Alto_Obj Then
            Ancho_IMG = Ancho_IMG * Alto_Obj / Alto_IMG
            Alto_IMG = Alto_Obj
        End If
        
        '--- Posición X e Y donde dibujar con PaintPicture
        Pos_x = (Ancho_Obj - Ancho_IMG) / 2
        Pos_y = (Alto_Obj - Alto_IMG) / 2
    End With
       
    '--- Dibuja la imagen
    Objeto.PaintPicture Pic, Pos_x, Pos_y, Ancho_IMG, Alto_IMG
       
    '--- Restaura el ScaleMode
    Objeto.ScaleMode = Old_Scale
    old_Path = Path_Imagen
       
    Exit Sub

ErrSub:
    MsgBox Err.Description, vbCritical
End Sub

Public Property Get ComputerName() As String
Dim sName As String
Dim lRetval As Long
Dim iPos As Integer
    
    sName = Space$(255)
    lRetval = GetComputerName(sName, 255)
    iPos = InStr(sName, Chr$(0))
    ComputerName = left$(sName, iPos - 1)
    
End Property

Public Function GeneraCorrelativoAnoMesNuevo(strTabla As String, strCod As String, PPrefijo As String, Optional indEmpresa As Boolean = True)
Dim rst As New ADODB.Recordset
Dim dateSys As Date
Dim strCond As String
Dim csql As String

    dateSys = getFechaSistema
    strCond = right(PPrefijo, 4) 'Right(CStr(Year(dateSys)), 2) & Format(Month(dateSys), "00")
    csql = "SELECT " & strCod & " FROM " & strTabla & " WHERE left(" & strCod & ",4) = '" & strCond & "' "
    
    If indEmpresa Then
        csql = csql & " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    csql = csql & " ORDER BY 1 DESC"
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        If right(Trim("" & rst.Fields(0)), 4) = "9999" Then
        
            GeneraCorrelativoAnoMesNuevo = GeneraCorrelativoAnoMesNuevo(strTabla, strCod, left(PPrefijo, 4) & Format(Val(right(PPrefijo, 2)) + 1, "00"), indEmpresa)
            
        Else
        
            GeneraCorrelativoAnoMesNuevo = strCond & Format((Val(right("" & rst.Fields(0), 4)) + 1), "0000")
            
        End If
        
    Else
        GeneraCorrelativoAnoMesNuevo = strCond & "0001"
    End If
    rst.Close: Set rst = Nothing
    
End Function

Public Sub GeneraTxt(StrMsgError As String, PRs As ADODB.Recordset, PGlsArchivo As String)
On Error GoTo Err
Dim i                           As Long
Dim cruta                       As String
Dim CLinea                      As String

    cruta = App.Path & "\Temporales\" & PGlsArchivo
    Open cruta For Output As #1
    
    Do While Not PRs.EOF
        
        CLinea = ""
        
        For i = 0 To PRs.Fields.Count - 1
            
            CLinea = CLinea & PRs.Fields(i) & "|"
            
        Next i
        
        Print #1, CLinea
        
        PRs.MoveNext
        
    Loop
    
    Close #1
    
    ShellEx cruta, essSW_MAXIMIZE, , , "open" ', Me.hwnd
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub AbrirRecordset(StrMsgError As String, PCn As ADODB.Connection, PRsC As ADODB.Recordset, PSqlC As String)
On Error GoTo NoConecta
    PRsC.Open "Select IdPruebaConeccion From PruebaConeccion", PCn, adOpenStatic, adLockReadOnly
    PRsC.Close: Set PRsC = Nothing
    
On Error GoTo Err
    
    Set PRsC = New ADODB.Recordset
    PRsC.Open PSqlC, PCn, adOpenStatic, adLockReadOnly
    
    Exit Sub
Err:
    If PRsC.State = 1 Then PRsC.Close: Set PRsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
    
NoConecta:
    abrirConexion StrMsgError, False
    If StrMsgError <> "" Then GoTo Err
    Resume
End Sub

Public Sub DocumentoElectronico(PForm As Form, pDocumento As String, pSerie As String, PNumero As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim pRecCab As ADODB.Recordset
Dim pRecDet As ADODB.Recordset
Dim pRecEmp As ADODB.Recordset
Dim CCarpeta                        As String
Dim CGlsCab                         As String
Dim CSqlC                           As String
Dim rsref                           As New ADODB.Recordset
Dim CIdDocumentoRef                 As String
Dim CIdSerieRef                     As String
Dim CIdNumeroRef                    As String
Dim CGlsDet                         As String
Dim RsC                             As New ADODB.Recordset
Dim RetVal
Dim strRuta As String

    strRUC = traerCampo("Empresas", "Ruc", "idEmpresa", glsEmpresa, False)
    ' pregunto si el XML anterior fue generado
    If Val(PNumero) > 1 Then
        CCarpeta = leeParametro("CARPETA_XML_VE")
        
        'consulta SI EL xml YA EXISTE en la carpeta
        strRuta = CCarpeta & "\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & PNumero & ".xml"
        If Existe(strRuta) Then
            StrMsgError = "El XML del documento electrónico ya existe"
            GoTo Err
        End If
        'consulta SI EL xml YA EXISTE en la carpeta LOG
        strRuta = CCarpeta & "_WORK\outputs\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & PNumero & ".xml"
        If Existe(strRuta) Then
            StrMsgError = "El XML del documento electrónico ya existe"
            GoTo Err
        End If
        'consulta si existe el documento anterior
        If Not (Existe(CCarpeta & "\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & Format(Val(PNumero) - 1, "00000000") & ".xml")) And Not (Existe(CCarpeta & "_WORK\outputs\R-" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & Format(Val(PNumero) - 1, "00000000") & ".xml")) And Not (Existe(CCarpeta & "_WORK\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & Format(Val(PNumero) - 1, "00000000") & ".xml")) Then
            StrMsgError = "Debe generar el XML del documento electrónico anterior"
            GoTo Err
        End If
    End If
    
    csql = "Select p.idTipoDocIdentidad,p.idDistrito,u.glsUbigeo,p.idPais,dv.idDocumento,dv.idSerie,dv.idDocVentas,dv.idMoneda,dv.IndVtaGratuita," & _
           "dv.TotalBaseImponible,dv.TotalExonerado,dv.TotalValorVenta,dv.TotalDescuento,dv.totalLetras,dv.fecEmision,dv.IdMotivoNCD,dv.ObsDocVentas," & _
           "dv.RucCliente,dv.GlsCliente,dv.TotalIGVVenta,dv.TotalPrecioVenta,dv.TotalDescuentoGlobalGravado,dv.TotalDescuentoGlobalExonerado " & _
           "from docventas dv inner join personas p on dv.idPerCliente=p.idpersona inner join Ubigeo u On p.IdDistrito = u.IdDistrito And p.IdPais = u.IdPais where dv.idEmpresa='" & glsEmpresa & "' and dv.idDocumento='" & pDocumento & "' and dv.idSerie='" & pSerie & "' and dv.idDocVentas='" & PNumero & "'"
    Set pRecCab = New ADODB.Recordset
    pRecCab.Open csql, Cn
    
    csql = "Select e.*, p.idDistrito, p.direccion, p.idPais, u.glsUbigeo from empresas e inner join personas p on e.idPersona=p.idPersona inner join Ubigeo u On p.IdDistrito = u.IdDistrito And p.IdPais = u.IdPais Where e.idEmpresa='" & glsEmpresa & "'; "
    Set pRecEmp = New ADODB.Recordset
    pRecEmp.Open csql, Cn
    
    If Not pRecCab.EOF And Not pRecEmp.EOF Then
        
        If pDocumento = "07" Then
            CGlsCab = "CreditNote"
            CGlsDet = "Credited"
        ElseIf pDocumento = "08" Then
            CGlsCab = "DebitNote"
            CGlsDet = "Debited"
        Else
            CGlsCab = "Invoice"
            CGlsDet = "Invoiced"
        End If
        
        STREMP = QuitarCaracteresEspeciales("" & pRecEmp("GlsEmpresa"))
        strDpt = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Mid("" & pRecEmp("idDistrito"), 1, 2), False, " idProv = '00' And idPais = '" & pRecEmp("idPais") & "' ")
        strPrv = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Mid("" & pRecEmp("idDistrito"), 3, 2), False, " idDpto = '" & Mid("" & pRecEmp("idDistrito"), 1, 2) & "' and idProv <> '00' and idDist = '00' And idPais = '" & pRecEmp("idPais") & "' ")
        
        'CERTIFICADO DIGITAL
        'strRuta = App.Path & "\Temporales\" & strRUC & ".cer"
        'IntFile = FreeFile
        'strFirma = ""
        'Open strRuta For Input As #IntFile
        'Do While Not EOF(IntFile)
        '    Line Input #IntFile, strLinea
        '    strFirma = strFirma & "" & vbCrLf
        'Loop
        'Close #IntFile
        
        CCarpeta = leeParametro("CARPETA_XML_VE")
        
        strRuta = CCarpeta & "\" & strRUC & "-" & pRecCab("idDocumento") & "-" & pRecCab("idSerie") & "-" & pRecCab("idDocVentas") & ".xml"
        IntFile = FreeFile
        Open strRuta For Output As #IntFile
        'Cabecera
        strLinea = ""
        strLinea = strLinea & "<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?> " & vbCrLf
        strLinea = strLinea & "<" & CGlsCab & " " & vbCrLf
        strLinea = strLinea & "    xmlns=""urn:oasis:names:specification:ubl:schema:xsd:" & CGlsCab & "-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:cac=""urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:cbc=""urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:ccts=""urn:un:unece:uncefact:documentation:2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:ds=""http://www.w3.org/2000/09/xmldsig#"" " & vbCrLf
        strLinea = strLinea & "    xmlns:ext=""urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:qdt=""urn:oasis:names:specification:ubl:schema:xsd:QualifiedDatatypes-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:sac=""urn:sunat:names:specification:ubl:peru:schema:xsd:SunatAggregateComponents-1"" " & vbCrLf
        strLinea = strLinea & "    xmlns:udt=""urn:un:unece:uncefact:data:specification:UnqualifiedDataTypesSchemaModule:2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""> " & vbCrLf
        strLinea = strLinea & "    <ext:UBLExtensions> " & vbCrLf
        strLinea = strLinea & "        <ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "            <ext:ExtensionContent> " & vbCrLf
        strLinea = strLinea & "                <sac:AdditionalInformation> " & vbCrLf
        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "                        <cbc:ID>1001</cbc:ID> " & vbCrLf
        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalBaseImponible")), "0.00") & "</cbc:PayableAmount> " & vbCrLf
        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
        
        If pDocumento <> "07" And pDocumento <> "08" Then
        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "                        <cbc:ID>1002</cbc:ID> " & vbCrLf
        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalExonerado")), "0.00") & "</cbc:PayableAmount> " & vbCrLf
        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "                        <cbc:ID>1003</cbc:ID> " & vbCrLf
        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(0, "0.00") & "</cbc:PayableAmount> " & vbCrLf
        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "                        <cbc:ID>1004</cbc:ID> " & vbCrLf
        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, pRecCab("TotalValorVenta"), 0), "0.00") & "</cbc:PayableAmount> " & vbCrLf
        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "                        <cbc:ID>2005</cbc:ID> " & vbCrLf
        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & pRecCab("TotalDescuento") & "</cbc:PayableAmount> " & vbCrLf
        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "                    <sac:AdditionalProperty> " & vbCrLf
        strLinea = strLinea & "                        <cbc:ID>" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "1002", "1000") & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "                        <cbc:Value>" & pRecCab("totalLetras") & "</cbc:Value> " & vbCrLf
        strLinea = strLinea & "                    </sac:AdditionalProperty> " & vbCrLf
        End If
        
        strLinea = strLinea & "                </sac:AdditionalInformation> " & vbCrLf
        strLinea = strLinea & "            </ext:ExtensionContent> " & vbCrLf
        strLinea = strLinea & "        </ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "        <ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "            <ext:ExtensionContent> " & vbCrLf
'        strLinea = strLinea & "                <ds:Signature Id=""SignatureSP""> " & vbCrLf
'        strLinea = strLinea & "                    <ds:SignedInfo> " & vbCrLf
'        strLinea = strLinea & "                        <ds:CanonicalizationMethod Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/> " & vbCrLf
'        strLinea = strLinea & "                        <ds:SignatureMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#rsa-sha1""/> " & vbCrLf
'        strLinea = strLinea & "                        <ds:Reference URI=""""> " & vbCrLf
'        strLinea = strLinea & "                            <ds:Transforms> " & vbCrLf
'        strLinea = strLinea & "                                <ds:Transform Algorithm=""http://www.w3.org/2000/09/xmldsig#envelopedsignature""/> " & vbCrLf
'        strLinea = strLinea & "                            </ds:Transforms> " & vbCrLf
'        strLinea = strLinea & "                            <ds:DigestMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#sha1""/> " & vbCrLf
'        strLinea = strLinea & "                            <ds:DigestValue>+pruib33lOapq6GSw58GgQLR8VGIGqANloj4EqB1cb4=</ds:DigestValue> " & vbCrLf
'        strLinea = strLinea & "                        </ds:Reference> " & vbCrLf
'        strLinea = strLinea & "                    </ds:SignedInfo> " & vbCrLf
'        strLinea = strLinea & "                    <ds:SignatureValue>Oatv5xMfFInuGqiX9SoLDTy2yuLf0tTlMFkWtkdw1z/Ss6kiDz+vIgZhgKfIaxp+JbVy57 " & vbCrLf
'        strLinea = strLinea & "GT52f1 " & vbCrLf
'        strLinea = strLinea & "8D6+WMYZ0xOxTK2mojNkJNewwTTXzqOqrrAlObs9YoS5JAQAMi/TwkR4brNniU9tVwyybirHxw0H " & vbCrLf
'        strLinea = strLinea & "WVzN2bB43yQd9hOlXzRUYpC8/sXw78h7ME3E/zeu882aOFySOnHWB63imBQGcYBV+LIGR/JW8ER+ " & vbCrLf
'        strLinea = strLinea & "0VLMLatdwPVRbrWmz1/NIy5CWp1xWMaM6fC/9SXV0O1Lqopk0UeX2I2yuf05QhmVfjgUu6GnS3m6 " & vbCrLf
'        strLinea = strLinea & "o6zM9J36iDvMVZyj7vbJTwI8SfWjTSNqxXlqPQ==</ds:SignatureValue> " & vbCrLf
'        strLinea = strLinea & "                    <ds:KeyInfo> " & vbCrLf
'        strLinea = strLinea & "                        <ds:X509Data> " & vbCrLf
'        strLinea = strLinea & "                            <!-- <ds:X509SubjectName>X509_SUBJECT_TEST</ds:X509SubjectName> --> " & vbCrLf
'        strLinea = strLinea & "                            <ds:X509Certificate>" & strFirma & "</ds:X509Certificate> " & vbCrLf
'        strLinea = strLinea & "                        </ds:X509Data> " & vbCrLf
'        strLinea = strLinea & "                    </ds:KeyInfo> " & vbCrLf
'        strLinea = strLinea & "                </ds:Signature> " & vbCrLf
        strLinea = strLinea & "            </ext:ExtensionContent> " & vbCrLf
        strLinea = strLinea & "        </ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "    </ext:UBLExtensions> " & vbCrLf
        strLinea = strLinea & "    <cbc:UBLVersionID>2.0</cbc:UBLVersionID> " & vbCrLf
        strLinea = strLinea & "    <cbc:CustomizationID>1.0</cbc:CustomizationID> " & vbCrLf
        strLinea = strLinea & "    <cbc:ID>" & pRecCab("idSerie") & "-" & pRecCab("idDocVentas") & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "    <cbc:IssueDate>" & Format(pRecCab("fecEmision"), "yyyy-mm-dd") & "</cbc:IssueDate> " & vbCrLf
        
        If pDocumento <> "07" And pDocumento <> "08" Then
        strLinea = strLinea & "    <cbc:InvoiceTypeCode>" & pRecCab("idDocumento") & "</cbc:InvoiceTypeCode> " & vbCrLf
        End If
        
        strLinea = strLinea & "    <cbc:DocumentCurrencyCode>" & pRecCab("idMoneda") & "</cbc:DocumentCurrencyCode> " & vbCrLf
        
        If pDocumento = "07" Or pDocumento = "08" Then
        
        CSqlC = "Select A.TipoDocReferencia,A.SerieDocReferencia,A.NumDocReferencia " & _
                "From DocReferencia A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.TipoDocOrigen = '" & pDocumento & "' And A.SerieDocOrigen = '" & pSerie & "' " & _
                "And A.NumDocOrigen = '" & PNumero & "'"
        AbrirRecordset StrMsgError, Cn, rsref, CSqlC: If StrMsgError <> "" Then GoTo Err
        If Not rsref.EOF Then
            CIdDocumentoRef = Trim("" & rsref.Fields("TipoDocReferencia"))
            CIdSerieRef = Trim("" & rsref.Fields("SerieDocReferencia"))
            CIdNumeroRef = Trim("" & rsref.Fields("NumDocReferencia"))
        End If
        rsref.Close: Set rsref = Nothing
        
        strLinea = strLinea & "    <cac:DiscrepancyResponse> " & vbCrLf
        strLinea = strLinea & "        <cbc:ReferenceID>" & CIdSerieRef & "-" & CIdNumeroRef & "</cbc:ReferenceID> " & vbCrLf
        
        CSqlC = "Select A.GlsMotivoNCD,A.IdCodigoVE " & _
                "From MotivosNCD A " & _
                "Where A.IdMotivoNCD = '" & Trim("" & pRecCab.Fields("IdMotivoNCD")) & "'"
        AbrirRecordset StrMsgError, Cn, RsC, CSqlC: If StrMsgError <> "" Then GoTo Err
        If Not RsC.EOF Then
        strLinea = strLinea & "        <cbc:ResponseCode>" & Trim("" & RsC.Fields("IdCodigoVE")) & "</cbc:ResponseCode> " & vbCrLf
        End If
        RsC.Close: Set RsC = Nothing
        
        strLinea = strLinea & "        <cbc:Description>" & QuitarCaracteresEspeciales("" & pRecCab.Fields("ObsDocVentas")) & "</cbc:Description> " & vbCrLf
        strLinea = strLinea & "    </cac:DiscrepancyResponse> " & vbCrLf
        strLinea = strLinea & "    <cac:BillingReference> " & vbCrLf
        strLinea = strLinea & "        <cac:InvoiceDocumentReference> " & vbCrLf
        strLinea = strLinea & "            <cbc:ID>" & CIdSerieRef & "-" & CIdNumeroRef & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "            <cbc:DocumentTypeCode>" & CIdDocumentoRef & "</cbc:DocumentTypeCode> " & vbCrLf
        strLinea = strLinea & "        </cac:InvoiceDocumentReference> " & vbCrLf
        strLinea = strLinea & "    </cac:BillingReference> " & vbCrLf
        End If
        If Trim(PForm.txt_OrdenCompra.Text) <> "" And pDocumento = "01" Then
        'Agregar la orden de compra
        strLinea = strLinea & "    <cac:OrderReference> " & vbCrLf
        strLinea = strLinea & "        <cbc:ID><![CDATA[" & Trim(PForm.txt_OrdenCompra.Text) & "]]></cbc:ID> " & vbCrLf
        strLinea = strLinea & "    </cac:OrderReference> " & vbCrLf
        End If
        strLinea = strLinea & "    <cac:Signature> " & vbCrLf
        strLinea = strLinea & "        <cbc:ID>IDSignSP</cbc:ID> " & vbCrLf
        strLinea = strLinea & "        <cac:SignatoryParty> " & vbCrLf
        strLinea = strLinea & "            <cac:PartyIdentification> " & vbCrLf
        strLinea = strLinea & "                <cbc:ID>" & strRUC & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "            </cac:PartyIdentification> " & vbCrLf
        strLinea = strLinea & "            <cac:PartyName> " & vbCrLf
        strLinea = strLinea & "                <cbc:Name><![CDATA[" & STREMP & "]]></cbc:Name> " & vbCrLf
        strLinea = strLinea & "            </cac:PartyName> " & vbCrLf
        strLinea = strLinea & "        </cac:SignatoryParty> " & vbCrLf
        strLinea = strLinea & "        <cac:DigitalSignatureAttachment> " & vbCrLf
        strLinea = strLinea & "            <cac:ExternalReference> " & vbCrLf
        strLinea = strLinea & "                <cbc:URI>#SignatureSP</cbc:URI> " & vbCrLf
        strLinea = strLinea & "            </cac:ExternalReference> " & vbCrLf
        strLinea = strLinea & "        </cac:DigitalSignatureAttachment> " & vbCrLf
        strLinea = strLinea & "    </cac:Signature> " & vbCrLf
        strLinea = strLinea & "    <cac:AccountingSupplierParty> " & vbCrLf
        strLinea = strLinea & "        <cbc:CustomerAssignedAccountID>" & strRUC & "</cbc:CustomerAssignedAccountID> " & vbCrLf
        strLinea = strLinea & "        <cbc:AdditionalAccountID>6</cbc:AdditionalAccountID> " & vbCrLf
        strLinea = strLinea & "        <cac:Party> " & vbCrLf
        strLinea = strLinea & "            <cac:PostalAddress> " & vbCrLf
        strLinea = strLinea & "                <cbc:ID>" & pRecEmp("idDistrito") & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "                <cbc:StreetName>" & STREMP & "</cbc:StreetName> " & vbCrLf
        strLinea = strLinea & "                <cbc:CitySubdivisionName></cbc:CitySubdivisionName> " & vbCrLf
        strLinea = strLinea & "                <cbc:CityName>" & strDpt & "</cbc:CityName> " & vbCrLf
        strLinea = strLinea & "                <cbc:CountrySubentity>" & strPrv & "</cbc:CountrySubentity> " & vbCrLf
        strLinea = strLinea & "                <cbc:District>" & pRecEmp("glsUbigeo") & "</cbc:District> " & vbCrLf
        strLinea = strLinea & "                <cac:Country> " & vbCrLf
        strLinea = strLinea & "                    <cbc:IdentificationCode>PE</cbc:IdentificationCode> " & vbCrLf
        strLinea = strLinea & "                </cac:Country> " & vbCrLf
        strLinea = strLinea & "            </cac:PostalAddress> " & vbCrLf
        strLinea = strLinea & "            <cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "                <cbc:RegistrationName><![CDATA[" & STREMP & "]]></cbc:RegistrationName> " & vbCrLf
        strLinea = strLinea & "            </cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "        </cac:Party> " & vbCrLf
        strLinea = strLinea & "    </cac:AccountingSupplierParty> " & vbCrLf
        strLinea = strLinea & "    <cac:AccountingCustomerParty> " & vbCrLf
        strLinea = strLinea & "        <cbc:CustomerAssignedAccountID>" & pRecCab("RucCliente") & "</cbc:CustomerAssignedAccountID> " & vbCrLf
        strLinea = strLinea & "        <cbc:AdditionalAccountID>" & pRecCab("idTipoDocIdentidad") & "</cbc:AdditionalAccountID> " & vbCrLf
        strLinea = strLinea & "        <cac:Party> " & vbCrLf
        strLinea = strLinea & "            <cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "                <cbc:RegistrationName>" & QuitarCaracteresEspeciales("" & pRecCab("GlsCliente")) & "</cbc:RegistrationName> " & vbCrLf
        strLinea = strLinea & "            </cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "        </cac:Party> " & vbCrLf
        strLinea = strLinea & "    </cac:AccountingCustomerParty> " & vbCrLf
        strLinea = strLinea & "    <cac:TaxTotal> " & vbCrLf
        strLinea = strLinea & "        <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecCab("TotalIGVVenta"), "0.00") & "</cbc:TaxAmount> " & vbCrLf
        strLinea = strLinea & "        <cac:TaxSubtotal> " & vbCrLf
        strLinea = strLinea & "            <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecCab("TotalIGVVenta"), "0.00") & "</cbc:TaxAmount> " & vbCrLf
        strLinea = strLinea & "            <cac:TaxCategory> " & vbCrLf
        strLinea = strLinea & "                <cac:TaxScheme> " & vbCrLf
        strLinea = strLinea & "                    <cbc:ID>1000</cbc:ID> " & vbCrLf
        strLinea = strLinea & "                    <cbc:Name>IGV</cbc:Name> " & vbCrLf
        strLinea = strLinea & "                    <cbc:TaxTypeCode>VAT</cbc:TaxTypeCode> " & vbCrLf
        strLinea = strLinea & "                </cac:TaxScheme> " & vbCrLf
        strLinea = strLinea & "            </cac:TaxCategory> " & vbCrLf
        strLinea = strLinea & "        </cac:TaxSubtotal> " & vbCrLf
        strLinea = strLinea & "    </cac:TaxTotal> " & vbCrLf
        
        If pDocumento = "08" Then
        strLinea = strLinea & "    <cac:RequestedMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecCab("TotalPrecioVenta"), "0.00") & "</cbc:PayableAmount> " & vbCrLf
        strLinea = strLinea & "    </cac:RequestedMonetaryTotal> " & vbCrLf
        Else
        strLinea = strLinea & "    <cac:LegalMonetaryTotal> " & vbCrLf
        If Val("" & pRecCab.Fields("TotalDescuentoGlobalGravado")) > 0 Or Val("" & pRecCab.Fields("TotalDescuentoGlobalExonerado")) > 0 Then
        strLinea = strLinea & "        <cbc:AllowanceTotalAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, Val("" & pRecCab.Fields("TotalDescuentoGlobalGravado")) + Val("" & pRecCab.Fields("TotalDescuentoGlobalExonerado"))), "0.00") & "</cbc:AllowanceTotalAmount> " & vbCrLf
        End If
        strLinea = strLinea & "        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalPrecioVenta")), "0.00") & "</cbc:PayableAmount> " & vbCrLf
        strLinea = strLinea & "    </cac:LegalMonetaryTotal> " & vbCrLf
        End If
        
        csql = "Select If(Afecto = '1','" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "11", "10") & "','" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "16", "30") & "') Afecto,Cantidad,TotalVVNeto,PVUnit,TotalIGVNeto,glsProducto,idProducto,VVUnit from docventasdet where idEmpresa='" & glsEmpresa & "' and idDocumento='" & pDocumento & "' and idSerie='" & pSerie & "' and idDocVentas='" & PNumero & "' Order by item "
        Set pRecDet = New ADODB.Recordset
        pRecDet.Open csql, Cn
        Print #IntFile, strLinea
        item = 0
        'Detalle
        Do While Not pRecDet.EOF
            item = item + 1
            strLinea = ""
            strLinea = strLinea & "    <cac:" & CGlsCab & "Line> " & vbCrLf
            strLinea = strLinea & "        <cbc:ID>" & item & "</cbc:ID> " & vbCrLf
            strLinea = strLinea & "        <cbc:" & CGlsDet & "Quantity unitCode=""NIU"">" & pRecDet("Cantidad") & "</cbc:" & CGlsDet & "Quantity> " & vbCrLf
            strLinea = strLinea & "        <cbc:LineExtensionAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecDet("TotalVVNeto")), "0.00") & "</cbc:LineExtensionAmount> " & vbCrLf
            strLinea = strLinea & "        <cac:PricingReference> " & vbCrLf
            strLinea = strLinea & "            <cac:AlternativeConditionPrice> " & vbCrLf
            strLinea = strLinea & "                <cbc:PriceAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecDet("PVUnit")), "0.00") & "</cbc:PriceAmount> " & vbCrLf
            strLinea = strLinea & "                <cbc:PriceTypeCode>01</cbc:PriceTypeCode> " & vbCrLf
            strLinea = strLinea & "            </cac:AlternativeConditionPrice> " & vbCrLf
            If Val("" & pRecCab.Fields("IndVtaGratuita")) = 1 Then
            strLinea = strLinea & "            <cac:AlternativeConditionPrice> " & vbCrLf
            strLinea = strLinea & "                <cbc:PriceAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("PVUnit"), "0.00") & "</cbc:PriceAmount> " & vbCrLf
            strLinea = strLinea & "                <cbc:PriceTypeCode>02</cbc:PriceTypeCode> " & vbCrLf
            strLinea = strLinea & "            </cac:AlternativeConditionPrice> " & vbCrLf
            End If
            strLinea = strLinea & "        </cac:PricingReference> " & vbCrLf
            strLinea = strLinea & "        <cac:TaxTotal> " & vbCrLf
            strLinea = strLinea & "            <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("TotalIGVNeto"), "0.00") & "</cbc:TaxAmount> " & vbCrLf
            strLinea = strLinea & "            <cac:TaxSubtotal> " & vbCrLf
            strLinea = strLinea & "                <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("TotalIGVNeto"), "0.00") & "</cbc:TaxAmount> " & vbCrLf
            strLinea = strLinea & "                <cac:TaxCategory> " & vbCrLf
            strLinea = strLinea & "                    <cbc:TaxExemptionReasonCode>" & Trim("" & pRecDet.Fields("Afecto")) & "</cbc:TaxExemptionReasonCode> " & vbCrLf
            strLinea = strLinea & "                    <cac:TaxScheme> " & vbCrLf
            strLinea = strLinea & "                        <cbc:ID>1000</cbc:ID> " & vbCrLf
            strLinea = strLinea & "                        <cbc:Name>IGV</cbc:Name> " & vbCrLf
            strLinea = strLinea & "                        <cbc:TaxTypeCode>VAT</cbc:TaxTypeCode> " & vbCrLf
            strLinea = strLinea & "                    </cac:TaxScheme> " & vbCrLf
            strLinea = strLinea & "                </cac:TaxCategory> " & vbCrLf
            strLinea = strLinea & "            </cac:TaxSubtotal> " & vbCrLf
            strLinea = strLinea & "        </cac:TaxTotal> " & vbCrLf
            strLinea = strLinea & "        <cac:Item> " & vbCrLf
            strLinea = strLinea & "            <cbc:Description>" & QuitarCaracteresEspeciales("" & pRecDet("glsProducto")) & "</cbc:Description> " & vbCrLf
            strLinea = strLinea & "            <cac:SellersItemIdentification> " & vbCrLf
            strLinea = strLinea & "                <cbc:ID>" & pRecDet("idProducto") & "</cbc:ID> " & vbCrLf
            strLinea = strLinea & "            </cac:SellersItemIdentification> " & vbCrLf
            strLinea = strLinea & "        </cac:Item> " & vbCrLf
            strLinea = strLinea & "        <cac:Price> " & vbCrLf
            strLinea = strLinea & "            <cbc:PriceAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecDet("VVUnit")), "0.00") & "</cbc:PriceAmount> " & vbCrLf
            strLinea = strLinea & "        </cac:Price> " & vbCrLf
            strLinea = strLinea & "    </cac:" & CGlsCab & "Line> " & vbCrLf
            Print #IntFile, strLinea
            pRecDet.MoveNext
        Loop
        'strLinea = strLinea & "</Invoice>" & vbCrLf
        strLinea = "</" & CGlsCab & ">" & vbCrLf
        Print #IntFile, strLinea
        Close #IntFile
    End If
    MsgBox "Se creo el archivo XML, en unos momentos se enviara al sistema SOL", vbInformation
    
    Cn.Execute "Update DocVentas Set IndEnviadoSunat = 1 Where IdEmpresa = '" & glsEmpresa & "' And IdDocumento = '" & pDocumento & "' And IdSerie = '" & pSerie & "' And IdDocVentas = '" & PNumero & "'"
    
    RetVal = ShellExecute(PForm.hwnd, "Open", App.Path & "\Release\FEapp2.exe", "", "", 0)
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Close #IntFile
    Exit Sub
    Resume
End Sub

Public Sub DocumentoElectronico21(PForm As Form, pDocumento As String, pSerie As String, PNumero As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim pRecCab As ADODB.Recordset
Dim pRecDet As ADODB.Recordset
Dim pRecEmp As ADODB.Recordset
Dim CCarpeta                        As String
Dim CGlsCab                         As String
Dim CSqlC                           As String
Dim rsref                           As New ADODB.Recordset
Dim CIdDocumentoRef                 As String
Dim CIdSerieRef                     As String
Dim CIdNumeroRef                    As String
Dim CGlsDet                         As String
Dim RsC                             As New ADODB.Recordset
Dim RetVal
Dim strRuta As String
Dim bolExportacion As Boolean

    strRUC = traerCampo("Empresas", "Ruc", "idEmpresa", glsEmpresa, False)
    bolExportacion = False
    
    ' pregunto si el XML anterior fue generado
    If Val(PNumero) > 1 Then
        CCarpeta = leeParametro("CARPETA_XML_VE")
        
        'consulta SI EL xml YA EXISTE en la carpeta
        strRuta = CCarpeta & "\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & PNumero & ".xml"
        If Existe(strRuta) Then
            StrMsgError = "El XML del documento electrónico ya existe"
            GoTo Err
        End If
        'consulta SI EL xml YA EXISTE en la carpeta LOG
        strRuta = CCarpeta & "_WORK\outputs\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & PNumero & ".xml"
        If Existe(strRuta) Then
            StrMsgError = "El XML del documento electrónico ya existe"
            GoTo Err
        End If
        'consulta si existe el documento anterior
        If Not (Existe(CCarpeta & "\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & Format(Val(PNumero) - 1, "00000000") & ".xml")) And Not (Existe(CCarpeta & "_WORK\outputs\R-" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & Format(Val(PNumero) - 1, "00000000") & ".xml")) And Not (Existe(CCarpeta & "_WORK\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & Format(Val(PNumero) - 1, "00000000") & ".xml")) Then
            StrMsgError = "Debe generar el XML del documento electrónico anterior"
            GoTo Err
        End If
    End If
    
    csql = "Select p.idTipoDocIdentidad,p.idDistrito,u.glsUbigeo,p.idPais,dv.idDocumento,dv.idSerie,dv.idDocVentas,dv.idMoneda,dv.IndVtaGratuita," & _
           "dv.TotalBaseImponible,dv.TotalExonerado,dv.TotalValorVenta,dv.TotalDescuento,dv.totalLetras,dv.fecEmision,dv.IdMotivoNCD,dv.ObsDocVentas," & _
           "dv.RucCliente,dv.GlsCliente,dv.TotalIGVVenta,dv.TotalPrecioVenta,dv.TotalDescuentoGlobalGravado,dv.TotalDescuentoGlobalExonerado " & _
           "from docventas dv inner join personas p on dv.idPerCliente=p.idpersona inner join Ubigeo u On p.IdDistrito = u.IdDistrito And p.IdPais = u.IdPais where dv.idEmpresa='" & glsEmpresa & "' and dv.idDocumento='" & pDocumento & "' and dv.idSerie='" & pSerie & "' and dv.idDocVentas='" & PNumero & "'"
    Set pRecCab = New ADODB.Recordset
    pRecCab.Open csql, Cn
    
    csql = "Select e.*, p.idDistrito, p.direccion, p.idPais, u.glsUbigeo from empresas e inner join personas p on e.idPersona=p.idPersona inner join Ubigeo u On p.IdDistrito = u.IdDistrito And p.IdPais = u.IdPais Where e.idEmpresa='" & glsEmpresa & "'; "
    Set pRecEmp = New ADODB.Recordset
    pRecEmp.Open csql, Cn
    
    'si es una exportacion
    If pRecCab("idTipoDocIdentidad") <> 6 And pRecCab("TotalIGVVenta") = 0 And pRecCab("idMoneda") <> "PEN" Then
        bolExportacion = True
    End If
    
    If Not pRecCab.EOF And Not pRecEmp.EOF Then
        
        If pDocumento = "07" Then
            CGlsCab = "CreditNote"
            CGlsDet = "Credited"
        ElseIf pDocumento = "08" Then
            CGlsCab = "DebitNote"
            CGlsDet = "Debited"
        Else
            CGlsCab = "Invoice"
            CGlsDet = "Invoiced"
        End If
        
        STREMP = QuitarCaracteresEspeciales("" & pRecEmp("GlsEmpresa"))
        strDpt = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Mid("" & pRecEmp("idDistrito"), 1, 2), False, " idProv = '00' And idPais = '" & pRecEmp("idPais") & "' ")
        strPrv = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Mid("" & pRecEmp("idDistrito"), 3, 2), False, " idDpto = '" & Mid("" & pRecEmp("idDistrito"), 1, 2) & "' and idProv <> '00' and idDist = '00' And idPais = '" & pRecEmp("idPais") & "' ")
        
        
        CCarpeta = leeParametro("CARPETA_XML_VE")
        
        strRuta = CCarpeta & "\" & strRUC & "-" & pRecCab("idDocumento") & "-" & pRecCab("idSerie") & "-" & pRecCab("idDocVentas") & ".xml"
        IntFile = FreeFile
        Open strRuta For Output As #IntFile
        'Cabecera
        strLinea = ""
        If pDocumento = "07" Or pDocumento = "08" Then
        strLinea = strLinea & "<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?> " & vbCrLf
        Else
        strLinea = strLinea & "<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?> " & vbCrLf
        End If
        
        strLinea = strLinea & "<" & CGlsCab & " " & vbCrLf
        strLinea = strLinea & "    xmlns=""urn:oasis:names:specification:ubl:schema:xsd:" & CGlsCab & "-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:cac=""urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:cbc=""urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"" " & vbCrLf
'        strLinea = strLinea & "    xmlns:ccts=""urn:un:unece:uncefact:documentation:2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:ds=""http://www.w3.org/2000/09/xmldsig#"" " & vbCrLf
        strLinea = strLinea & "    xmlns:ext=""urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2"">" & vbCrLf
'        strLinea = strLinea & "    xmlns:qdt=""urn:oasis:names:specification:ubl:schema:xsd:QualifiedDatatypes-2"" " & vbCrLf
'        strLinea = strLinea & "    xmlns:sac=""urn:sunat:names:specification:ubl:peru:schema:xsd:SunatAggregateComponents-1"" " & vbCrLf
'        strLinea = strLinea & "    xmlns:udt=""urn:un:unece:uncefact:data:specification:UnqualifiedDataTypesSchemaModule:2"" " & vbCrLf
'        strLinea = strLinea & "    xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""> " & vbCrLf
        strLinea = strLinea & "    <ext:UBLExtensions> " & vbCrLf
'        strLinea = strLinea & "        <ext:UBLExtension> " & vbCrLf
'        strLinea = strLinea & "            <ext:ExtensionContent> " & vbCrLf
'        strLinea = strLinea & "                <sac:AdditionalInformation> " & vbCrLf
'        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:ID>1001</cbc:ID> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalBaseImponible")), "0.00") & "</cbc:PayableAmount> " & vbCrLf
'        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
'
'        If pDocumento <> "07" And pDocumento <> "08" Then
'        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:ID>1002</cbc:ID> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalExonerado")), "0.00") & "</cbc:PayableAmount> " & vbCrLf
'        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
'        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:ID>1003</cbc:ID> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(0, "0.00") & "</cbc:PayableAmount> " & vbCrLf
'        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
'        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:ID>1004</cbc:ID> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, pRecCab("TotalValorVenta"), 0), "0.00") & "</cbc:PayableAmount> " & vbCrLf
'        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
'        strLinea = strLinea & "                    <sac:AdditionalMonetaryTotal> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:ID>2005</cbc:ID> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & pRecCab("TotalDescuento") & "</cbc:PayableAmount> " & vbCrLf
'        strLinea = strLinea & "                    </sac:AdditionalMonetaryTotal> " & vbCrLf
'        strLinea = strLinea & "                    <sac:AdditionalProperty> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:ID>" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "1002", "1000") & "</cbc:ID> " & vbCrLf
'        strLinea = strLinea & "                        <cbc:Value>" & pRecCab("totalLetras") & "</cbc:Value> " & vbCrLf
'        strLinea = strLinea & "                    </sac:AdditionalProperty> " & vbCrLf
'        End If
'
'        strLinea = strLinea & "                </sac:AdditionalInformation> " & vbCrLf
'        strLinea = strLinea & "            </ext:ExtensionContent> " & vbCrLf
'        strLinea = strLinea & "        </ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "        <ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "            <ext:ExtensionContent> " & vbCrLf
        strLinea = strLinea & "            </ext:ExtensionContent> " & vbCrLf
        strLinea = strLinea & "        </ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "    </ext:UBLExtensions> " & vbCrLf
        strLinea = strLinea & "    <cbc:UBLVersionID>2.1</cbc:UBLVersionID> " & vbCrLf
        strLinea = strLinea & "    <cbc:CustomizationID>2.0</cbc:CustomizationID> " & vbCrLf
        strLinea = strLinea & "    <cbc:ID>" & pRecCab("idSerie") & "-" & pRecCab("idDocVentas") & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "    <cbc:IssueDate>" & Format(pRecCab("fecEmision"), "yyyy-mm-dd") & "</cbc:IssueDate> " & vbCrLf
        strLinea = strLinea & "    <cbc:IssueTime>" & Format(Now, "hh:MM:ss") & "</cbc:IssueTime> " & vbCrLf
        
        If pDocumento <> "07" And pDocumento <> "08" Then
        strLinea = strLinea & "    <cbc:DueDate>" & Format(pRecCab("fecEmision"), "yyyy-mm-dd") & "</cbc:DueDate> " & vbCrLf
        'strLinea = strLinea & "    <cbc:InvoiceTypeCode>" & pRecCab("idDocumento") & "</cbc:InvoiceTypeCode> " & vbCrLf
        strLinea = strLinea & "    <cbc:InvoiceTypeCode listID=""" & IIf(Not bolExportacion, "0101", "0200") & """ listAgencyName=""PE:SUNAT"" listName=""SUNAT:Identificador de Tipo de Documento"" listURI=""urn:pe:gob:sunat:cpe:see:gem:catalogos:catalogo01"">" & pRecCab("idDocumento") & "</cbc:InvoiceTypeCode> " & vbCrLf
        End If
        strLinea = strLinea & "    <cbc:Note languageLocaleID=""1000"">" & pRecCab("totalLetras") & "</cbc:Note> " & vbCrLf
        'strLinea = strLinea & "    <cbc:DocumentCurrencyCode>" & pRecCab("idMoneda") & "</cbc:DocumentCurrencyCode> " & vbCrLf
        strLinea = strLinea & "    <cbc:DocumentCurrencyCode listID=""ISO 4217 Alpha"" listName=""Currency"" listAgencyName=""United Nations Economic Commission for Europe"">" & pRecCab("idMoneda") & "</cbc:DocumentCurrencyCode> " & vbCrLf
        If pDocumento = "01" Or pDocumento = "03" Then
        strLinea = strLinea & "    <cbc:LineCountNumeric>2</cbc:LineCountNumeric> " & vbCrLf
        End If
        
        If pDocumento = "07" Or pDocumento = "08" Then
        
        CSqlC = "Select A.TipoDocReferencia,A.SerieDocReferencia,A.NumDocReferencia " & _
                "From DocReferencia A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.TipoDocOrigen = '" & pDocumento & "' And A.SerieDocOrigen = '" & pSerie & "' " & _
                "And A.NumDocOrigen = '" & PNumero & "'"
        AbrirRecordset StrMsgError, Cn, rsref, CSqlC: If StrMsgError <> "" Then GoTo Err
        If Not rsref.EOF Then
            CIdDocumentoRef = Trim("" & rsref.Fields("TipoDocReferencia"))
            CIdSerieRef = Trim("" & rsref.Fields("SerieDocReferencia"))
            CIdNumeroRef = Trim("" & rsref.Fields("NumDocReferencia"))
        End If
        rsref.Close: Set rsref = Nothing
        
        strLinea = strLinea & "    <cac:DiscrepancyResponse> " & vbCrLf
        strLinea = strLinea & "        <cbc:ReferenceID>" & CIdSerieRef & "-" & CIdNumeroRef & "</cbc:ReferenceID> " & vbCrLf
        
        CSqlC = "Select A.GlsMotivoNCD,A.IdCodigoVE " & _
                "From MotivosNCD A " & _
                "Where A.IdMotivoNCD = '" & Trim("" & pRecCab.Fields("IdMotivoNCD")) & "'"
        AbrirRecordset StrMsgError, Cn, RsC, CSqlC: If StrMsgError <> "" Then GoTo Err
        If Not RsC.EOF Then
        strLinea = strLinea & "        <cbc:ResponseCode>" & Trim("" & RsC.Fields("IdCodigoVE")) & "</cbc:ResponseCode> " & vbCrLf
        End If
        RsC.Close: Set RsC = Nothing
        
        strLinea = strLinea & "        <cbc:Description>" & QuitarCaracteresEspeciales("" & pRecCab.Fields("ObsDocVentas")) & "</cbc:Description> " & vbCrLf
        strLinea = strLinea & "    </cac:DiscrepancyResponse> " & vbCrLf
        strLinea = strLinea & "    <cac:BillingReference> " & vbCrLf
        strLinea = strLinea & "        <cac:InvoiceDocumentReference> " & vbCrLf
        strLinea = strLinea & "            <cbc:ID>" & CIdSerieRef & "-" & CIdNumeroRef & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "            <cbc:DocumentTypeCode>" & CIdDocumentoRef & "</cbc:DocumentTypeCode> " & vbCrLf
        strLinea = strLinea & "        </cac:InvoiceDocumentReference> " & vbCrLf
        strLinea = strLinea & "    </cac:BillingReference> " & vbCrLf
        End If
        If Trim(PForm.txt_OrdenCompra.Text) <> "" And pDocumento = "01" Then
        'Agregar la orden de compra
        strLinea = strLinea & "    <cac:OrderReference> " & vbCrLf
        strLinea = strLinea & "        <cbc:ID><![CDATA[" & Trim(PForm.txt_OrdenCompra.Text) & "]]></cbc:ID> " & vbCrLf
        strLinea = strLinea & "    </cac:OrderReference> " & vbCrLf
        End If
                
        strLinea = strLinea & "    <cac:Signature> " & vbCrLf
        If pDocumento = "07" Or pDocumento = "08" Then
        strLinea = strLinea & "        <cbc:ID>" & strRUC & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "        <cbc:Note>" & STREMP & "</cbc:Note> " & vbCrLf
        Else
        strLinea = strLinea & "        <cbc:ID>IDSignSP</cbc:ID> " & vbCrLf
        End If
        strLinea = strLinea & "        <cac:SignatoryParty> " & vbCrLf
        strLinea = strLinea & "            <cac:PartyIdentification> " & vbCrLf
        strLinea = strLinea & "                <cbc:ID>" & strRUC & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "            </cac:PartyIdentification> " & vbCrLf
        strLinea = strLinea & "            <cac:PartyName> " & vbCrLf
        strLinea = strLinea & "                <cbc:Name><![CDATA[" & STREMP & "]]></cbc:Name> " & vbCrLf
        strLinea = strLinea & "            </cac:PartyName> " & vbCrLf
        strLinea = strLinea & "        </cac:SignatoryParty> " & vbCrLf
        strLinea = strLinea & "        <cac:DigitalSignatureAttachment> " & vbCrLf
        strLinea = strLinea & "            <cac:ExternalReference> " & vbCrLf
        strLinea = strLinea & "                <cbc:URI>#SignatureSP</cbc:URI> " & vbCrLf
        strLinea = strLinea & "            </cac:ExternalReference> " & vbCrLf
        strLinea = strLinea & "        </cac:DigitalSignatureAttachment> " & vbCrLf
        strLinea = strLinea & "    </cac:Signature> " & vbCrLf
        strLinea = strLinea & "    <cac:AccountingSupplierParty> " & vbCrLf
'        strLinea = strLinea & "        <cbc:CustomerAssignedAccountID>" & strRUC & "</cbc:CustomerAssignedAccountID> " & vbCrLf
'        strLinea = strLinea & "        <cbc:AdditionalAccountID>6</cbc:AdditionalAccountID> " & vbCrLf
'        strLinea = strLinea & "        <cac:Party> " & vbCrLf
'        strLinea = strLinea & "            <cac:PostalAddress> " & vbCrLf
'        strLinea = strLinea & "                <cbc:ID>" & pRecEmp("idDistrito") & "</cbc:ID> " & vbCrLf
'        strLinea = strLinea & "                <cbc:StreetName>" & STREMP & "</cbc:StreetName> " & vbCrLf
'        strLinea = strLinea & "                <cbc:CitySubdivisionName></cbc:CitySubdivisionName> " & vbCrLf
'        strLinea = strLinea & "                <cbc:CityName>" & strDpt & "</cbc:CityName> " & vbCrLf
'        strLinea = strLinea & "                <cbc:CountrySubentity>" & strPrv & "</cbc:CountrySubentity> " & vbCrLf
'        strLinea = strLinea & "                <cbc:District>" & pRecEmp("glsUbigeo") & "</cbc:District> " & vbCrLf
'        strLinea = strLinea & "                <cac:Country> " & vbCrLf
'        strLinea = strLinea & "                    <cbc:IdentificationCode>PE</cbc:IdentificationCode> " & vbCrLf
'        strLinea = strLinea & "                </cac:Country> " & vbCrLf
'        strLinea = strLinea & "            </cac:PostalAddress> " & vbCrLf
'        strLinea = strLinea & "            <cac:PartyLegalEntity> " & vbCrLf
'        strLinea = strLinea & "                <cbc:RegistrationName><![CDATA[" & STREMP & "]]></cbc:RegistrationName> " & vbCrLf
'        strLinea = strLinea & "            </cac:PartyLegalEntity> " & vbCrLf
'        strLinea = strLinea & "        </cac:Party> " & vbCrLf
'        strLinea = strLinea & "    </cac:AccountingSupplierParty> " & vbCrLf
'        strLinea = strLinea & "    <cac:AccountingCustomerParty> " & vbCrLf
'        strLinea = strLinea & "        <cbc:CustomerAssignedAccountID>" & pRecCab("RucCliente") & "</cbc:CustomerAssignedAccountID> " & vbCrLf
'        strLinea = strLinea & "        <cbc:AdditionalAccountID>" & pRecCab("idTipoDocIdentidad") & "</cbc:AdditionalAccountID> " & vbCrLf
'        strLinea = strLinea & "        <cac:Party> " & vbCrLf
'        strLinea = strLinea & "            <cac:PartyLegalEntity> " & vbCrLf
'        strLinea = strLinea & "                <cbc:RegistrationName>" & QuitarCaracteresEspeciales("" & pRecCab("GlsCliente")) & "</cbc:RegistrationName> " & vbCrLf
'        strLinea = strLinea & "            </cac:PartyLegalEntity> " & vbCrLf
'        strLinea = strLinea & "        </cac:Party> " & vbCrLf
'        strLinea = strLinea & "    </cac:AccountingCustomerParty> " & vbCrLf
        
        strLinea = strLinea & "    <cac:Party> " & vbCrLf
        strLinea = strLinea & "      <cac:PartyIdentification> " & vbCrLf
        strLinea = strLinea & "        <cbc:ID schemeID=""6"" schemeName=""SUNAT:Identificador de Documento de Identidad"" schemeAgencyName=""PE:SUNAT"" schemeURI=""urn:pe:gob:sunat:cpe:see:gem:catalogos:catalogo06"">" & strRUC & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "      </cac:PartyIdentification> " & vbCrLf
        strLinea = strLinea & "      <cac:PartyName> " & vbCrLf
        strLinea = strLinea & "        <cbc:Name>" & STREMP & "</cbc:Name> " & vbCrLf
        strLinea = strLinea & "      </cac:PartyName> " & vbCrLf
        strLinea = strLinea & "      <cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "        <cbc:RegistrationName>" & STREMP & "</cbc:RegistrationName> " & vbCrLf
        strLinea = strLinea & "        <cac:RegistrationAddress> " & vbCrLf
        strLinea = strLinea & "          <cbc:AddressTypeCode>" & pRecEmp("idDistrito") & "</cbc:AddressTypeCode> " & vbCrLf
        strLinea = strLinea & "        </cac:RegistrationAddress> " & vbCrLf
        strLinea = strLinea & "      </cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "    </cac:Party> " & vbCrLf
        strLinea = strLinea & "  </cac:AccountingSupplierParty> " & vbCrLf
        strLinea = strLinea & "  <cac:AccountingCustomerParty> " & vbCrLf
        strLinea = strLinea & "    <cac:Party> " & vbCrLf
        strLinea = strLinea & "      <cac:PartyIdentification> " & vbCrLf
        strLinea = strLinea & "        <cbc:ID schemeID=""" & pRecCab("idTipoDocIdentidad") & """ schemeName=""SUNAT:Identificador de Documento de Identidad"" schemeAgencyName=""PE:SUNAT"" schemeURI=""urn:pe:gob:sunat:cpe:see:gem:catalogos:catalogo06"">" & pRecCab("RucCliente") & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "      </cac:PartyIdentification> " & vbCrLf
        strLinea = strLinea & "      <cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "        <cbc:RegistrationName>" & QuitarCaracteresEspeciales("" & pRecCab("GlsCliente")) & "</cbc:RegistrationName> " & vbCrLf
        strLinea = strLinea & "      </cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "    </cac:Party> " & vbCrLf
        strLinea = strLinea & "  </cac:AccountingCustomerParty> " & vbCrLf
        
        strLinea = strLinea & "   <cac:TaxTotal> " & vbCrLf
        strLinea = strLinea & "       <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecCab("TotalIGVVenta"), "0.00") & "</cbc:TaxAmount> " & vbCrLf
        strLinea = strLinea & "       <cac:TaxSubtotal> " & vbCrLf
        strLinea = strLinea & "           <cbc:TaxableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecCab("TotalValorVenta"), "0.00") & "</cbc:TaxableAmount>" & vbCrLf
        strLinea = strLinea & "           <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecCab("TotalIGVVenta"), "0.00") & "</cbc:TaxAmount> " & vbCrLf
        strLinea = strLinea & "           <cac:TaxCategory> " & vbCrLf
        strLinea = strLinea & "               <cbc:ID schemeID=""UN/ECE 5305"" schemeName=""Tax Category Identifier"" schemeAgencyName=""United Nations Economic Commission for Europe"">" & IIf(Not bolExportacion, "S", "G") & "</cbc:ID>" & vbCrLf
        strLinea = strLinea & "               <cac:TaxScheme> " & vbCrLf
        If Not bolExportacion Then
        strLinea = strLinea & "                   <cbc:ID  schemeID=""UN/ECE 5153"" schemeAgencyID=""6"">1000</cbc:ID> " & vbCrLf
        Else
        strLinea = strLinea & "                   <cbc:ID schemeName=""Codigo de tributos"" schemeAgencyName=""PE:SUNAT"" schemeURI=""urn:pe:gob:sunat:cpe:see:gem:catalogos:catalogo05"">9995</cbc:ID>" & vbCrLf
        End If
        strLinea = strLinea & "                   <cbc:Name>" & IIf(Not bolExportacion, "IGV", "EXP") & "</cbc:Name> " & vbCrLf
        strLinea = strLinea & "                   <cbc:TaxTypeCode>" & IIf(Not bolExportacion, "VAT", "FRE") & "</cbc:TaxTypeCode> " & vbCrLf
        strLinea = strLinea & "               </cac:TaxScheme> " & vbCrLf
        strLinea = strLinea & "           </cac:TaxCategory> " & vbCrLf
        strLinea = strLinea & "    </cac:TaxSubtotal> " & vbCrLf
        
        If (pDocumento = "01" Or pDocumento = "03") And Not bolExportacion Then
        'INICIO
        strLinea = strLinea & "    <cac:TaxSubtotal> " & vbCrLf
        strLinea = strLinea & "         <cbc:TaxableAmount currencyID=""" & pRecCab("idMoneda") & """>0.00</cbc:TaxableAmount> " & vbCrLf
        strLinea = strLinea & "         <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>0.00</cbc:TaxAmount> " & vbCrLf
        strLinea = strLinea & "         <cac:TaxCategory> " & vbCrLf
        strLinea = strLinea & "            <cac:TaxScheme> " & vbCrLf
        strLinea = strLinea & "                <cbc:ID>9996</cbc:ID> " & vbCrLf
        strLinea = strLinea & "               <cbc:Name>GRA</cbc:Name> " & vbCrLf
        strLinea = strLinea & "               <cbc:TaxTypeCode>FRE</cbc:TaxTypeCode> " & vbCrLf
        strLinea = strLinea & "            </cac:TaxScheme> " & vbCrLf
        strLinea = strLinea & "            </cac:TaxCategory> " & vbCrLf
        strLinea = strLinea & "    </cac:TaxSubtotal> " & vbCrLf
        
        strLinea = strLinea & "    <cac:TaxSubtotal> " & vbCrLf
        strLinea = strLinea & "         <cbc:TaxableAmount currencyID=""" & pRecCab("idMoneda") & """>0.00</cbc:TaxableAmount> " & vbCrLf
        strLinea = strLinea & "         <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>0.00</cbc:TaxAmount> " & vbCrLf
        strLinea = strLinea & "         <cac:TaxCategory> " & vbCrLf
        strLinea = strLinea & "            <cac:TaxScheme> " & vbCrLf
        strLinea = strLinea & "               <cbc:ID>9997</cbc:ID> " & vbCrLf
        strLinea = strLinea & "               <cbc:Name>EXO</cbc:Name> " & vbCrLf
        strLinea = strLinea & "               <cbc:TaxTypeCode>VAT</cbc:TaxTypeCode> " & vbCrLf
        strLinea = strLinea & "            </cac:TaxScheme> " & vbCrLf
        strLinea = strLinea & "         </cac:TaxCategory> " & vbCrLf
        strLinea = strLinea & "    </cac:TaxSubtotal> " & vbCrLf
        ' FIN
        End If
        
        strLinea = strLinea & "    </cac:TaxTotal> " & vbCrLf
        
        If pDocumento = "08" Then
        strLinea = strLinea & "    <cac:RequestedMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecCab("TotalPrecioVenta"), "0.00") & "</cbc:PayableAmount> " & vbCrLf
        strLinea = strLinea & "    </cac:RequestedMonetaryTotal> " & vbCrLf
        Else
        strLinea = strLinea & "    <cac:LegalMonetaryTotal> " & vbCrLf
        If Val("" & pRecCab.Fields("TotalDescuentoGlobalGravado")) > 0 Or Val("" & pRecCab.Fields("TotalDescuentoGlobalExonerado")) > 0 Then
        strLinea = strLinea & "        <cbc:AllowanceTotalAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, Val("" & pRecCab.Fields("TotalDescuentoGlobalGravado")) + Val("" & pRecCab.Fields("TotalDescuentoGlobalExonerado"))), "0.00") & "</cbc:AllowanceTotalAmount> " & vbCrLf
        End If
        strLinea = strLinea & "        <cbc:LineExtensionAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalValorVenta")), "0.00") & "</cbc:LineExtensionAmount> " & vbCrLf
        strLinea = strLinea & "        <cbc:PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalPrecioVenta")), "0.00") & "</cbc:PayableAmount> " & vbCrLf
        strLinea = strLinea & "    </cac:LegalMonetaryTotal> " & vbCrLf
        End If
        
        csql = "Select If(Afecto = '1','" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "11", "10") & "','" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "16", "30") & "') Afecto,Cantidad,TotalVVNeto,PVUnit,TotalIGVNeto,glsProducto,idProducto,VVUnit from docventasdet where idEmpresa='" & glsEmpresa & "' and idDocumento='" & pDocumento & "' and idSerie='" & pSerie & "' and idDocVentas='" & PNumero & "' Order by item "
        Set pRecDet = New ADODB.Recordset
        pRecDet.Open csql, Cn
        Print #IntFile, strLinea
        item = 0
        'Detalle
        Do While Not pRecDet.EOF
            item = item + 1
            strLinea = ""
            strLinea = strLinea & "    <cac:" & CGlsCab & "Line> " & vbCrLf
            strLinea = strLinea & "        <cbc:ID>" & item & "</cbc:ID> " & vbCrLf
            strLinea = strLinea & "        <cbc:" & CGlsDet & "Quantity unitCode=""NIU"">" & pRecDet("Cantidad") & "</cbc:" & CGlsDet & "Quantity> " & vbCrLf
            strLinea = strLinea & "        <cbc:LineExtensionAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecDet("TotalVVNeto")), "0.00") & "</cbc:LineExtensionAmount> " & vbCrLf
            strLinea = strLinea & "        <cac:PricingReference> " & vbCrLf
            strLinea = strLinea & "            <cac:AlternativeConditionPrice> " & vbCrLf
            strLinea = strLinea & "                <cbc:PriceAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecDet("PVUnit")), "0.00") & "</cbc:PriceAmount> " & vbCrLf
            strLinea = strLinea & "                <cbc:PriceTypeCode>01</cbc:PriceTypeCode> " & vbCrLf
            strLinea = strLinea & "            </cac:AlternativeConditionPrice> " & vbCrLf
            If Val("" & pRecCab.Fields("IndVtaGratuita")) = 1 Then
            strLinea = strLinea & "            <cac:AlternativeConditionPrice> " & vbCrLf
            strLinea = strLinea & "                <cbc:PriceAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("PVUnit"), "0.00") & "</cbc:PriceAmount> " & vbCrLf
            strLinea = strLinea & "                <cbc:PriceTypeCode>02</cbc:PriceTypeCode> " & vbCrLf
            strLinea = strLinea & "            </cac:AlternativeConditionPrice> " & vbCrLf
            End If
            strLinea = strLinea & "        </cac:PricingReference> " & vbCrLf
            strLinea = strLinea & "        <cac:TaxTotal> " & vbCrLf
            strLinea = strLinea & "            <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("TotalIGVNeto"), "0.00") & "</cbc:TaxAmount> " & vbCrLf
            strLinea = strLinea & "            <cac:TaxSubtotal> " & vbCrLf
            strLinea = strLinea & "                <cbc:TaxableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("TotalVVNeto"), "0.00") & "</cbc:TaxableAmount>" & vbCrLf
            strLinea = strLinea & "                <cbc:TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("TotalIGVNeto"), "0.00") & "</cbc:TaxAmount> " & vbCrLf
            strLinea = strLinea & "                <cac:TaxCategory> " & vbCrLf
            strLinea = strLinea & "                    <cbc:Percent>18.00</cbc:Percent> " & vbCrLf
            strLinea = strLinea & "                    <cbc:TaxExemptionReasonCode listAgencyName=""PE: SUNAT"" listName=""SUNAT:Codigo de Tipo de Afectacion del IGV"" listURI=""urn:pe:gob:sunat:cpe:see:gem:catalogos:catalogo07"">" & IIf(Not bolExportacion, Trim("" & pRecDet.Fields("Afecto")), "40") & "</cbc:TaxExemptionReasonCode> " & vbCrLf
            strLinea = strLinea & "                    <cac:TaxScheme> " & vbCrLf
            strLinea = strLinea & "                        <cbc:ID>" & IIf(Not bolExportacion, "1000", "9995") & "</cbc:ID> " & vbCrLf
            strLinea = strLinea & "                        <cbc:Name>" & IIf(Not bolExportacion, "IGV", "EXP") & "</cbc:Name> " & vbCrLf
            strLinea = strLinea & "                        <cbc:TaxTypeCode>" & IIf(Not bolExportacion, "VAT", "FRE") & "</cbc:TaxTypeCode> " & vbCrLf
            strLinea = strLinea & "                    </cac:TaxScheme> " & vbCrLf
            strLinea = strLinea & "                </cac:TaxCategory> " & vbCrLf
            strLinea = strLinea & "            </cac:TaxSubtotal> " & vbCrLf
            strLinea = strLinea & "        </cac:TaxTotal> " & vbCrLf
            strLinea = strLinea & "        <cac:Item> " & vbCrLf
            strLinea = strLinea & "            <cbc:Description>" & QuitarCaracteresEspeciales("" & pRecDet("glsProducto")) & "</cbc:Description> " & vbCrLf
            'strLinea = strLinea & "            <cac:SellersItemIdentification> " & vbCrLf
            'strLinea = strLinea & "                <cbc:ID>" & pRecDet("idProducto") & "</cbc:ID> " & vbCrLf
            'strLinea = strLinea & "            </cac:SellersItemIdentification> " & vbCrLf
            If bolExportacion Then
            strLinea = strLinea & "            <cac:CommodityClassification> " & vbCrLf
            strLinea = strLinea & "                <cbc:ItemClassificationCode listID=""UNSPSC"" listAgencyName=""GS1 US"" listName=""Item Classification"">50000000</cbc:ItemClassificationCode>" & vbCrLf
            strLinea = strLinea & "            </cac:CommodityClassification> " & vbCrLf
            End If
            strLinea = strLinea & "        </cac:Item> " & vbCrLf
            strLinea = strLinea & "        <cac:Price> " & vbCrLf
            strLinea = strLinea & "            <cbc:PriceAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecDet("VVUnit")), "0.00") & "</cbc:PriceAmount> " & vbCrLf
            strLinea = strLinea & "        </cac:Price> " & vbCrLf
            strLinea = strLinea & "    </cac:" & CGlsCab & "Line> " & vbCrLf
            Print #IntFile, strLinea
            pRecDet.MoveNext
        Loop
        'strLinea = strLinea & "</Invoice>" & vbCrLf
        strLinea = "</" & CGlsCab & ">" & vbCrLf
        Print #IntFile, strLinea
        Close #IntFile
    End If
    MsgBox "Se creo el archivo XML, en unos momentos se enviara al sistema SOL", vbInformation
    
    Cn.Execute "Update DocVentas Set IndEnviadoSunat = 1 Where IdEmpresa = '" & glsEmpresa & "' And IdDocumento = '" & pDocumento & "' And IdSerie = '" & pSerie & "' And IdDocVentas = '" & PNumero & "'"
    
    RetVal = ShellExecute(PForm.hwnd, "Open", App.Path & "\Release\FEapp2.exe", "", "", 0)
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Close #IntFile
    Exit Sub
    Resume
End Sub

Public Sub DocumentoElectronicoGuia(PForm As Form, pDocumento As String, pSerie As String, PNumero As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim pRecCab As ADODB.Recordset
Dim pRecDet As ADODB.Recordset
Dim pRecEmp As ADODB.Recordset
Dim CCarpeta                        As String
Dim CGlsCab                         As String
Dim CSqlC                           As String
Dim rsref                           As New ADODB.Recordset
Dim CIdDocumentoRef                 As String
Dim CIdSerieRef                     As String
Dim CIdNumeroRef                    As String
Dim CGlsDet                         As String
Dim RsC                             As New ADODB.Recordset
Dim RetVal

    csql = "Select p.idTipoDocIdentidad,p.idDistrito,u.glsUbigeo,p.idPais,dv.idDocumento,dv.idSerie,dv.idDocVentas,dv.idMoneda,dv.IndVtaGratuita," & _
           "dv.TotalBaseImponible,dv.TotalExonerado,dv.TotalValorVenta,dv.TotalDescuento,dv.totalLetras,dv.fecEmision,dv.IdMotivoNCD,dv.ObsDocVentas," & _
           "dv.RucCliente,dv.GlsCliente,dv.TotalIGVVenta,dv.TotalPrecioVenta,dv.TotalDescuentoGlobalGravado,dv.TotalDescuentoGlobalExonerado,V.IdChofer,dv.Placa,dv.Llegada " & _
           "from docventas dv " & _
           "inner join personas p on dv.idPerCliente=p.idpersona " & _
           "inner join Ubigeo u On p.IdDistrito = u.IdDistrito And p.IdPais = u.IdPais " & _
           "Left Join Vehiculos V On dv.IdEmpresa = V.IdEmpresa And dv.IdVehiculo = V.IdVehiculo " & _
           "where dv.idEmpresa='" & glsEmpresa & "' and dv.idDocumento='" & pDocumento & "' and dv.idSerie='" & pSerie & "' and dv.idDocVentas='" & PNumero & "'"
    Set pRecCab = New ADODB.Recordset
    pRecCab.Open csql, Cn
    
    csql = "Select e.*, p.idDistrito, p.direccion, p.idPais, u.glsUbigeo from empresas e inner join personas p on e.idPersona=p.idPersona inner join Ubigeo u On p.IdDistrito = u.IdDistrito And p.IdPais = u.IdPais Where e.idEmpresa='" & glsEmpresa & "'; "
    Set pRecEmp = New ADODB.Recordset
    pRecEmp.Open csql, Cn
    
    If Not pRecCab.EOF And Not pRecEmp.EOF Then
        
        CGlsCab = "DespatchAdvice"
        CGlsDet = "Delivered"
        
        strRUC = "" & pRecEmp("RUC")
        STREMP = QuitarCaracteresEspeciales("" & "" & pRecEmp("GlsEmpresa"))
        strDpt = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Mid("" & pRecEmp("idDistrito"), 1, 2), False, " idProv = '00' And idPais = '" & pRecEmp("idPais") & "' ")
        strPrv = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Mid("" & pRecEmp("idDistrito"), 3, 2), False, " idDpto = '" & Mid("" & pRecEmp("idDistrito"), 1, 2) & "' and idProv <> '00' and idDist = '00' And idPais = '" & pRecEmp("idPais") & "' ")
        
        'CERTIFICADO DIGITAL
        'strRuta = App.Path & "\Temporales\" & strRUC & ".cer"
        'IntFile = FreeFile
        'strFirma = ""
        'Open strRuta For Input As #IntFile
        'Do While Not EOF(IntFile)
        '    Line Input #IntFile, strLinea
        '    strFirma = strFirma & "" & vbCrLf
        'Loop
        'Close #IntFile
        
        CCarpeta = leeParametro("CARPETA_XML_VE")
        
        strRuta = CCarpeta & "\" & strRUC & "-09-" & pRecCab("idSerie") & "-" & pRecCab("idDocVentas") & ".xml"
        IntFile = FreeFile
        Open strRuta For Output As #IntFile
        'Cabecera
        strLinea = ""
        strLinea = strLinea & "<?xml version=""1.0"" encoding=""UTF-8""?> " & vbCrLf
        strLinea = strLinea & "<" & CGlsCab & " " & vbCrLf
        strLinea = strLinea & "    xmlns=""urn:oasis:names:specification:ubl:schema:xsd:" & CGlsCab & "-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:ds=""http://www.w3.org/2000/09/xmldsig#"" " & vbCrLf
        strLinea = strLinea & "    xmlns:cac=""urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:cbc=""urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"" " & vbCrLf
        strLinea = strLinea & "    xmlns:ext=""urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2""> " & vbCrLf
        strLinea = strLinea & "    <ext:UBLExtensions> " & vbCrLf
        strLinea = strLinea & "        <ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "            <ext:ExtensionContent/> " & vbCrLf
        strLinea = strLinea & "        </ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "        <ext:UBLExtension " & vbCrLf
        strLinea = strLinea & "            xmlns="""" " & vbCrLf
        strLinea = strLinea & "            xmlns:ar=""urn:oasis:names:specification:ubl:schema:xsd:ApplicationResponse-2""> " & vbCrLf
        strLinea = strLinea & "            <ext:ExtensionContent> " & vbCrLf
                                               'FIRMA
        strLinea = strLinea & "            </ext:ExtensionContent> " & vbCrLf
        strLinea = strLinea & "        </ext:UBLExtension> " & vbCrLf
        strLinea = strLinea & "    </ext:UBLExtensions> " & vbCrLf
        strLinea = strLinea & "    <cbc:UBLVersionID>2.1</cbc:UBLVersionID> " & vbCrLf
        strLinea = strLinea & "    <cbc:CustomizationID>1.0</cbc:CustomizationID> " & vbCrLf
        strLinea = strLinea & "    <cbc:ID>" & pSerie & "-" & PNumero & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "    <cbc:IssueDate>" & Format(pRecCab.Fields("FecEmision"), "yyyy-mm-dd") & "</cbc:IssueDate> " & vbCrLf
        strLinea = strLinea & "    <cbc:IssueTime>" & pRecCab.Fields("FecEmision") & "</cbc:IssueTime> " & vbCrLf
        strLinea = strLinea & "    <cbc:DespatchAdviceTypeCode>09</cbc:DespatchAdviceTypeCode> " & vbCrLf
        strLinea = strLinea & "    <cbc:Note>" & pRecCab.Fields("ObsDocVentas") & "</cbc:Note> " & vbCrLf
        strLinea = strLinea & "    <cac:DespatchSupplierParty> " & vbCrLf
        strLinea = strLinea & "        <cbc:CustomerAssignedAccountID schemeID=""6"">" & strRUC & "</cbc:CustomerAssignedAccountID> " & vbCrLf
        strLinea = strLinea & "        <cac:Party> " & vbCrLf
        strLinea = strLinea & "            <cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "                <cbc:RegistrationName>" & STREMP & "</cbc:RegistrationName> " & vbCrLf
        strLinea = strLinea & "            </cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "        </cac:Party> " & vbCrLf
        strLinea = strLinea & "    </cac:DespatchSupplierParty> " & vbCrLf
        strLinea = strLinea & "    <cac:DeliveryCustomerParty> " & vbCrLf
        strLinea = strLinea & "        <cbc:CustomerAssignedAccountID schemeID=""6"">" & strRUC & "</cbc:CustomerAssignedAccountID> " & vbCrLf
        strLinea = strLinea & "        <cac:Party> " & vbCrLf
        strLinea = strLinea & "            <cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "                <cbc:RegistrationName>" & STREMP & "</cbc:RegistrationName> " & vbCrLf
        strLinea = strLinea & "            </cac:PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "        </cac:Party> " & vbCrLf
        strLinea = strLinea & "    </cac:DeliveryCustomerParty> " & vbCrLf
        strLinea = strLinea & "    <cac:Shipment> " & vbCrLf
        strLinea = strLinea & "        <cbc:ID>1</cbc:ID> " & vbCrLf
        strLinea = strLinea & "        <cbc:HandlingCode>04</cbc:HandlingCode> " & vbCrLf
        strLinea = strLinea & "        <cbc:Information/> " & vbCrLf
        strLinea = strLinea & "        <cac:ShipmentStage> " & vbCrLf
        strLinea = strLinea & "            <cbc:TransportModeCode>2</cbc:TransportModeCode> " & vbCrLf
        strLinea = strLinea & "            <cac:TransitPeriod> " & vbCrLf
        strLinea = strLinea & "                <cbc:StartDate>" & Format(pRecCab.Fields("FecEmision"), "yyyy-mm-dd") & "</cbc:StartDate> " & vbCrLf
        strLinea = strLinea & "            </cac:TransitPeriod> " & vbCrLf
        strLinea = strLinea & "            <cac:TransportMeans> " & vbCrLf
        strLinea = strLinea & "                <cac:RoadTransport> " & vbCrLf
        strLinea = strLinea & "                    <cbc:LicensePlateID>" & pRecCab.Fields("Placa") & "</cbc:LicensePlateID> " & vbCrLf
        strLinea = strLinea & "                </cac:RoadTransport> " & vbCrLf
        strLinea = strLinea & "            </cac:TransportMeans> " & vbCrLf
        strLinea = strLinea & "            <cac:DriverPerson> " & vbCrLf
        strLinea = strLinea & "                <cbc:ID schemeID=""1"">" & traerCampo("Personas", "Ruc", "IdPersona", Trim("" & pRecCab.Fields("IdChofer")), False) & "</cbc:ID> " & vbCrLf 'Falta DNI de Conductor
        strLinea = strLinea & "            </cac:DriverPerson> " & vbCrLf
        strLinea = strLinea & "        </cac:ShipmentStage> " & vbCrLf
        strLinea = strLinea & "        <cac:Delivery> " & vbCrLf
        strLinea = strLinea & "            <cac:DeliveryAddress> " & vbCrLf
        strLinea = strLinea & "                <cbc:ID>" & traerCampo("Personas", "IdDistrito", "IdPersona", glsPersonaEmpresa, False) & "</cbc:ID> " & vbCrLf 'Falta Ubigeo Llegada
        strLinea = strLinea & "                <cbc:StreetName>" & pRecCab.Fields("Llegada") & "</cbc:StreetName> " & vbCrLf 'Falta Direccion
        strLinea = strLinea & "            </cac:DeliveryAddress> " & vbCrLf
        strLinea = strLinea & "        </cac:Delivery> " & vbCrLf
        strLinea = strLinea & "        <cac:TransportHandlingUnit> " & vbCrLf
        strLinea = strLinea & "            <cbc:ID>" & pRecCab.Fields("Placa") & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "        </cac:TransportHandlingUnit> " & vbCrLf
        strLinea = strLinea & "        <cac:OriginAddress> " & vbCrLf
        strLinea = strLinea & "            <cbc:ID>" & traerCampo("Personas", "IdDistrito", "Ruc", "" & strRUC, False) & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "            <cbc:StreetName>" & traerCampo("Personas", "Direccion", "Ruc", "" & strRUC, False) & "</cbc:StreetName> " & vbCrLf
        strLinea = strLinea & "        </cac:OriginAddress> " & vbCrLf
        strLinea = strLinea & "    </cac:Shipment> " & vbCrLf
        
        csql = "Select If(Afecto = '1','" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "11", "10") & "','" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "16", "30") & "') Afecto,Cantidad,TotalVVNeto,PVUnit,TotalIGVNeto,glsProducto,idProducto,VVUnit from docventasdet where idEmpresa='" & glsEmpresa & "' and idDocumento='" & pDocumento & "' and idSerie='" & pSerie & "' and idDocVentas='" & PNumero & "' Order by item "
        Set pRecDet = New ADODB.Recordset
        pRecDet.Open csql, Cn
        Print #IntFile, strLinea
        item = 0
        'Detalle
        Do While Not pRecDet.EOF
        item = item + 1
        strLinea = ""
        strLinea = strLinea & "    <cac:" & CGlsCab & "Line> " & vbCrLf
        strLinea = strLinea & "        <cbc:ID>" & item & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "        <cbc:" & CGlsDet & "Quantity unitCode=""C62"">" & pRecDet("Cantidad") & "</cbc:" & CGlsDet & "Quantity> " & vbCrLf
        strLinea = strLinea & "        <cac:OrderLineReference> " & vbCrLf
        strLinea = strLinea & "            <cbc:LineID>" & item & "</cbc:LineID> " & vbCrLf
        strLinea = strLinea & "        </cac:OrderLineReference> " & vbCrLf
        strLinea = strLinea & "        <cac:Item> " & vbCrLf
        strLinea = strLinea & "            <cbc:Name>" & pRecDet("GlsProducto") & "</cbc:Name> " & vbCrLf
        strLinea = strLinea & "            <cac:SellersItemIdentification> " & vbCrLf
        strLinea = strLinea & "                <cbc:ID>" & pRecDet("IdProducto") & "</cbc:ID> " & vbCrLf
        strLinea = strLinea & "            </cac:SellersItemIdentification> " & vbCrLf
        strLinea = strLinea & "        </cac:Item> " & vbCrLf
        strLinea = strLinea & "    </cac:" & CGlsCab & "Line> " & vbCrLf
        
            Print #IntFile, strLinea
            pRecDet.MoveNext
        Loop
        
        strLinea = "</" & CGlsCab & ">" & vbCrLf
        Print #IntFile, strLinea
        Close #IntFile
    End If
    MsgBox "Se creo el archivo XML, en unos momentos se enviara al sistema SOL", vbInformation
    
    Cn.Execute "Update DocVentas Set IndEnviadoSunat = 1 Where IdEmpresa = '" & glsEmpresa & "' And IdDocumento = '" & pDocumento & "' And IdSerie = '" & pSerie & "' And IdDocVentas = '" & PNumero & "'"
    
    RetVal = ShellExecute(PForm.hwnd, "Open", App.Path & "\Release\FEapp2.exe", "", "", 0)
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Close #IntFile
    Exit Sub
    Resume
End Sub

Public Sub DocumentoElectronicoAceptaFacturaBoleta(PForm As Form, pDocumento As String, pSerie As String, PNumero As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim pRecCab As ADODB.Recordset
Dim pRecDet As ADODB.Recordset
Dim pRecEmp As ADODB.Recordset
Dim CCarpeta                        As String
Dim CGlsCab                         As String
Dim CSqlC                           As String
Dim rsref                           As New ADODB.Recordset
Dim CIdDocumentoRef                 As String
Dim CIdSerieRef                     As String
Dim CIdNumeroRef                    As String
Dim CGlsDet                         As String
Dim RsC                             As New ADODB.Recordset
Dim RetVal
Dim strRuta As String

    strRUC = traerCampo("Empresas", "Ruc", "idEmpresa", glsEmpresa, False)
    ' pregunto si el XML anterior fue generado
    If Val(PNumero) > 1 Then
        CCarpeta = leeParametro("CARPETA_XML_VE")
        
        'consulta SI EL xml YA EXISTE en la carpeta
        strRuta = CCarpeta & "\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & PNumero & ".xml"
        If Existe(strRuta) Then
            StrMsgError = "El XML del documento electrónico ya existe"
            GoTo Err
        End If
        'consulta SI EL xml YA EXISTE en la carpeta LOG
        strRuta = CCarpeta & "_WORK\outputs\R-" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & PNumero & ".xml"
        If Existe(strRuta) Then
            StrMsgError = "El XML del documento electrónico ya existe"
            GoTo Err
        End If
        'consulta la carpeta de envio
        strRuta = CCarpeta & "\" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & Format(Val(PNumero), "00000000") - 1 & ".xml"
        If Existe(strRuta) Then
            StrMsgError = "Debe generar el XML del documento electrónico anterior"
            GoTo Err
        End If
        'consulta la carpeta log
        strRuta = CCarpeta & "_WORK\outputs\R-" & strRUC & "-" & pDocumento & "-" & pSerie & "-" & Format(Val(PNumero), "00000000") - 1 & ".xml"
        If Existe(strRuta) Then
            StrMsgError = "Debe generar el XML del documento electrónico anterior"
            GoTo Err
        End If
    End If
    
    csql = "Select p.idTipoDocIdentidad,p.idDistrito,u.glsUbigeo,p.idPais,dv.idDocumento,dv.idSerie,dv.idDocVentas,dv.idMoneda,dv.IndVtaGratuita," & _
           "dv.TotalBaseImponible,dv.TotalExonerado,dv.TotalValorVenta,dv.TotalDescuento,dv.totalLetras,dv.fecEmision,dv.IdMotivoNCD,dv.ObsDocVentas," & _
           "dv.RucCliente,dv.GlsCliente,dv.TotalIGVVenta,dv.TotalPrecioVenta,dv.TotalDescuentoGlobalGravado,dv.TotalDescuentoGlobalExonerado " & _
           "from docventas dv inner join personas p on dv.idPerCliente=p.idpersona inner join Ubigeo u On p.IdDistrito = u.IdDistrito And p.IdPais = u.IdPais where dv.idEmpresa='" & glsEmpresa & "' and dv.idDocumento='" & pDocumento & "' and dv.idSerie='" & pSerie & "' and dv.idDocVentas='" & PNumero & "'"
    Set pRecCab = New ADODB.Recordset
    pRecCab.Open csql, Cn
    
    csql = "Select e.*, p.idDistrito, p.direccion, p.idPais, u.glsUbigeo from empresas e inner join personas p on e.idPersona=p.idPersona inner join Ubigeo u On p.IdDistrito = u.IdDistrito And p.IdPais = u.IdPais Where e.idEmpresa='" & glsEmpresa & "'; "
    Set pRecEmp = New ADODB.Recordset
    pRecEmp.Open csql, Cn
    
    If Not pRecCab.EOF And Not pRecEmp.EOF Then
        
        If pDocumento = "07" Then
            CGlsCab = "CreditNote"
            CGlsDet = "Credited"
        ElseIf pDocumento = "08" Then
            CGlsCab = "DebitNote"
            CGlsDet = "Debited"
        Else
            CGlsCab = "Invoice"
            CGlsDet = "Invoiced"
        End If
        
        strRUC = "" & pRecEmp("RUC")
        STREMP = QuitarCaracteresEspeciales("" & "" & pRecEmp("GlsEmpresa"))
        strDpt = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Mid("" & pRecEmp("idDistrito"), 1, 2), False, " idProv = '00' And idPais = '" & pRecEmp("idPais") & "' ")
        strPrv = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Mid("" & pRecEmp("idDistrito"), 3, 2), False, " idDpto = '" & Mid("" & pRecEmp("idDistrito"), 1, 2) & "' and idProv <> '00' and idDist = '00' And idPais = '" & pRecEmp("idPais") & "' ")
        
        'CERTIFICADO DIGITAL
        'strRuta = App.Path & "\Temporales\" & strRUC & ".cer"
        'IntFile = FreeFile
        'strFirma = ""
        'Open strRuta For Input As #IntFile
        'Do While Not EOF(IntFile)
        '    Line Input #IntFile, strLinea
        '    strFirma = strFirma & "" & vbCrLf
        'Loop
        'Close #IntFile
        
        CCarpeta = leeParametro("CARPETA_XML_VE")
        
        strRuta = CCarpeta & "\" & strRUC & "-" & pRecCab("idDocumento") & "-" & pRecCab("idSerie") & "-" & pRecCab("idDocVentas") & ".xml"
        IntFile = FreeFile
        Open strRuta For Output As #IntFile
        'Cabecera
        strLinea = ""
        strLinea = strLinea & "<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?> " & vbCrLf
        strLinea = strLinea & "<" & CGlsCab & " " & vbCrLf
        strLinea = strLinea & "    <IssueDate>" & Format(pRecCab("fecEmision"), "yyyy-mm-dd") & "</IssueDate> " & vbCrLf
        If pDocumento <> "07" And pDocumento <> "08" Then
        strLinea = strLinea & "    <InvoiceTypeCode>" & pRecCab("idDocumento") & "</InvoiceTypeCode> " & vbCrLf
        End If
        strLinea = strLinea & "    <ID>" & pRecCab("idSerie") & "-" & pRecCab("idDocVentas") & "</ID> " & vbCrLf
        strLinea = strLinea & "    <DocumentCurrencyCode>" & pRecCab("idMoneda") & "</DocumentCurrencyCode> " & vbCrLf
        
        If pDocumento = "07" Or pDocumento = "08" Then
        CSqlC = "Select A.TipoDocReferencia,A.SerieDocReferencia,A.NumDocReferencia " & _
                "From DocReferencia A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.TipoDocOrigen = '" & pDocumento & "' And A.SerieDocOrigen = '" & pSerie & "' " & _
                "And A.NumDocOrigen = '" & PNumero & "' And A.TipoDocReferencia = '01'"
        AbrirRecordset StrMsgError, Cn, rsref, CSqlC: If StrMsgError <> "" Then GoTo Err
        If Not rsref.EOF Then
            CIdDocumentoRef = Trim("" & rsref.Fields("TipoDocReferencia"))
            CIdSerieRef = Trim("" & rsref.Fields("SerieDocReferencia"))
            CIdNumeroRef = Trim("" & rsref.Fields("NumDocReferencia"))
        End If
        rsref.Close: Set rsref = Nothing
        
        strLinea = strLinea & "    <BillingReference> " & vbCrLf
        strLinea = strLinea & "        <InvoiceDocumentReference> " & vbCrLf
        strLinea = strLinea & "            <ID>" & CIdSerieRef & "-" & CIdNumeroRef & "</ID> " & vbCrLf
        strLinea = strLinea & "            <DocumentTypeCode>" & CIdDocumentoRef & "</DocumentTypeCode> " & vbCrLf
        strLinea = strLinea & "        </InvoiceDocumentReference> " & vbCrLf
        strLinea = strLinea & "    </BillingReference> " & vbCrLf
        
        End If
        strLinea = strLinea & "    <AccountingSupplierParty> " & vbCrLf
        strLinea = strLinea & "        <CustomerAssignedAccountID>" & strRUC & "</CustomerAssignedAccountID> " & vbCrLf
        strLinea = strLinea & "        <AdditionalAccountID>6</AdditionalAccountID> " & vbCrLf
        strLinea = strLinea & "        <Party> " & vbCrLf
        strLinea = strLinea & "            <PartyName> " & vbCrLf
        strLinea = strLinea & "                <Name>" & STREMP & "</Name> " & vbCrLf
        strLinea = strLinea & "            </PartyName> " & vbCrLf
        strLinea = strLinea & "            <PostalAddress> " & vbCrLf
        strLinea = strLinea & "                <ID>" & pRecEmp("idDistrito") & "</ID> " & vbCrLf
        strLinea = strLinea & "                <StreetName>" & STREMP & "</StreetName> " & vbCrLf
        strLinea = strLinea & "                <CitySubdivisionName></CitySubdivisionName> " & vbCrLf
        strLinea = strLinea & "                <CityName>" & strDpt & "</CityName> " & vbCrLf
        strLinea = strLinea & "                <CountrySubentity>" & strPrv & "</CountrySubentity> " & vbCrLf
        strLinea = strLinea & "                <District>" & pRecEmp("glsUbigeo") & "</District> " & vbCrLf
        strLinea = strLinea & "                <Country> " & vbCrLf
        strLinea = strLinea & "                    <IdentificationCode>PE</IdentificationCode> " & vbCrLf
        strLinea = strLinea & "                </Country> " & vbCrLf
        strLinea = strLinea & "            </PostalAddress> " & vbCrLf
        strLinea = strLinea & "            <PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "                <RegistrationName>" & STREMP & "</RegistrationName> " & vbCrLf
        strLinea = strLinea & "            </PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "        </Party> " & vbCrLf
        strLinea = strLinea & "    </AccountingSupplierParty> " & vbCrLf
        strLinea = strLinea & "    <AccountingCustomerParty> " & vbCrLf
        strLinea = strLinea & "        <CustomerAssignedAccountID>" & pRecCab("RucCliente") & "</CustomerAssignedAccountID> " & vbCrLf
        strLinea = strLinea & "        <AdditionalAccountID>" & pRecCab("idTipoDocIdentidad") & "</AdditionalAccountID> " & vbCrLf
        strLinea = strLinea & "        <Party> " & vbCrLf
        strLinea = strLinea & "            <PartyName> " & vbCrLf
        strLinea = strLinea & "                <Name>" & pRecCab("GlsCliente") & "</Name> " & vbCrLf
        strLinea = strLinea & "            </PartyName> " & vbCrLf
        'strLinea = strLinea & "            <PostalAddress> " & vbCrLf
        'strLinea = strLinea & "                <ID>" & pRecEmp("idDistrito") & "</ID> " & vbCrLf
        'strLinea = strLinea & "                <StreetName>" & pRecEmp("glsEmpresa") & "</StreetName> " & vbCrLf
        'strLinea = strLinea & "                <CitySubdivisionName></CitySubdivisionName> " & vbCrLf
        'strLinea = strLinea & "                <CityName>" & strDpt & "</CityName> " & vbCrLf
        'strLinea = strLinea & "                <CountrySubentity>" & strPrv & "</CountrySubentity> " & vbCrLf
        'strLinea = strLinea & "                <District>" & pRecEmp("glsUbigeo") & "</District> " & vbCrLf
        'strLinea = strLinea & "                <Country> " & vbCrLf
        'strLinea = strLinea & "                    <IdentificationCode>PE</IdentificationCode> " & vbCrLf
        'strLinea = strLinea & "                </Country> " & vbCrLf
        'strLinea = strLinea & "            </PostalAddress> " & vbCrLf
        strLinea = strLinea & "            <PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "                <RegistrationName>" & pRecCab("GlsCliente") & "</RegistrationName> " & vbCrLf
        strLinea = strLinea & "            </PartyLegalEntity> " & vbCrLf
        strLinea = strLinea & "        </Party> " & vbCrLf
        strLinea = strLinea & "    </AccountingCustomerParty> " & vbCrLf
        strLinea = strLinea & "    <TaxTotal> " & vbCrLf
        strLinea = strLinea & "        <TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecCab("TotalIGVVenta"), "0.00") & "</TaxAmount> " & vbCrLf
        strLinea = strLinea & "        <TaxSubtotal> " & vbCrLf
        strLinea = strLinea & "            <TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecCab("TotalIGVVenta"), "0.00") & "</TaxAmount> " & vbCrLf
        strLinea = strLinea & "            <TaxCategory> " & vbCrLf
        strLinea = strLinea & "                <TaxScheme> " & vbCrLf
        strLinea = strLinea & "                    <ID>1000</ID> " & vbCrLf
        strLinea = strLinea & "                    <Name>IGV</Name> " & vbCrLf
        strLinea = strLinea & "                    <TaxTypeCode>VAT</TaxTypeCode> " & vbCrLf
        strLinea = strLinea & "                </TaxScheme> " & vbCrLf
        strLinea = strLinea & "            </TaxCategory> " & vbCrLf
        strLinea = strLinea & "        </TaxSubtotal> " & vbCrLf
        strLinea = strLinea & "    </TaxTotal> " & vbCrLf
        strLinea = strLinea & "    <LegalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "        <PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalPrecioVenta")), "0.00") & "</PayableAmount> " & vbCrLf
        strLinea = strLinea & "    </LegalMonetaryTotal> " & vbCrLf
        csql = "Select If(Afecto = '1','" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "11", "10") & "','" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "16", "30") & "') Afecto,Cantidad,TotalVVNeto,PVUnit,TotalIGVNeto,glsProducto,idProducto,VVUnit from docventasdet where idEmpresa='" & glsEmpresa & "' and idDocumento='" & pDocumento & "' and idSerie='" & pSerie & "' and idDocVentas='" & PNumero & "' Order by item "
        Set pRecDet = New ADODB.Recordset
        pRecDet.Open csql, Cn
        Print #IntFile, strLinea
        item = 0
        'Detalle
        Do While Not pRecDet.EOF
        item = item + 1
        strLinea = ""
        strLinea = strLinea & "    <" & CGlsCab & "Line> " & vbCrLf
        strLinea = strLinea & "        <ID>" & item & "</ID> " & vbCrLf
        strLinea = strLinea & "        <" & CGlsDet & "Quantity unitCode=""NIU"">" & pRecDet("Cantidad") & "</" & CGlsDet & "Quantity> " & vbCrLf
        strLinea = strLinea & "        <LineExtensionAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecDet("TotalVVNeto")), "0.00") & "</LineExtensionAmount> " & vbCrLf
        strLinea = strLinea & "        <PricingReference> " & vbCrLf
        strLinea = strLinea & "            <AlternativeConditionPrice> " & vbCrLf
        strLinea = strLinea & "                <PriceAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecDet("PVUnit")), "0.00") & "</PriceAmount> " & vbCrLf
        strLinea = strLinea & "                <PriceTypeCode>01</PriceTypeCode> " & vbCrLf
        strLinea = strLinea & "            </AlternativeConditionPrice> " & vbCrLf
        If Val("" & pRecCab.Fields("IndVtaGratuita")) = 1 Then
        strLinea = strLinea & "            <AlternativeConditionPrice> " & vbCrLf
        strLinea = strLinea & "                <PriceAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("PVUnit"), "0.00") & "</PriceAmount> " & vbCrLf
        strLinea = strLinea & "                <PriceTypeCode>02</PriceTypeCode> " & vbCrLf
        strLinea = strLinea & "            </AlternativeConditionPrice> " & vbCrLf
        End If
        strLinea = strLinea & "        </PricingReference> " & vbCrLf
        strLinea = strLinea & "        <TaxTotal> " & vbCrLf
        strLinea = strLinea & "            <TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("TotalIGVNeto"), "0.00") & "</TaxAmount> " & vbCrLf
        strLinea = strLinea & "            <TaxSubtotal> " & vbCrLf
        strLinea = strLinea & "                <TaxAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(pRecDet("TotalIGVNeto"), "0.00") & "</TaxAmount> " & vbCrLf
        strLinea = strLinea & "                <TaxCategory> " & vbCrLf
        strLinea = strLinea & "                    <TaxExemptionReasonCode>" & Trim("" & pRecDet.Fields("Afecto")) & "</TaxExemptionReasonCode> " & vbCrLf
        strLinea = strLinea & "                    <TaxScheme> " & vbCrLf
        strLinea = strLinea & "                        <ID>1000</ID> " & vbCrLf
        strLinea = strLinea & "                        <Name>IGV</Name> " & vbCrLf
        strLinea = strLinea & "                        <TaxTypeCode>VAT</TaxTypeCode> " & vbCrLf
        strLinea = strLinea & "                    </TaxScheme> " & vbCrLf
        strLinea = strLinea & "                </TaxCategory> " & vbCrLf
        strLinea = strLinea & "            </TaxSubtotal> " & vbCrLf
        strLinea = strLinea & "        </TaxTotal> " & vbCrLf
        strLinea = strLinea & "        <Item> " & vbCrLf
        strLinea = strLinea & "            <Description>" & pRecDet("glsProducto") & "</Description> " & vbCrLf
        strLinea = strLinea & "            <SellersItemIdentification> " & vbCrLf
        strLinea = strLinea & "                <ID>" & pRecDet("idProducto") & "</ID> " & vbCrLf
        strLinea = strLinea & "            </SellersItemIdentification> " & vbCrLf
        strLinea = strLinea & "        </Item> " & vbCrLf
        strLinea = strLinea & "        <Price> " & vbCrLf
        strLinea = strLinea & "            <PriceAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecDet("VVUnit")), "0.00") & "</PriceAmount> " & vbCrLf
        strLinea = strLinea & "        </Price> " & vbCrLf
        strLinea = strLinea & "    </" & CGlsCab & "Line> " & vbCrLf
        
            Print #IntFile, strLinea
            pRecDet.MoveNext
        Loop
        strLinea = "    <AdditionalInformation> " & vbCrLf
        strLinea = strLinea & "        <AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "            <ID>1001</ID> " & vbCrLf
        strLinea = strLinea & "            <PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalBaseImponible")), "0.00") & "</PayableAmount> " & vbCrLf
        strLinea = strLinea & "        </AdditionalMonetaryTotal> " & vbCrLf
        
        If pDocumento <> "07" And pDocumento <> "08" Then
        strLinea = strLinea & "        <AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "            <ID>1002</ID> " & vbCrLf
        strLinea = strLinea & "            <PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, 0, pRecCab("TotalExonerado")), "0.00") & "</PayableAmount> " & vbCrLf
        strLinea = strLinea & "        </AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "        <AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "            <ID>1003</ID> " & vbCrLf
        strLinea = strLinea & "            <PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(0, "0.00") & "</PayableAmount> " & vbCrLf
        strLinea = strLinea & "        </AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "        <AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "            <ID>1004</ID> " & vbCrLf
        strLinea = strLinea & "            <PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & Format(IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, pRecCab("TotalValorVenta"), 0), "0.00") & "</PayableAmount> " & vbCrLf
        strLinea = strLinea & "        </AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "        <AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "            <ID>2005</ID> " & vbCrLf
        strLinea = strLinea & "            <PayableAmount currencyID=""" & pRecCab("idMoneda") & """>" & pRecCab("TotalDescuento") & "</PayableAmount> " & vbCrLf
        strLinea = strLinea & "        </AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "        <AdditionalMonetaryTotal> " & vbCrLf
        strLinea = strLinea & "            <ID>" & IIf(Val("" & pRecCab.Fields("IndVtaGratuita")) = 1, "1002", "1000") & "</ID> " & vbCrLf
        strLinea = strLinea & "            <Value>" & pRecCab("totalLetras") & "</Value> " & vbCrLf
        strLinea = strLinea & "        </AdditionalMonetaryTotal> " & vbCrLf
        End If
        
        strLinea = strLinea & "    </sac:AdditionalInformation> " & vbCrLf
        strLinea = strLinea & "</" & CGlsCab & ">" & vbCrLf
        Print #IntFile, strLinea
        Close #IntFile
        
    End If
    
    MsgBox "Se creo el archivo XML, en unos momentos se enviara al sistema de ACEPTA", vbInformation
    
    'Cn.Execute "Update DocVentas Set IndEnviadoSunat = 1 Where IdEmpresa = '" & glsEmpresa & "' And IdDocumento = '" & pDocumento & "' And IdSerie = '" & pSerie & "' And IdDocVentas = '" & PNumero & "'"
    
'    RetVal = ShellExecute(PForm.hwnd, "Open", App.Path & "\Release\FEapp2.exe", "", "", 0)
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Close #IntFile
    Exit Sub
    Resume
End Sub



Public Function Existe(strPath As String) As Boolean
    Dim intGet As Integer
On Local Error GoTo lblExiste
    PubStrErr = ""
    intGet = GetAttr(strPath)
    Existe = True
Exit Function
lblExiste:
    PubStrErr = Err.Number & " " & Err.Description
    Existe = False
End Function

Public Function QuitarCaracteresEspeciales(pPalabra As String) As String
    If InStr(1, pPalabra, "Á") > 0 Then pPalabra = Replace(pPalabra, "Á", "A")
    If InStr(1, pPalabra, "É") > 0 Then pPalabra = Replace(pPalabra, "É", "E")
    If InStr(1, pPalabra, "Í") > 0 Then pPalabra = Replace(pPalabra, "Í", "I")
    If InStr(1, pPalabra, "Ó") > 0 Then pPalabra = Replace(pPalabra, "Ó", "O")
    If InStr(1, pPalabra, "Ú") > 0 Then pPalabra = Replace(pPalabra, "Ú", "U")
    If InStr(1, pPalabra, "á") > 0 Then pPalabra = Replace(pPalabra, "á", "a")
    If InStr(1, pPalabra, "é") > 0 Then pPalabra = Replace(pPalabra, "é", "e")
    If InStr(1, pPalabra, "í") > 0 Then pPalabra = Replace(pPalabra, "í", "i")
    If InStr(1, pPalabra, "ó") > 0 Then pPalabra = Replace(pPalabra, "ó", "o")
    If InStr(1, pPalabra, "ú") > 0 Then pPalabra = Replace(pPalabra, "ú", "u")
    If InStr(1, pPalabra, "&") > 0 Then pPalabra = Replace(pPalabra, "&", "&amp;")
    If InStr(1, pPalabra, "<") > 0 Then pPalabra = Replace(pPalabra, "&", "&lt;")
    If InStr(1, pPalabra, ">") > 0 Then pPalabra = Replace(pPalabra, "&", "&gt;")
    If InStr(1, pPalabra, "\") > 0 Then pPalabra = Replace(pPalabra, "&", "&quot;")
    If InStr(1, pPalabra, "'") > 0 Then pPalabra = Replace(pPalabra, "&", "&apos;")
    If InStr(1, pPalabra, vbCrLf) > 0 Then pPalabra = Replace(pPalabra, vbCrLf, "")
    If InStr(1, pPalabra, vbTab) > 0 Then pPalabra = Replace(pPalabra, vbTab, "")
    QuitarCaracteresEspeciales = Trim(pPalabra)
End Function

