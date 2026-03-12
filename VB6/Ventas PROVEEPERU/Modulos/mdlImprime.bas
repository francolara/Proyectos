Attribute VB_Name = "mdlImprime"

Dim s As String
Dim sCabecera As String
Dim margenIzq As Long, longTotal As Long

'Const cSeparadores = " ªº\!|@#$%&/()=?¿'¡[]*+{}<>,.-;:_"
Dim Linea_l         As String
Dim xfila           As Integer
Dim ncol            As Integer
Const cSeparadores = " ,.;:_"
Private sSeparadores As String

Public Enum ePropperWrapConstants
    pwLeft = 0
    pwMid = 1
    pwRight = 2
End Enum
''''Public Sub ImprimeCodigoBarra(ByVal indTipo As Integer, ByVal codproducto As String, ByVal numVale As String, ByRef strMsgError As String, Optional ByVal dblCantidad As Double = 0)
'''''indTipo 0 = por producto,    1 = por bloque
''''On Error GoTo ERR
''''Dim objPrinter As New PrinterAPI.clsPrinter
''''Dim rsp As New ADODB.Recordset
''''Dim StrCodBarra As String
''''Dim BlnFoundPrinter As Boolean
''''Dim BlnFoundData As Boolean
''''Dim intPar As Long
''''Dim i As Integer
''''Dim intParTotal As Long
''''Dim indParTotal As Boolean
''''Dim strPrecio As String
''''Dim strTalla As String
''''Dim intCantidad As Integer
''''
''''    If (objPrinter.SetPrinter("Generica / Solo Texto") = False) Then
''''        If (objPrinter.SetPrinter("Generic / Text Only") = False) Then
''''            MsgBox "No se Encuentra instalada la Impresora " & NombreImpresora_sp & "o " & NombreImpresora_us, vbInformation
''''            Exit Sub
''''        End If
''''    End If
''''
''''    StrGlsEmpresa = traerCampo("empresas", "GlsEmpresa", "idEmpresa", glsEmpresa, False)
''''    If indTipo = 0 Then
''''        csql = "SELECT v.idProducto,p.GlsProducto,t.GlsTallaPeso,l.PVUnit,v.Cantidad " & _
''''               "FROM valesdet v,productos p, tallapeso t, preciosventa l " & _
''''               "WHERE v.idEmpresa = p.idEmpresa " & _
''''                 "AND v.idProducto = p.idProducto " & _
''''                 "AND p.idEmpresa = t.idEmpresa " & _
''''                 "AND p.idTallaPeso = t.idTallaPeso " & _
''''                 "AND p.idEmpresa = l.idEmpresa " & _
''''                 "AND p.idProducto = l.idProducto " & _
''''                 "AND p.idUMCompra = l.idUM " & _
''''                 "AND v.idValesCab = '" & numVale & "' AND p.idProducto = '" & codproducto & "' AND l.idLista = '" & glsListaVentas & "'"
''''        intCantidadTotal = 1
''''
''''    ElseIf indTipo = 1 Then
''''        csql = "SELECT v.idProducto,p.GlsProducto,t.GlsTallaPeso,l.PVUnit,v.Cantidad " & _
''''               "FROM valesdet v,productos p, tallapeso t, preciosventa l " & _
''''               "WHERE v.idEmpresa = p.idEmpresa " & _
''''                 "AND v.idProducto = p.idProducto " & _
''''                 "AND p.idEmpresa = t.idEmpresa " & _
''''                 "AND p.idTallaPeso = t.idTallaPeso " & _
''''                 "AND p.idEmpresa = l.idEmpresa " & _
''''                 "AND p.idProducto = l.idProducto " & _
''''                 "AND p.idUMCompra = l.idUM " & _
''''                 "AND v.idValesCab = '" & numVale & "' AND l.idLista = '" & glsListaVentas & "'"
''''        intCantidadTotal = Val("" & traerCampo("valesdet", "SUM(Cantidad)", "idSucursal", glsSucursal, True, " idValesCab = '" & numVale & "'"))
''''
''''    ElseIf indTipo = 2 Then
''''        csql = "SELECT p.idProducto,p.GlsProducto,t.GlsTallaPeso,l.PVUnit," & CStr(dblCantidad) & " AS Cantidad " & _
''''               "FROM productos p, tallapeso t, preciosventa l " & _
''''               "WHERE p.idEmpresa = t.idEmpresa " & _
''''                 "AND p.idTallaPeso = t.idTallaPeso " & _
''''                 "AND p.idEmpresa = l.idEmpresa " & _
''''                 "AND p.idProducto = l.idProducto " & _
''''                 "AND p.idUMCompra = l.idUM " & _
''''                 "AND p.idProducto = '" & codproducto & "' AND l.idLista = '" & glsListaVentas & "'"
'        intCantidadTotal = 1
'    End If
'
'    rsp.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
'    intPar = 0
'    If indTipo = 0 Or indTipo = 2 Then
'        If Not rsp.EOF Then
'            intParTotal = intCantidadTotal * Val("" & rsp.Fields("Cantidad"))
'        Else
'            strMsgError = "El producto no tiene precio"
'            GoTo ERR
'        End If
'    Else
'        intParTotal = intCantidadTotal
'    End If
'
'    indParTotal = True
'    If intParTotal Mod 2 Then
'        indParTotal = False
'    End If
'
'    Do While (Not rsp.EOF)
'        strPrecio = Format(rsp.Fields("PVUnit").Value, "##,##0.00")
'        strTalla = Trim$(rsp.Fields("GlsTallaPeso").Value)
'        intCantidad = Val("" & rsp.Fields("Cantidad"))
'
'        For i = 1 To intCantidad
'            intPar = intPar + 1
'            objPrinter.PrintDataLn Chr$(2) & "L"
'            objPrinter.PrintDataLn "A2"
'            objPrinter.PrintDataLn "D11"
'            objPrinter.PrintDataLn "z"
'            objPrinter.PrintDataLn "PN"
'            objPrinter.PrintDataLn "H10"
'
'            StrCodBarra = Trim$(rsp.Fields("idProducto").Value)
'            If intPar Mod 2 Then
'                objPrinter.PrintDataLn "191100300610140" & strPrecio 'PRECIO -30
'                objPrinter.PrintDataLn "191100100280010" & strTalla 'TALLA
'
'                objPrinter.PrintDataLn "191100100610070" & StrGlsEmpresa
'                objPrinter.PrintDataLn "191100100500010" & left(Trim$(rsp.Fields("GlsProducto").Value), 38)
'                objPrinter.PrintDataLn "191100300280133" & StrCodBarra
'                objPrinter.PrintDataLn "1e2201600010016B" & StrCodBarra
'
'                objPrinter.PrintDataLn "^01"     ' Numero de Copias
'                objPrinter.PrintDataLn "Q0001"   ' Numero de Etiquetas
'
'            Else
'                objPrinter.PrintDataLn "191100300610350" & strPrecio 'PRECIO
'                objPrinter.PrintDataLn "191100100280220" & strTalla 'TALLA
'
'                objPrinter.PrintDataLn "191100100610280" & StrGlsEmpresa
'                objPrinter.PrintDataLn "191100100500220" & left(Trim$(rsp.Fields("GlsProducto").Value), 38)
'                objPrinter.PrintDataLn "191100300280342" & StrCodBarra
'                objPrinter.PrintDataLn "1e2201600010228B" & StrCodBarra
'
'                objPrinter.PrintDataLn "^01"     ' Numero de Copias
'                objPrinter.PrintDataLn "Q0001"   ' Numero de Etiquetas
'                objPrinter.PrintDataLn "E"       ' Enviar la Impresion
'
'            End If
'
'            If indParTotal = False And intPar = intParTotal Then
'                objPrinter.PrintDataLn "E"       ' Enviar la Impresion
'            End If
'            BlnFoundData = True
'        Next
'        rsp.MoveNext
'    Loop
'    rsp.Close: Set rsp = Nothing
'
'    If (BlnFoundData) Then
'        Printer.EndDoc
'    Else
'        strMsgError = "No se Realizó ninguna Impresión, No hay Datos"
'        GoTo ERR
'    End If
'
'    Exit Sub
'
'ERR:
'    If rsp.State = 1 Then rsp.Close: Set rsp = Nothing
'    If strMsgError = "" Then strMsgError = ERR.Description
'End Sub
Public Sub imprimeDocVentas(strTD As String, strNumDoc As String, strSerie As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim wcont, wsum             As Integer, wfila As Integer, wcolu As Integer, intScale As Integer
Dim rst                     As New ADODB.Recordset, rstObj As New ADODB.Recordset, p As Object
Dim indPrinter              As Boolean, indterceros As Boolean, boolEst As Boolean
Dim strCampos               As String, strTipoFecDoc As String, strImprimeTicket As String, StrTipoTicket As String, strRUCCliente As String
Dim StrGlsCliente           As String, StrTotalIGV As String, StrGlsVendedorCampo As String, StrDirecCliente As String, StrMotivoTraslado As String
Dim StrCodigoCliente        As String, cselect As String, nfontletra As String, indventasterceros As String
Dim numEntreLineasAdicional As Integer, intRegistros As Integer, nFilRef As Integer, NTamanoLetra As Integer
Dim TbConsultaRef           As New ADODB.Recordset
Dim cSqlRef                 As String, IdTiendaCli As String, rstienda  As New ADODB.Recordset
Dim CodDistrito             As String, codPais As String, Gls_Pais As String, Gls_Depa As String, Gls_Prov As String, Gls_Distrito As String, Cad_Mysql As String
Dim impxsp                  As Integer, impysp As Integer, LongSp As Integer, RsC As New ADODB.Recordset
Dim CSqlC                   As String, strIdDocumento As String, strIdDocventas As String, strIdSerie As String, stridEmpresa As String
Dim stridSucursal           As String, stridPersona As String, strglspersona As String, strRUCPersona As String, stridTd As String, strcadll As String
Dim rsrecorset              As New ADODB.Recordset, RsD  As New ADODB.Recordset
Dim rucEmpresa              As String, StrTexto As String, strimpNF As String, strimpSF As String, StrimpFecf As String, FecRefNC As String
Dim numDocorigenNC          As String, serieDocorigenNC As String, tipodocorigenNC As String, rsref  As New ADODB.Recordset
Dim nfiladetalle            As Integer, nvarfila As Integer, nvarfilatotal As Integer
Dim intLong                 As Integer, intX As Integer, intY As Integer, intDec As Integer, Long_total As Integer, Long_Acumu As Integer, contadorImp As Integer
Dim strFechaDR              As String, strTipoDato As String, iddocumentoRef As String, GlsMarca As String
Dim TotFlete                As Double, strMtoSinDsc As Double, dblPorcDsc As Double, dblMtoTotal As Double
Dim strcaddref              As String, StrGlsMG As String, ccadenafecha As String, strGlosaDscto As String, strMO As String, dblMtoTotEnt As String, dblMtoTotVuelto As String
Dim CCampoDirCliente        As String, StrGlsMotivoTraslado As String, NReferencias  As Long, NItem As Long, IndVG As String, StrIndDirRecojo As String, Gls_Distritox   As String, Gls_Provx As String, Gls_Depax As String, StrDirRecojo As String

    CCampoDirCliente = "dirCliente"
    
    If leeParametro("IMPRIME_DIRECCION_1LINEA") = "S" Then CCampoDirCliente = ""
    
    rucEmpresa = traerCampo("empresas", "ruc", "idEmpresa", glsEmpresa, False)
     
    '--- SELECCIONAMOS IMPRESORA
    PredeterminaImpresora indPrinter, strTD, p, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Printer.ScaleMode = 6
    Printer.FontName = Trim((traerCampo("sucursales", "TipoLetra", "idSucursal", glsSucursal, True)))
    If (Val(traerCampo("empresas", "NTamanoLetra", "idEmpresa", glsEmpresa, False) & "")) > 0 Then Printer.FontSize = (Val(traerCampo("empresas", "NTamanoLetra", "idEmpresa", glsEmpresa, False) & ""))
    strTipoFecDoc = traerCampo("documentos", "TipoImpFecha", "idDocumento", strTD, False)
    strImprimeTicket = traerCampo("documentos", "indImprimeTicket", "idDocumento", strTD, False)
    
    numEntreLineasAdicional = Val("" & traerCampo("seriesdocumento", "espacioLineasImp", "idSerie", strSerie, True, " idsucursal = '" & glsSucursal & "' and iddocumento = '" & strTD & "'"))
    'Pierooooooo
    
    ImprimeDocVentasParte1 StrMsgError, strTD, strNumDoc, strSerie, StrTipoTicket, strRUCCliente, StrGlsCliente, StrTotalIGV, StrGlsVendedorCampo, _
    StrDirecCliente, StrMotivoTraslado, StrCodigoCliente, IdTiendaCli, RsD, StrGlsMotivoTraslado
    If StrMsgError <> "" Then GoTo Err
    ImprimeEtiquetas StrMsgError, strRUCCliente, strTD, strSerie, StrGlsCliente, StrTotalIGV, StrGlsMotivoTraslado, strNumDoc
    If StrMsgError <> "" Then GoTo Err
    
    '-------------------------------------------------------------------
    '--- CABECERA
    '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
    intRegistros = Val(traerCampo("objdocventas", "count(*)", "idDocumento", strTD, True, "tipoObj = 'C' and trim(GlsCampo) <> '' and indImprime = 1 and idSerie = '" & strSerie & "' "))
    csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'C' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "' ORDER BY IMPY,IMPX"
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rst.EOF
        strCampos = strCampos & "" & rst.Fields(0) & ","
        rst.MoveNext
    Loop
    strCampos = left(strCampos, Len(strCampos) - 1)
    rst.Close
    
    '---- AGREGADO EL 06/05/10 VENTAS A TERCEROS DEPENDIENDO SI EL CAMPO IND VENTAS TERCEROS ESTA CON 1
    indterceros = False
    If strTD = "86" Then
        indventasterceros = traerCampo("clientes", "indventasterceros", "idcliente", StrCodigoCliente, True)
        If indventasterceros = "1" Then
            indterceros = True
        End If
    End If
    
    '-----------------------------------------------------------------------------------------------------------------------------------
    '--- Traemos la data de lo campos seleccionados arriba
    '-----------------------------------------------------------------------------------------------------------------------------------
    csql = "SELECT " & strCampos & " , '" & StrGlsMotivoTraslado & "' As GlsMotivoNCD FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF Then
        For i = 0 To rst.Fields.Count - 1
            '--- Traemos datos de impresion por en nombre del campo de la tabla objDocventas
            If (strTD = "86" And rst.Fields(i).Name = "idMotivoTraslado") Or (strTD = "07" And rst.Fields(i).Name = "idMotivoTraslado") Then
                csql = "SELECT 'X' AS valor,'T' AS tipoDato, 1 AS impLongitud, impX, impY,0 AS Decimales,0 as intNumFilas FROM impMotivosTraslados WHERE idEmpresa = '" & glsEmpresa & "' and idDocumento = '" & strTD & "' AND idSerie = '" & strSerie & "' AND idMotivoTraslado = '" & Trim(rst.Fields(i) & "") & "'"
            Else
                csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales,intNumFilas,Identificador FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "' "
            End If
            
            If rst.Fields(i).Name = "GlsMoneda" And rucEmpresa = "20566047668" Then
                
                csql = "SELECT '" & IIf((rst.Fields(i) & "") = "SOLES", "SOLES", "DOLARES") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales,intNumFilas FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' "
                
            End If
            
            rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rstObj.EOF Then
                
                If rst.Fields(i).Name = "PeriodoAD" Then 'Telefono del Cliente
                    
                    ImprimeXY traerCampo("DocVentas A Inner Join Personas B On A.IdPerCliente = B.IdPersona", "B.Telefonos", "A.IdDocumento", strTD, True, "IdSerie = '" & strSerie & "' And IdDocVentas = '" & strNumDoc & "'"), "T", rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                
                ElseIf rst.Fields(i).Name = "idPerVendedorCampo" And rucEmpresa = "20505322674" Then
                    
                    ImprimeAbreviaturaVen StrMsgError, rstObj.Fields("Identificador"), (rst.Fields(i) & "")
                    If StrMsgError <> "" Then GoTo Err
                    
                ElseIf (rst.Fields(i).Name = "FecEmision" Or rst.Fields(i).Name = "FecIniTraslado") And strTipoFecDoc = "S" Then
                    
                    If rucEmpresa = "20462608056" Then
                        
                        ImprimeXY Format(Day(rstObj.Fields("valor")), "00") & " de " & strArregloMes(Val(Month(rstObj.Fields("valor")))) & " de " & Year(rstObj.Fields("valor")), "T", 25, rstObj.Fields("impY"), rstObj.Fields("impX") + 13, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    
                    Else
                    
                        '--- IMPRIME EL DIA
                        ImprimeXY Format(Day(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                        
                        If rucEmpresa = "20305948277" Then
                            
                            ImprimeXY strArregloMes(Val(Month(rstObj.Fields("valor")))) & "", "T", 10, rstObj.Fields("impY"), rstObj.Fields("impX") + 10, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                        
                            ImprimeXY right((Year(rstObj.Fields("valor"))), 2) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 65 - 3, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            
                        ElseIf rucEmpresa = "20296745317" Then
                            
                            ImprimeXY strArregloMes(Val(Month(rstObj.Fields("valor")))) & "", "T", 10, rstObj.Fields("impY"), rstObj.Fields("impX") + 25, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                        
                            ImprimeXY right((Year(rstObj.Fields("valor"))), 2) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 70, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        
                        ElseIf rucEmpresa = "20536550764" Then
                            
                            ImprimeXY strArregloMes(Val(Month(rstObj.Fields("valor")))) & "", "T", 10, rstObj.Fields("impY"), rstObj.Fields("impX") + 13, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                            
                            ImprimeXY right((Year(rstObj.Fields("valor"))), 2) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 57, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        
                        Else
                            
                            ImprimeXY strArregloMes(Val(Month(rstObj.Fields("valor")))) & "", "T", 10, rstObj.Fields("impY"), rstObj.Fields("impX") + 10, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                            
                            ImprimeXY right((Year(rstObj.Fields("valor"))), 2) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 30, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        
                        End If
                    
                    End If
                    
                    If StrMsgError <> "" Then GoTo Err
                ElseIf rst.Fields(i).Name = "FecEmision" And strTipoFecDoc = "C" Then   '---- PARA CADILLO
                    ccadenafecha = Format(Day(rstObj.Fields("valor")), "00") & " de " & strArregloMes(Val(Month(rstObj.Fields("valor")))) & " del " & right((Year(rstObj.Fields("valor"))), 4)
                    ImprimeXY ccadenafecha, "T", 30, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    
                ElseIf (rst.Fields(i).Name = "GlsVehiculo" Or rst.Fields(i).Name = "FecIniTraslado" Or rst.Fields(i).Name = "idPerChofer") And rucEmpresa = "20542001969" Then
                    MMGuiaRemison rstObj, rst, i, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    
                ElseIf rst.Fields(i).Name = "llegada" Then
                    '----------------------------------------------------------------------------------------------------------------------------
                    strIdDocumento = traerCampo("DocReferenciaEmpresas", "tipoDocReferencia", "idEmpresaOrigen", glsEmpresa, False, " idSucursalOrigen='" & glsSucursal & "' And tipoDocOrigen ='" & strTD & "' And numDocOrigen  = '" & strNumDoc & "' And serieDocOrigen='" & strSerie & "' ")
                    strIdDocventas = traerCampo("DocReferenciaEmpresas", "numDocReferencia", "idEmpresaOrigen", glsEmpresa, False, " idSucursalOrigen='" & glsSucursal & "' And tipoDocOrigen ='" & strTD & "' And numDocOrigen  = '" & strNumDoc & "' And serieDocOrigen='" & strSerie & "' ")
                    strIdSerie = traerCampo("DocReferenciaEmpresas", "serieDocReferencia", "idEmpresaOrigen", glsEmpresa, False, " idSucursalOrigen='" & glsSucursal & "' And tipoDocOrigen ='" & strTD & "' And numDocOrigen  = '" & strNumDoc & "' And serieDocOrigen='" & strSerie & "' ")
                    stridEmpresa = traerCampo("DocReferenciaEmpresas", "idEmpresaReferencia", "idEmpresaOrigen", glsEmpresa, False, " idSucursalOrigen='" & glsSucursal & "' And tipoDocOrigen ='" & strTD & "' And numDocOrigen  = '" & strNumDoc & "' And serieDocOrigen='" & strSerie & "' ")
                    stridSucursal = traerCampo("DocReferenciaEmpresas", "idSucursalReferencia", "idEmpresaOrigen", glsEmpresa, False, " idSucursalOrigen='" & glsSucursal & "' And tipoDocOrigen ='" & strTD & "' And numDocOrigen  = '" & strNumDoc & "' And serieDocOrigen='" & strSerie & "' ")
                    stridPersona = Trim("" & traerCampo("Docventas", "idPercliente", "idDocventas", strIdDocventas, False, " idEmpresa ='" & stridEmpresa & "' And idSucursal = '" & stridSucursal & "' And idDocumento = '" & strIdDocumento & "' And idSerie = '" & strIdSerie & "' "))
                    
                    stridTd = Trim("" & traerCampo("Docventas", "idTienda", "idDocventas", strIdDocventas, False, " idEmpresa ='" & stridEmpresa & "' And idSucursal = '" & stridSucursal & "' And idDocumento = '" & strIdDocumento & "' And idSerie = '" & strIdSerie & "' "))
                    
                    strglspersona = Trim("" & traerCampo("Personas", "GlsPersona", "idPersona", stridPersona, False))
                    strRUCPersona = Trim("" & traerCampo("Personas", "RUC", "idPersona", stridPersona, False))
                    
                    If Len(Trim(traerCampo("DocReferenciaEmpresas", "numDocOrigen", "idEmpresaOrigen", glsEmpresa, False, " idSucursalOrigen='" & glsSucursal & "' And tipoDocOrigen ='" & strTD & "' And numDocOrigen  = '" & strNumDoc & "' And serieDocOrigen='" & strSerie & "' "))) > 0 Then
                        If strIdDocumento = "01" Then
                            stridTd = Trim("" & traerCampo("Docventas", "idTienda", "idDocventas", strNumDoc, True, "idSucursal = '" & glsSucursal & "' And idDocumento = '" & strTD & "' And idSerie = '" & strSerie & "'"))
                            If Len(Trim(stridTd)) > 0 Then
                                CodDistrito = Trim("" & traerCampo("tiendascliente", "idDistrito", "idtdacli", stridTd, True))
                                codPais = Trim("" & traerCampo("tiendascliente", "idPais", "idtdacli", stridTd, True))
                            Else
                                CodDistrito = Trim("" & traerCampo("personas", "idDistrito", "idpersona", stridPersona, False))
                                codPais = Trim("" & traerCampo("personas", "idPais", "idpersona", stridPersona, False))
                            End If
                            
                        Else
                            If stridTd = "" Then
                                CodDistrito = Trim("" & traerCampo("personas", "idDistrito", "idpersona", stridPersona, False))
                                codPais = Trim("" & traerCampo("personas", "idPais", "idpersona", stridPersona, False))
                            Else
                                CodDistrito = Trim("" & traerCampo("tiendascliente", "idDistrito", "idtdacli", stridTd, False, "IdEmpresa = '" & stridEmpresa & "'"))
                                codPais = Trim("" & traerCampo("tiendascliente", "idPais", "idtdacli", stridTd, False, "IdEmpresa = '" & stridEmpresa & "'"))
                            End If
                        End If
                          
                        Cad_Mysql = " select idDpto, idProv FROM ubigeo where iddistrito = '" & CodDistrito & "' and idPais = '" & codPais & "' "
                        If rsrecorset.State = 1 Then rsrecorset.Close
                        rsrecorset.Open Cad_Mysql, Cn, adOpenStatic, adLockReadOnly
                                  
                        If Not rsrecorset.EOF Then
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                
                                Gls_Distrito = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", CodDistrito, False)
                                
                                ImprimeXY Gls_Distrito & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            End If
                              
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                              
                                Gls_Prov = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Trim("" & rsrecorset.Fields("idProv")), False, " idDpto = '" & Trim("" & rsrecorset.Fields("idDpto")) & "' and idProv <> '00' and idDist = '00'")
                              
                                ImprimeXY Gls_Prov & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            End If
                              
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                            
                                Gls_Depa = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Trim("" & rsrecorset.Fields("idDpto")), False, " idProv = '00'")
                                
                                ImprimeXY Gls_Depa & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                                  
                                '--- Imprime glosa cliente
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", "GlsPersona", True, " iddocumento = '" & strTD & "' and Campo = 'GlsPersona' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", "GlsPersona", True, " iddocumento = '" & strTD & "' and Campo = 'GlsPersona' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", "GlsPersona", True, " iddocumento = '" & strTD & "' and Campo = 'GlsPersona' "))
                              
                                ImprimeXY strglspersona & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                                  
                                '--- Imprim ruc cliente
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", "RUCPersona", True, " iddocumento = '" & strTD & "' and Campo = 'RUCPersona' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", "RUCPersona", True, " iddocumento = '" & strTD & "' and Campo = 'RUCPersona' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", "RUCPersona", True, " iddocumento = '" & strTD & "' and Campo = 'RUCPersona' "))
                              
                                ImprimeXY strRUCPersona & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            End If
                        End If
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    
                    Else
                        ' ----------------------------------------------------------------------------------------------------------------------------
                        If Len(Trim("" & traerCampo("objimpespeciales", "iddocumento", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' "))) > 0 Then
                            Gls_Pais = "": Gls_Depa = "": Gls_Prov = "": Gls_Distrito = "": impxsp = 0: impysp = 0
                            
                            If rucEmpresa = "20566047668" Or rucEmpresa = "20505322674" Then
                                CodDistrito = Trim("" & traerCampo("Personas", "idDistrito", "idPersona", StrCodigoCliente, False))
                                codPais = Trim("" & traerCampo("Personas", "idPais", "idPersona", StrCodigoCliente, False))
                            Else
                                CodDistrito = Trim("" & traerCampo("tiendascliente", "idDistrito", "idtdacli", IdTiendaCli & "' and idPersona='" & StrCodigoCliente, True))
                                codPais = Trim("" & traerCampo("tiendascliente", "idPais", "idtdacli", IdTiendaCli, True))
                            End If
                            
                            Cad_Mysql = " select idDpto, idProv " & _
                                        " FROM ubigeo " & _
                                        " where iddistrito = '" & CodDistrito & "' and idPais = '" & codPais & "' "
                            If rstienda.State = 1 Then rstienda.Close
                            rstienda.Open Cad_Mysql, Cn, adOpenStatic, adLockReadOnly
                        
                            If Not rstienda.EOF Then
                                If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))) > 0 Then
                                    impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                    impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                    LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                    
                                    Gls_Distrito = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", CodDistrito, False, " IDPAIS = '" & codPais & "' ")
                                    
                                    ImprimeXY Gls_Distrito & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                    If StrMsgError <> "" Then GoTo Err
                            End If
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                            
                                Gls_Prov = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Trim("" & rstienda.Fields("idProv")), False, " idDpto = '" & Trim("" & rstienda.Fields("idDpto")) & "' and idProv <> '00' and idDist = '00' AND IDPAIS = '" & codPais & "'")
                            
                                ImprimeXY Gls_Prov & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            End If
                            
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                            
                                Gls_Depa = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Trim("" & rstienda.Fields("idDpto")), False, " idProv = '00' AND IDPAIS = '" & codPais & "' ")
                                
                                ImprimeXY Gls_Depa & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            End If
                        End If
                        
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                                                
                    ElseIf rucEmpresa = "20430471750" Then
                        strcadll = rstObj.Fields("valor") & ""
                        ImprimeXY left(rstObj.Fields("valor"), 54) & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        ImprimeXY Mid(strcadll, 55, Len(strcadll)) & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + 4, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                    Else
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                End If
                
                ElseIf rst.Fields(i).Name = "Partida" Then
                    ' luis 15/04/2018
                    StrIndDirRecojo = Val(Trim("" & traerCampo("motivostraslados", "IndDirRecojo", "idMotivoTraslado", Trim("" & rst.Fields("idMotivoTraslado")), False)))

                    If Val(StrIndDirRecojo) = 1 Then
                        DireccionRecojo strTD, strSerie, strNumDoc, StrDirRecojo, Gls_Distritox, Gls_Provx, Gls_Depax, StrMsgError
                    Else
                        StrIndDirRecojo = "": Gls_Distritox = "": Gls_Provx = "": Gls_Depax = ""
                    End If
                   '********************************************************
                    If Len(Trim("" & traerCampo("objimpespeciales", "count(*)", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' "))) > 0 Then
                        Gls_Pais = "": Gls_Depa = "": Gls_Prov = "": Gls_Distrito = ""
                        
                        TraeDistrito StrMsgError, rucEmpresa, StrCodigoCliente, strTD, strSerie, strNumDoc, CodDistrito
                        If StrMsgError <> "" Then GoTo Err
                        'CodDistrito = Trim("" & traerCampo("Personas", "idDistrito", "idPersona", IIf(rucEmpresa = "20544632192", StrCodigoCliente, glsSucursal), False))
                        codPais = Trim("" & traerCampo("Personas", "idPais", "idPersona", IIf(rucEmpresa = "20544632192", StrCodigoCliente, glsSucursal), False))
                        Cad_Mysql = " select idDpto, idProv FROM ubigeo where iddistrito = '" & CodDistrito & "' and idPais = '" & codPais & "' "
                        If rstienda.State = 1 Then rstienda.Close
                        rstienda.Open Cad_Mysql, Cn, adOpenStatic, adLockReadOnly
                        If Not rstienda.EOF Then
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                           
                                Gls_Distrito = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", CodDistrito, False)
                                ImprimeXY IIf(Val(StrIndDirRecojo) = 1, Gls_Distritox, Gls_Distrito & ""), rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            End If
                            
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                            
                                Gls_Prov = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Trim("" & rstienda.Fields("idProv")), False, " idDpto = '" & Trim("" & rstienda.Fields("idDpto")) & "' and idProv <> '00' and idDist = '00'")
                                ImprimeXY IIf(Val(StrIndDirRecojo) = 1, Gls_Provx, Gls_Prov & ""), rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            End If
                            
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                
                                Gls_Depa = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Trim("" & rstienda.Fields("idDpto")), False, " idProv = '00'")
                                ImprimeXY IIf(Val(StrIndDirRecojo) = 1, Gls_Depax, Gls_Depa & ""), rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            End If
                        End If
                        rstienda.Close: Set rstienda = Nothing
                        If CCampoDirCliente = "dirCliente" And Len("" & rstObj.Fields("valor")) > rstObj.Fields("impLongitud") Then
                            ImprimeXY Mid(rstObj.Fields("valor") & "", 1, rstObj.Fields("impLongitud")), rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                            ImprimeXY Mid(rstObj.Fields("valor") & "", rstObj.Fields("impLongitud") + 1), rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + 2, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        Else
                            ImprimeXY IIf(Val(StrIndDirRecojo) = 1, StrDirRecojo, rstObj.Fields("valor") & ""), rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        End If
                        If StrMsgError <> "" Then GoTo Err
                                                                                            
                    Else
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                
                ElseIf (rst.Fields(i).Name = "TipoCambio" And strTD = "86") Then
                    
                    If rucEmpresa = "20305948277" Then
                        If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'NUMERO' "))) > 0 Then
                        
                            strimpNF = traerCampo("Docventas", "idDocVentas", "RIGHT(glsdocreferencia,8)", rst.Fields("idDocventas").Value, True, "MID(glsdocreferencia,5,3)= '" & rst.Fields("idSerie").Value & "'  And idSucursal ='" & glsSucursal & "'")
                            If Len("" & strimpNF) = 0 Then
                                strimpNF = Trim("" & traerCampo("docreferencia", "numDocReferencia", "tipoDocOrigen", "86", True, " idsucursal = '" & glsSucursal & "' and numDocOrigen = '" & rst.Fields("idDocventas").Value & "' and serieDocOrigen = '" & rst.Fields("idSerie").Value & "' "))
                            End If
                            
                            impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'NUMERO' "))
                            impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'NUMERO' "))
                            LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'NUMERO' "))
                            
                            ImprimeXY strimpNF & "", "T", LongSp, impysp, impxsp, 0, 0, StrMsgError
                        End If
                        
                        If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'SERIE' "))) > 0 Then
                            strimpSF = traerCampo("Docventas", "idserie", "MID(glsdocreferencia,5,3)", rst.Fields("idSerie").Value, True, "RIGHT(glsdocreferencia,8)='" & rst.Fields("idDocventas").Value & "'  And idSucursal ='" & glsSucursal & "'")
                            If Len("" & strimpSF) = 0 Then
                                strimpSF = Trim("" & traerCampo("docreferencia", "serieDocOrigen", "tipoDocOrigen", "86", True, " idsucursal = '" & glsSucursal & "' and numDocOrigen = '" & rst.Fields("idDocventas").Value & "' and serieDocOrigen = '" & rst.Fields("idSerie").Value & "' "))
                            End If
                            
                            impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'SERIE' "))
                            impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'SERIE' "))
                            LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'SERIE' "))
                            
                            ImprimeXY strimpSF & "", "T", LongSp, impysp, impxsp, 0, 0, StrMsgError
                        End If
                        
                        If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'FECHA' "))) > 0 Then
                        
                            StrimpFecf = Format(traerCampo("Docventas", "FecEmision", "MID(glsdocreferencia,5,3)", rst.Fields("idSerie").Value, True, "RIGHT(glsdocreferencia,8)='" & rst.Fields("idDocventas").Value & "'  And idSucursal ='" & glsSucursal & "'"), "dd/mm/yyyy")
                            
                            impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'FECHA' "))
                            impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'FECHA' "))
                            LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'FECHA' "))
                            
                            ImprimeXY StrimpFecf & "", "T", LongSp, impysp, impxsp, 0, 0, StrMsgError
                        End If
                    
                    ElseIf rucEmpresa = "20430471750" Then
                        '--- Se esta utilizando el campo del TC para imprimir el documento de referencia de la guia para Alsisac
                        '--- SE HA COMENTADO PARA PODER EJECUTAR, ERROR DE PROCEDIMIENTO DEMASIADO GRANDE
                        csql = "Select  Concat(d.AbreDocumento,'',dr.serieDocOrigen,'/',dr.numDocOrigen)  as Nro_Comp " & _
                                "From DocReferencia dr Inner Join Docventas dv   On dr.tipoDocReferencia = dv.idDocumento   And dr.numDocReferencia = dv.idDocventas   And dr.serieDocReferencia = dv.idSerie   And dr.idEmpresa = dv.idEmpresa " & _
                                "Inner Join Documentos d   On d.idDocumento = dr.tipoDocOrigen " & _
                                "Where Numdocreferencia = '" & strNumDoc & "' And TipoDocreferencia = '" & strTD & "' And SerieDocreferencia = '" & strSerie & "' And tipodocorigen not in (99)  And EstDocventas <> 'Anu' "
                        If rsref.State = 1 Then rsref.Close
                        rsref.Open csql, Cn, adOpenStatic, adLockReadOnly

                        If Not rsref.EOF Then
                            rsref.MoveFirst
                            Do While Not rsref.EOF
                                strimpSF = strimpSF + Space(1) + rsref.Fields("Nro_Comp").Value & "" + " --"
                                rsref.MoveNext
                            Loop

                            strimpSF = left(strimpSF, Len(strimpSF) - 2)
                            ImprimeXY strimpSF, rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                        End If
                        
                    Else
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                    
                ElseIf rst.Fields(i).Name = "FecPago" And strTD = "07" And leeParametro("IMPRIME_FECHA_FACTURA") = "1" Then 'imprime la fecha del documento de referencia
                        
                    strFechaDR = traerCampo("Docventas", "FecEmision", "idDocventas", "" & right(rst.Fields("GlsDocReferencia").Value, 8), True, "idSerie= '" & "" & Mid(rst.Fields("GlsDocReferencia").Value, 5, 3) & "' And idSucursal='" & glsSucursal & "' ")
                    ImprimeXY strFechaDR & "", "T", 100, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                
                    '--- Recuperamos el RUC para la empresa Apimas solo para clientes del Pais Bolivia Concatenamos al Ruc N.I.T.
                ElseIf strTD = "01" And rucEmpresa = "20305948277" And rst.Fields(i).Name = "RUCCliente" And traerCampo("Personas", "idPais", "Ruc", rstObj.Fields("valor"), False) = "02005" Then
                    ImprimeXY "N.I.T." & rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                
                '--- Se está utilizando el campo vehiculo para imprimir la fecha del documento de referencia  de NC para ITS-ahora tambien para Salas PQS 18/12/14 - ahora también para Nota de Debito Solvet 02/02/15
                ElseIf (strTD = "07" And rucEmpresa = "20257354041" And rst.Fields(i).Name = "idVehiculo") Or (strTD = "08" And rucEmpresa = "20566047668" And rst.Fields(i).Name = "FecPago") Or (strTD = "08" And rucEmpresa = "20504973990" And rst.Fields(i).Name = "FecPago") Then
                
                    numDocorigenNC = traerCampo("DocReferencia", "numDocReferencia", "numDocOrigen", strNumDoc, True, "  serieDocOrigen= '" & strSerie & "' And tipoDocOrigen = '" & strTD & "' And idSucursal = '" & glsSucursal & "'  ")
                    serieDocorigenNC = traerCampo("DocReferencia", "serieDocReferencia", "numDocOrigen", strNumDoc, True, "  serieDocOrigen= '" & strSerie & "' And tipoDocOrigen = '" & strTD & "' And idSucursal = '" & glsSucursal & "'  ")
                    tipodocorigenNC = traerCampo("DocReferencia", "tipoDocReferencia", "numDocOrigen", strNumDoc, True, "  serieDocOrigen= '" & strSerie & "' And tipoDocOrigen = '" & strTD & "' And idSucursal = '" & glsSucursal & "'  ")
                                                              
                    FecRefNC = traerCampo("Docventas", "FecEmision", "idDocventas", numDocorigenNC, True, "idSerie = '" & serieDocorigenNC & "'  And idSucursal ='" & glsSucursal & "' And idDocumento = '" & tipodocorigenNC & "'  ")
                    ImprimeXY FecRefNC & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + IIf(rucEmpresa = "20566047668", 0, wsum), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                
                '-----------------------------------------------------------------------------------------
                '---- AQUI TENER EN CUENTA PARA LOS OTROS CLIENTES - LA DIRECCION SE IMPRIME EN 02 LINEAS
                '-----------------------------------------------------------------------------------------
                ElseIf rst.Fields(i).Name = "ObsDocVentas" Or rst.Fields(i).Name = "ObsDocVentas2" Or rst.Fields(i).Name = CCampoDirCliente Then
                        xfila = Val(rstObj.Fields("impY"))
                        ncol = Val(rstObj.Fields("impX"))
                        Linea_l = rstObj.Fields("valor") & ""
                      
                        margenIzq = 0
                        If margenIzq < 0 Or margenIzq > 40 Then
                            margenIzq = 0
                        End If
                        
                        longTotal = Val("" & rstObj.Fields("impLongitud"))
                        If longTotal < margenIzq Or longTotal > 136 Then
                            longTotal = Val("" & rstObj.Fields("impLongitud"))
                        End If
                        
                        sCabecera = Linea_l
                        s = LoopPropperWrap(sCabecera, longTotal, pwLeft)
                        sCabecera = ""
                            
                        ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila, ncol, Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                            
                        xfila = xfila + 4
                        wsum = wsum + 4
                        nvarfila = nvarfila + 4
                        filas_detalle = 0
                        
                        If Len(Linea_l) <= Len(s) Then
                        Else
                            Do While Len(s) > 0
                                '--- Añadirle el margen izquierdo
                                sCabecera = sCabecera & Space$(margenIzq) & s & vbCrLf
                                s = LoopPropperWrap()
                                xMemo = s 'sCabecera
                                
                                ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila, ncol, Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                                xfila = xfila + 4
                                wsum = wsum + 4
                                nvarfila = nvarfila + 4
                                filas_detalle = filas_detalle + 4
                            Loop
                        End If
                
                Else
                    If indterceros = True Then
                        If rst.Fields(i).Name = "GlsCliente" Or rst.Fields(i).Name = "RUCCliente" Or rst.Fields(i).Name = "llegada" Then
                        Else
                            ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                        End If
                    Else
                        If rst.Fields(i).Name = "vtercerosCliente" Or rst.Fields(i).Name = "vterceroRuc" Or rst.Fields(i).Name = "vtercerosdireccion" Then
                        Else
                             ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                             If StrMsgError <> "" Then GoTo Err
                        End If
                    End If
                End If
            End If
            rstObj.Close
        Next
    End If
    rst.Close
    '-----------------------------------------------------------------------------------------------------------------------------------
    '--- DETALLE
    '-----------------------------------------------------------------------------------------------------------------------------------
    wcont = 1
    wsum = 0
    '--- SI LA EMPRESA ES SINCHI O PIC HACE UNA IMPRESION PARTICULAR
    If rucEmpresa = "20119041148" Or rucEmpresa = "20119040842" Or rucEmpresa = "20530611681" Then      ''' SINCHI      PIC     ACUICULTURA
        If glsSucursal <> "08090001" And (strTD = "01" Or strTD = "07") Then
            '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
            strCampos = ""
            csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "'  ORDER BY impY,impX"
            rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            Do While Not rst.EOF
                strCampos = strCampos & "d." & "" & rst.Fields(0) & ","
                rst.MoveNext
            Loop
            strCampos = left(strCampos, Len(strCampos) - 1)
            rst.Close
            
            '--- Traemos la data de los campos seleccionados arriba
            csql = "SELECT " & strCampos & " FROM docventasdet d, productos p " & _
                   "WHERE d.idProducto = p.idProducto AND d.idEmpresa = p.idEmpresa " & _
                   "AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' " & _
                   "AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "' AND p.idNivel <> '10010062' order by d.item"
            rst.Open csql, Cn, adOpenStatic, adLockReadOnly
            wsum = 0
            Do While Not rst.EOF
                For i = 0 To rst.Fields.Count - 1
                    '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
                    csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "' "
                    rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                    
                    Printer.FontSize = 9

                    If Not rstObj.EOF Then
                        strTipoDato = rstObj.Fields("tipoDato")
                        intLong = rstObj.Fields("impLongitud")
                        intX = rstObj.Fields("impX")
                        intY = rstObj.Fields("impY")
                        intDec = Val("" & rstObj.Fields("Decimales"))
                            
                        ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                            
                    End If
                    rstObj.Close
                    Printer.FontSize = IIf(NTamanoLetra > 0, NTamanoLetra, Printer.FontSize)
                Next
                rst.MoveNext
                wsum = wsum + IIf(intScale = 6, IIf(glsEmpresa = "03" And strSerie = "001", 5, 4), 1) + numEntreLineasAdicional
                
            Loop
            rst.Close
            
            csql = "SELECT TotalVVNeto FROM docventasdet d, productos p " & _
                   "WHERE d.idProducto = p.idProducto AND d.idEmpresa = p.idEmpresa " & _
                   "AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' " & _
                   "AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "' AND p.idNivel = '10010062' "
            rst.Open csql, Cn, adOpenStatic, adLockReadOnly
            
            Do While Not rst.EOF
                If glsEmpresa = "03" Then
                    ImprimeXY Format(Val(rst.Fields(0) & ""), "0.00"), "N", 10, 95, 97, 2, 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                Else
                    ImprimeXY Format(Val(rst.Fields(0) & ""), "0.00"), "N", 10, 95, 97, 2, 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                End If
                
                rst.MoveNext
            Loop
            rst.Close
            
        Else
            '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
            strCampos = ""
            csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "'  ORDER BY impY,impX"
            rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            Do While Not rst.EOF
                strCampos = strCampos & "" & rst.Fields(0) & ","
                rst.MoveNext
            Loop
            strCampos = left(strCampos, Len(strCampos) - 1)
            rst.Close
            
            '--- Traemos la data de los campos seleccionados arriba
            csql = "SELECT " & strCampos & " FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
            rst.Open csql, Cn, adOpenStatic, adLockReadOnly
            wsum = 0
            Do While Not rst.EOF
                For i = 0 To rst.Fields.Count - 1
                    '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
                    csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "' "
                    rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                    If Not rstObj.EOF Then
                        strTipoDato = rstObj.Fields("tipoDato")
                        intLong = rstObj.Fields("impLongitud")
                        intX = rstObj.Fields("impX")
                        intY = rstObj.Fields("impY")
                        intDec = Val("" & rstObj.Fields("Decimales"))
                            
                        ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                            
                    End If
                    rstObj.Close
                Next
                rst.MoveNext
                wsum = wsum + IIf(intScale = 6, 1, 1) + numEntreLineasAdicional
            Loop
            rst.Close
        End If
        
    Else
        'Medio Mundo Venta directa de productos imprime el código
        If (rucEmpresa = "20542001969" And strTD = "01") Then
            'Luis 15/04/2018
            'If traerCampo("DocventasDet", "idCentroCosto", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento = '" & strTD & "' And idSucursal = '" & glsSucursal & "' ") = "" Then
            '    csql = "Update objdocventas Set indImprime = '1' WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and idDocumento = '" & strTD & "' and idSerie = '999'   And GlsCampo = 'CodigoRapido'"
            '    Cn.Execute (csql)
                
            'Else
            '    Cn.Execute "Update objdocventas Set indImprime = '0' WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999'   And GlsCampo = 'CodigoRapido'"
            'End If
        End If
        
        '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
        strCampos = ""
        csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "'  ORDER BY impY,impX"
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        
        If rucEmpresa = "20504973990" Then
            'Luis 15/04/2018
            'Do While Not rst.EOF
            '    strCampos = strCampos & "d." & "" & rst.Fields(0) & ","
            '    rst.MoveNext
            'Loop
        Else
            Do While Not rst.EOF
                
                If rst.Fields(0) = "VVUnit" Or rst.Fields(0) = "TotalVVNeto" Or rst.Fields(0) = "PVUnit" Or rst.Fields(0) = "TotalPVNeto" Then
                    If Trim("" & traerCampo("DocVentas", "If(IndTransGratuitaMP = '1','1',IndTransGratuita)", "IdSucursal", glsSucursal, True, "IdDocumento = '" & strTD & "' And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strNumDoc & "'")) = "1" Then
                        strCampos = strCampos & "0.00 " & rst.Fields(0) & ","
                    Else
                        If rst.Fields(0) = "TotalVVNeto" And Trim("" & traerCampo("DocVentas", "IndVtaGratuita", "IdSucursal", glsSucursal, True, "IdDocumento = '" & strTD & "' And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strNumDoc & "'")) = "1" Then
                            strCampos = strCampos & "0.00 " & rst.Fields(0) & ","
                        Else
                            strCampos = strCampos & "" & rst.Fields(0) & ","
                        End If
                    End If
                Else
                    strCampos = strCampos & "" & rst.Fields(0) & ","
                End If
                rst.MoveNext
            Loop
        End If
        strCampos = left(strCampos, Len(strCampos) - 1)
        rst.Close
        
        If rucEmpresa = "20504973990" Then '--- SOLVET
            '--- Traemos la data de los campos seleccionados arriba
            'csql = "SELECT " & strCampos & " FROM docventasdet d, productos p WHERE d.idProducto = p.idProducto AND d.idEmpresa = p.idEmpresa AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "' AND p.idNivel <> '11080037' order by d.item"
            csql = "SELECT " & strCampos & " FROM docventasdet d, productos p WHERE d.idProducto = p.idProducto AND d.idEmpresa = p.idEmpresa AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "' order by d.item"
        
        Else '--- OTROS CLIENTES
            '--- Traemos la data de los campos seleccionados arriba
            csql = "SELECT " & strCampos & " FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' order by item"
        End If
        
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        wsum = 0
        NItem = 0
        Do While Not rst.EOF
            
            NItem = NItem + 1
            
            For i = 0 To rst.Fields.Count - 1
                '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
                csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "' "
                rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rstObj.EOF Then
                
                    strTipoDato = rstObj.Fields("tipoDato")
                    intLong = rstObj.Fields("impLongitud")
                    intX = rstObj.Fields("impX")
                    intY = rstObj.Fields("impY")
                    intDec = Val("" & rstObj.Fields("Decimales"))
                     
                    If rst.Fields(i).Name = "GlsProducto" Then
                        Long_total = 0
                        Long_Acumu = 0
                        contadorImp = 0
                        
                        strGlosaDscto = ""
                        
                        'Para MM si el item del Pro. tiene descuento imprime glosa
                        If rucEmpresa = "20542001969" And strTD = "01" Then
                            strMtoSinDsc = traerCampo("DocventasDet", "Round(TotalPVBruto,2)", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' And idProducto = '" & rst.Fields("idProducto").Value & "' ")
                            strMO = IIf(traerCampo("Docventas", "idMoneda", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' ") = "PEN", "S/", "U$$")
                            dblPorcDsc = traerCampo("DocventasDet", "PorDcto", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' And idProducto = '" & rst.Fields("idProducto").Value & "' ")
                            dblMtoTotal = traerCampo("DocventasDet", "Round(TotalVVNeto,2)", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' And idProducto = '" & rst.Fields("idProducto").Value & "' ")
                        
                            If traerCampo("DocventasDet", "PorDcto", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' And idProducto = '" & rst.Fields("idProducto").Value & "' ") > 0 Then
                                strGlosaDscto = "Al monto " & strMO & "" & strMtoSinDsc & " aplicar descuento del " & dblPorcDsc & "%, total " & strMO & "" & dblMtoTotal & " "
                            Else
                                strGlosaDscto = ""
                            End If
                            
                            Linea_l = rst.Fields(i) & "" + strGlosaDscto
                            
                        Else
                            
                            Linea_l = rst.Fields(i)
                            
                        End If
                        
                        Linea_l = Linea_l & Trim("" & TraeGlosaTransGratuita(StrMsgError, strTD, strSerie, strNumDoc, NItem))
                        If StrMsgError <> "" Then GoTo Err
                        
                        Long_total = Len(Trim(Linea_l)) ' + Len(Trim(strGlosaDscto))
                        nfiladetalle = 0: nvarfila = 0: nvarfilatotal = 0
                        xfila = Val(rstObj.Fields("impY")) + wsum
                        ncol = Val(rstObj.Fields("impX"))
                
                        margenIzq = 0
                        If margenIzq < 0 Or margenIzq > 40 Then
                            margenIzq = 0
                        End If
                        
                        longTotal = Val("" & rstObj.Fields("impLongitud"))
                        If longTotal < margenIzq Or longTotal > 136 Then
                            longTotal = Val("" & rstObj.Fields("impLongitud"))
                        End If
                        
                        sCabecera = Linea_l
                        s = LoopPropperWrap(sCabecera, longTotal, pwLeft)
                        sCabecera = ""
                        
                        ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila, ncol, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                        
                        xfila = xfila + 4
                        filas_detalle = 0
                        Long_Acumu = Long_Acumu + Len(s)
                        
                        'If Len(Linea_l) <= Len(s) Then
                        If Len(Linea_l) < Len(s) Then
                            nvarfila = nvarfila + 4
                            
                        Else
                            Do While Long_total > Long_Acumu
                               '--- Añadirle el margen izquierdo
                                nvarfila = nvarfila + 4
                                sCabecera = sCabecera & Space$(margenIzq) & s & vbCrLf
                                s = LoopPropperWrap()
                                xMemo = s
                                ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila, ncol, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                                xfila = xfila + 4
                                filas_detalle = filas_detalle + 4
                                Long_Acumu = Long_Acumu + Len(s) + Len(vbCrLf)
                                If Long_Acumu = Long_Acumu + Len(s) Then
                                    contadorImp = contadorImp + 1
                                End If
                                If contadorImp = 10 Then Exit Do
                            Loop
                            'nvarfila = nvarfila + IIf(rucEmpresa = "20430471750" Or rucEmpresa = "20266578572" Or rucEmpresa = "20552185286" Or rucEmpresa = "20552184981" Or rucEmpresa = "20566047668", 0, 4)
                        End If
                 
                    
                    ElseIf rst.Fields(i).Name = "idMarca" Then
                        If rucEmpresa = "20305948277" Then
                            GlsMarca = Trim("" & traerCampo("Marcas", "GlsMarca", "idmarca", rst.Fields(i), True))
                            ImprimeXY GlsMarca, strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                            
                        Else
                            ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                        End If
                        
                    Else
                        If rucEmpresa = "20305948277" Then
                            If rst.Fields(i).Name = "idProducto" Then
                                If Len(Trim(traerCampo("productosclientes", "Codigo", "idproducto", Trim("" & rst.Fields(i)), True, " idclIente = '" & StrCodigoCliente & "' "))) = 0 Then
                                    If Len(Trim("" & traerCampo("Productos", "CodigoRapido", "idproducto", rst.Fields(i), True))) = 0 Then
                                        ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                                        If StrMsgError <> "" Then GoTo Err
                                    Else
                                        ImprimeXY right(Trim("" & traerCampo("Productos", "CodigoRapido", "idproducto", rst.Fields(i), True)), 6) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                                        If StrMsgError <> "" Then GoTo Err
                                    End If
                                Else
                                    ImprimeXY Trim(traerCampo("productosclientes", "Codigo", "idproducto", Trim("" & rst.Fields(i)), True, " idclIente = '" & StrCodigoCliente & "' ")) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                                    If StrMsgError <> "" Then GoTo Err
                                End If
                            Else
                                ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            End If
                        Else
                            ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                        End If
                    End If
                 End If
                rstObj.Close
            Next
            
            If Trim(leeParametro("GENERA_VALE_FORMULA") & "") = "S" Then
                
                CSqlC = "Select B.IdFabricante,B.GlsProducto,A.Cantidad From DocVentasDetFormula A Inner Join Productos B On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdDocumento = '" & strTD & "' And A.IdSerie = '" & strSerie & "' And A.IdDocVentas = '" & strNumDoc & "'"
                RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
                
                Do While Not RsC.EOF
                    xfila = xfila + 4
                    ImprimeXY RsC.Fields("IdFabricante"), "T", 18, xfila, 0, 0, 0, StrMsgError
                    ImprimeXY RsC.Fields("GlsProducto"), "T", 58, xfila, 33, 0, 0, StrMsgError
                    ImprimeXY RsC.Fields("Cantidad"), "N", 10, xfila, 138, 2, 0, StrMsgError
                    
                    RsC.MoveNext
                Loop
                RsC.Close
            End If
            rst.MoveNext
            wsum = wsum + IIf(intScale = 6, IIf(rucEmpresa = "20504973990" Or rucEmpresa = "20305948277" Or rucEmpresa = "20257354041" Or rucEmpresa = "20430471750" Or rucEmpresa = "20511137137" Or rucEmpresa = "20509571792" Or rucEmpresa = "20542001969" Or rucEmpresa = "20544632192" Or rucEmpresa = "20388197804" Or rucEmpresa = "20552543549", 0, 4), 1) + numEntreLineasAdicional + 2 + nvarfila
        Loop
        rst.Close
        
    End If
    '------------------------------------------------------------------------------------------------
    If strImprimeTicket = "S" Then
        wsum = wsum + IIf(intScale = 6, 4, 1)
    Else
        wsum = 0
    End If
    
    boolEst = False
    strCampos = ""
    csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "' ORDER BY IMPY,IMPX "
    
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        
        If rst.Fields(0) = "TotalValorVenta" Or rst.Fields(0) = "TotalIGVVenta" Or rst.Fields(0) = "TotalPrecioVenta" Or rst.Fields(0) = "totalLetras" Then
            If Trim("" & traerCampo("DocVentas", "If(IndTransGratuitaMP = '1','1',IndTransGratuita)", "IdSucursal", glsSucursal, True, "IdDocumento = '" & strTD & "' And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strNumDoc & "'")) = "1" Then
                If rst.Fields(0) = "totalLetras" Then
                    strCampos = strCampos & "ConCat(If(IdDocumento = '03','','SON: '),'CERO Y 0/100 SOLES') " & rst.Fields(0) & ","
                Else
                    strCampos = strCampos & "0.00 " & rst.Fields(0) & ","
                End If
            Else
                strCampos = strCampos & "" & rst.Fields(0) & ","
            End If
        Else
            strCampos = strCampos & "" & rst.Fields(0) & ","
        End If
        rst.MoveNext
    Loop
    If Len(strCampos) > 0 Then
        strCampos = left(strCampos, Len(strCampos) - 1)
    End If
    rst.Close
    
    '-----------------------------------------------------------------------------------------------------------------------------------
    '--- Traemos la data de lo campos seleccionados arriba
    '-----------------------------------------------------------------------------------------------------------------------------------
    If Len(strCampos) > 0 Then
        csql = "SELECT " & strCampos & ",IdMoneda FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        
        If Not rst.EOF Then
            For i = 0 To rst.Fields.Count - 1
                '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
                If strTD = "86" And rst.Fields(i).Name = "idMotivoTraslado" Then
                    csql = "SELECT 'X' AS valor,'T' AS tipoDato, 1 AS impLongitud, impX, impY,0 AS Decimales FROM impMotivosTraslados WHERE idEmpresa = '" & glsEmpresa & "' AND idSerie = '" & strSerie & "' AND idMotivoTraslado = '" & Trim(rst.Fields(i) & "") & "'"
                Else
                    csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & IIf(intRegistros = 0, "999", strSerie) & "'"
                End If
                
                rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rstObj.EOF Then
                    If strImprimeTicket = "S" Then
                        If rst.Fields(i).Name = "TotalPrecioVenta" Then
                            If rst.Fields("IdMoneda") = "PEN" Then
                                ImprimeXY "Total :  S/ ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
                            Else
                                ImprimeXY "Total : US$. ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
                            End If
                        End If
                    End If
                    
                    If strImprimeTicket = "S" Then
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), intY + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        wsum = wsum + IIf(intScale = 6, 4, 1)
                        
                        'Si es Ticket imprime Total y Vuelto
                        If boolEst = False Then
                            dblMtoTotEnt = Trim("" & traerCampo("pagosdocventas p  INNER JOIN formaspagos f ON P.idFormadePago = f.idFormaPago AND p.idEmpresa = f.idEmpresa INNER JOIN  monedas m ON p.idMoneda = m.idMoneda", "Sum(p.MontoOri)", "p.idSucursal", glsSucursal, False, "p.idEmpresa = '" & glsEmpresa & "' AND p.idDocumento = '" & strTD & "' AND p.idDocVentas = '" & strNumDoc & "' AND p.idSerie = '" & strSerie & "' "))
                        
                            ImprimeXY "Efectivo :  " & IIf(rst.Fields("IdMoneda") = "PEN", "S/", "US$.") & " ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
                            
                            ImprimeXY Format(IIf(dblMtoTotEnt = "", "0.00", dblMtoTotEnt), "0.00") & "", "N", "10", intY + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            wsum = wsum + IIf(intScale = 6, 4, 1)
                            
                            If rst.Fields("IdMoneda") = "PEN" Then
                                ImprimeXY "Vuelto :  S/ ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
                            Else
                                ImprimeXY "Vuelto : US$. ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
                            End If
                            
                            dblMtoTotVuelto = Trim("" & traerCampo("movcajasdet m INNER JOIN monedas o ON m.idMoneda = o.idMoneda", "Sum(m.ValMonto)", "m.idDocumento", strTD, True, "m.idDocVentas = '" & strNumDoc & "' AND m.idSerie = '" & strSerie & "' AND m.idTipoMovCaja = '99990003' And m.idSucursal = '" & glsSucursal & "'"))
                            ImprimeXY Format(IIf(dblMtoTotVuelto = "", "0.00", dblMtoTotVuelto), "0.00") & "", "N", "10", intY + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            wsum = wsum + IIf(intScale = 6, 4, 1)
                            
                            boolEst = True
                        End If
                    Else
                        
                        IndVG = traerCampo("Docventas", "indVtaGratuita", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento='" & strTD & "'")
                        
                        If strTD = "86" Or strTD = "07" Then
                            If strTD = "86" And rucEmpresa = "20542001969" Then ' Para MM
                                MMGuiaRemison rstObj, rst, i, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            Else
                                ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            End If
                        
                        Else
                            If rucEmpresa = "20119041148" Or rucEmpresa = "20119040842" Or rucEmpresa = "20530611681" Then
                                
                                SinchiImpSp rst.Fields(i).Name, rstObj.Fields("valor"), rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), rstObj.Fields("Decimales"), wsum, iddocumentoRef, StrMsgError, strNumDoc, strSerie, strTD, rst.Fields("idMoneda").Value
                                If StrMsgError <> "" Then GoTo Err
                                                                    
                            ElseIf rucEmpresa = "20257354041" And strTD = "01" And rst.Fields(i).Name = "GlsDocReferencia" Then
                                cSqlRef = "Select SerieDocReferencia,NumDocReferencia " & _
                                            "From DocReferencia " & _
                                            "Where TipoDocOrigen = '" & strTD & "' And NumDocOrigen = '" & strNumDoc & "' And SerieDocOrigen = '" & strSerie & "' " & _
                                            "And IdEmpresa = '" & glsEmpresa & "' And TipoDocReferencia = '86' Order By SerieDocReferencia,NumDocReferencia"
                                
                                With TbConsultaRef
                                    .Open cSqlRef, Cn, adOpenStatic, adLockReadOnly
                                    If Not .EOF Then
                                        nFilRef = rstObj.Fields("impY") + wsum
                                        .MoveFirst
                                        Do While Not .EOF
                                            ImprimeXY .Fields("SerieDocReferencia") & "-" & .Fields("NumDocReferencia"), rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), nFilRef, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                            nFilRef = nFilRef + 4
                                            .MoveNext
                                        Loop
                                    End If
                                    .Close: Set TbConsultaRef = Nothing
                                End With
                            
                            ElseIf rst.Fields(i).Name = "GlsDocReferencia" Then
                                
                                If (rucEmpresa = "20509571792" And strTD = "01") Or (rucEmpresa = "20566047668" And strTD = "01") Or (rucEmpresa = "20505322674" And strTD = "01") Or (rucEmpresa = "20296745317" And strTD = "01") Or (rucEmpresa = "20462608056" And strTD = "01") Then
                                
                                    cSqlRef = "Select SerieDocReferencia,NumDocReferencia " & _
                                                "From DocReferencia " & _
                                                "Where TipoDocOrigen = '" & strTD & "' And NumDocOrigen = '" & strNumDoc & "' And SerieDocOrigen = '" & strSerie & "' " & _
                                                "And IdEmpresa = '" & glsEmpresa & "' And TipoDocReferencia = '86' Order By SerieDocReferencia,NumDocReferencia"
                                    
                                    With TbConsultaRef
                                        .Open cSqlRef, Cn, adOpenStatic, adLockReadOnly
                                        If Not .EOF Then
                                            NReferencias = Val(leeParametro("REFERENCIAS_POR_LINEA"))
                                            NReferencias = IIf(NReferencias = 0, 3, NReferencias)
                                            nFilRef = rstObj.Fields("impY") + wsum
                                            NVueltas = 0
                                            .MoveFirst
                                            Do While Not .EOF
                                                .MoveNext
                                                If Not .EOF Then
                                                    .MovePrevious
                                                    ImprimeXY Val("" & .Fields("NumDocReferencia")) & "/", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), nFilRef, rstObj.Fields("impX") + (NVueltas * 9), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                                Else
                                                    .MovePrevious
                                                    ImprimeXY Val("" & .Fields("NumDocReferencia")), rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), nFilRef, rstObj.Fields("impX") + (NVueltas * 9), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                                End If
                                                
                                                NVueltas = NVueltas + 1
                                                If NVueltas / NReferencias = 1 Then
                                                    nFilRef = nFilRef + 4
                                                    NVueltas = 0
                                                End If
                                                .MoveNext
                                            Loop
                                        End If
                                        .Close: Set TbConsultaRef = Nothing
                                    End With
                                
                                Else
                                    
                                    ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                    If StrMsgError <> "" Then GoTo Err
                                    
                                End If
                                
                                '--- Para Cadillo
                            ElseIf (rucEmpresa = "20100664179" Or rucEmpresa = "20381829279") And strTD = "01" And rst.Fields(i).Name = "GlsDocReferencia" Then
                                cSqlRef = "Select Concat(AbreDocumento,' ',SerieDocReferencia,'-',NumDocReferencia) as NumDocReferencia, SerieDocReferencia " & _
                                            "From DocReferencia r Inner Join Documentos d   On r.TipoDocReferencia = d.idDocumento " & _
                                            "Where TipoDocOrigen = '" & strTD & "' And NumDocOrigen = '" & strNumDoc & "' And SerieDocOrigen = '" & strSerie & "' " & _
                                            "And IdEmpresa = '" & glsEmpresa & "' And TipoDocReferencia = '86' Order By SerieDocReferencia,NumDocReferencia"
                                
                                With TbConsultaRef
                                    .Open cSqlRef, Cn, adOpenStatic, adLockReadOnly
                                    If Not .EOF Then
                                        nFilRef = rstObj.Fields("impY") + wsum
                                        NVueltas = 0
                                        .MoveFirst
                                        Do While Not .EOF
                                            .MoveNext
                                            If Not .EOF Then
                                                .MovePrevious
                                                ImprimeXY "" & .Fields("NumDocReferencia") & "/", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), nFilRef, rstObj.Fields("impX") + (NVueltas * 25), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                            Else
                                                .MovePrevious
                                                ImprimeXY "" & .Fields("NumDocReferencia"), rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), nFilRef, rstObj.Fields("impX") + (NVueltas * 25), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                            End If
                                            
                                            NVueltas = NVueltas + 1
                                            If NVueltas / 3 = 1 Then
                                                nFilRef = nFilRef + 4
                                                NVueltas = 0
                                            End If
                                            .MoveNext
                                        Loop
                                    End If
                                    .Close: Set TbConsultaRef = Nothing
                                End With
                            ElseIf strTD = "01" And rst.Fields(i).Name = "GlsDocReferencia" Then
                                strcaddref = rstObj.Fields("valor") & ""
                                ImprimeXY left(rstObj.Fields("valor"), 36) & "", "T", 100, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                ImprimeXY Trim(Mid(strcaddref, 36, Len(strcaddref)) & ""), "T", 100, rstObj.Fields("impY") + 3, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            ElseIf rst.Fields(i).Name = "GlsDocReferencia" And rucEmpresa = "20542001969" Then ' Para MM
                                MMGuiaRemison rstObj, rst, i, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                            ElseIf rucEmpresa = "20566047668" And rst.Fields(i).Name = "TotalValorVenta" And IndVG = "1" Then  'Salas Evalua Campo para las muestras gratuitas
                                ImprimeXY "0.00" & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            ElseIf rucEmpresa = "20566047668" And rst.Fields(i).Name = "TotalIGVVenta" And IndVG = "1" Then  'Salas Evalua Campo para las muestras gratuitas
                                ImprimeXY "0.00" & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            ElseIf (rucEmpresa = "20504973990" Or rucEmpresa = "20566047668") And rst.Fields(i).Name = "TotalPrecioVenta" And IndVG = "1" Then  'Solvet y Salas Evalua Campo para las muestras gratuitas
                                ImprimeXY "0.00" & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            ElseIf (rucEmpresa = "20504973990" Or rucEmpresa = "20566047668") And rst.Fields(i).Name = "totalLetras" And IndVG = "1" Then   'Solvet y Salas Evalua Campo para las muestras gratuitas
                                ImprimeXY "SON: 0.00 " & IIf(rst.Fields("idMoneda").Value = "PEN", "SOLES", "DOLARES AMERICANOS") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            Else
                                ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            End If
                        End If
                    End If
                    If StrMsgError <> "" Then GoTo Err
                End If
                rstObj.Close
            Next
        End If
        rst.Close
    End If
    
    '-----------------------------------------------------------------------------------------------------------------------------------
    '--- IMPRIME ETIQUETAS FINAL
    '-----------------------------------------------------------------------------------------------------------------------------------
    wsum = IIf(strImprimeTicket = "S", wsum + IIf(intScale = 6, 4, 1), 0)
    
    csql = "SELECT Etiqueta,impX,impY,indDirSucursal,indSerieEtiquetera,indSoloTicketFactura,indRUCCliente,indRazonSocial,indIGVTotal,indHora,indVendedor,indDirecCliente,indUsuario " & _
           "FROM objetiquetasdoc " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & strTD & "' AND idSerie = '" & _
           strSerie & "'  AND tipoObj = 'T' ORDER BY impY,impX"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    Do While Not rst.EOF
         StrTexto = ""
        If (rst.Fields(6) & "") = 0 And (rst.Fields(7) & "") = 0 And (rst.Fields(8) & "") = 0 And (rst.Fields(9) & "") = 0 And (rst.Fields(10) & "") = 0 Then
            If (rst.Fields(5) & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields(0) & ""
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = rst.Fields(0) & ""
                End If
            End If
        End If
                        
        If (rst.Fields(3) & "") = 1 Then StrTexto = StrTexto + traerDireccionSucursal
        If (rst.Fields(4) & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields(12) & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        '08001 FACTURA -- '08002 BOLETA
        If (rst.Fields(6) & "") = 1 Then
            If (rst.Fields(5) & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = (rst.Fields(0) & "") + strRUCCliente
                End If
            Else
                StrTexto = (rst.Fields(0) & "") + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 And (((rst.Fields("indSoloTicketFactura") & "") = 1 And StrTipoTicket = "08001") Or StrTipoTicket = "08002") Then StrTexto = (rst.Fields(0) & "") + StrGlsCliente  '2
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = (rst.Fields(0) & "") + StrTotalIGV
                End If
            Else
                StrTexto = (rst.Fields(0) & "") + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indHora") & "") = 1 Then
            StrTexto = (rst.Fields(0) & "") + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        End If
                           
        If (rst.Fields("indVendedor") & "") And (((rst.Fields("indSoloTicketFactura") & "") = 1 And StrTipoTicket = "08001") Or (StrTipoTicket = "08002")) = 1 Then StrTexto = (rst.Fields(0) & "") + StrGlsVendedorCampo
        
        If (rst.Fields("indDirecCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = (rst.Fields(0) & "") + StrDirecCliente
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = (rst.Fields(0) & "") + StrDirecCliente
                End If
            End If
        End If
        
        If Len(Trim(StrTexto)) > 0 Then
            If strImprimeTicket = "S" And strTD = "12" Then
                ImprimeXY StrTexto, "T", Len(StrTexto), intY + wsum, rst.Fields("impX"), 0, 0, StrMsgError
                wsum = wsum + IIf(intScale = 6, 4, 1)
            Else
                ImprimeXY StrTexto, "T", Len(StrTexto), rst.Fields("impY"), rst.Fields("impX"), 0, 0, StrMsgError
                wsum = wsum + IIf(intScale = 6, 4, 1)
            End If
        End If
        If StrMsgError <> "" Then GoTo Err
        rst.MoveNext
    Loop
    rst.Close
                                                              
    imprimeDocVentas_2 strTD, strNumDoc, strSerie, StrMsgError
    If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Printer.KillDoc
    Exit Sub
End Sub

Private Sub TraeDistrito(StrMsgError As String, rucEmpresa As String, PIdCliente As String, PIdDocumento As String, PIdSerie As String, PIdDocVentas As String, CodDistrito As String)
On Error GoTo Err
Dim CSqlC                       As String
Dim RsC                         As New ADODB.Recordset

    If rucEmpresa = "20544632192" Then
    
        CSqlC = "Select B.RucCliente,B.IdSucursal " & _
                "From DocReferenciaEmpresas A " & _
                "Inner Join DocVentas B " & _
                    "On '01' = B.IdEmpresa And A.TipoDocReferencia = B.IdDocumento And A.SerieDocReferencia = B.IdSerie And A.NumDocReferencia = B.IdDocVentas " & _
                "Where A.TipoDocOrigen = '" & PIdDocumento & "' And A.SerieDocOrigen = '" & PIdSerie & "' And A.NumDocOrigen = '" & PIdDocVentas & "'"
        AbrirRecordset StrMsgError, Cn, RsC, CSqlC: If StrMsgError <> "" Then GoTo Err
        If Not RsC.EOF Then
            
            If Trim("" & RsC.Fields("RucCliente")) = "20305948277" Then
                CodDistrito = Trim("" & traerCampo("Personas", "idDistrito", "idPersona", Trim("" & RsC.Fields("IdSucursal")), False))
            Else
                CodDistrito = Trim("" & traerCampo("Personas", "idDistrito", "idPersona", PIdCliente, False))
            End If
        
        End If
        
        RsC.Close: Set RsC = Nothing
        
    Else
        CodDistrito = Trim("" & traerCampo("Personas", "idDistrito", "idPersona", glsSucursal, False))
    End If
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Function TraeGlosaTransGratuita(StrMsgError As String, strTD As String, strSerie As String, strNumDoc As String, NItem As Long)
On Error GoTo Err
Dim CSqlC                                       As String
Dim RsC                                         As New ADODB.Recordset
Dim CGlosaTransGratuita                         As String

    If Val("" & traerCampo("DocVentas", "IndTransGratuitaMP", "IdSucursal", glsSucursal, True, "IdDocumento = '" & strTD & "' And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strNumDoc & "'")) = 1 Or Val("" & traerCampo("DocVentas", "IndTransGratuita", "IdSucursal", glsSucursal, True, "IdDocumento = '" & strTD & "' And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strNumDoc & "'")) = 1 Then
                            
        CSqlC = "Select A.IndTransGratuitaMP,C.Simbolo,B.TotalVVNeto,B.TotalIgvNeto,B.TotalPVNeto " & _
                "From DocVentas A " & _
                "Inner Join Monedas C " & _
                    "On A.IdMoneda = C.IdMoneda " & _
                "Inner Join DocVentasDet B " & _
                    "On A.IdEmpresa = B.IdEmpresa And A.IdSucursal = B.IdSucursal And A.IdDocumento = B.IdDocumento " & _
                    "And A.IdSerie = B.IdSerie And A.IdDocVentas = B.IdDocVentas " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdDocumento = '" & strTD & "' " & _
                "And A.IdSerie = '" & strSerie & "' And A.IdDocVentas = '" & strNumDoc & "' And B.Item = " & NItem & ""
        RsC.Open CSqlC, Cn, adOpenKeyset, adLockReadOnly
        If Not RsC.EOF Then
            
            CGlosaTransGratuita = " - " & IIf(Val("" & RsC.Fields("IndTransGratuitaMP")) = 1, "DEMOSTRACION CON FINES PROMOCIONALES VV ", "") & Trim("" & RsC.Fields("Simbolo"))
            
            CGlosaTransGratuita = CGlosaTransGratuita & Trim(Format(Val("" & RsC.Fields("TotalVVNeto")), "0.00")) & " + IGV " & Trim("" & RsC.Fields("Simbolo"))
            
            CGlosaTransGratuita = CGlosaTransGratuita & Trim(Format(Val("" & RsC.Fields("TotalIgvNeto")), "0.00")) & " = VT " & Trim("" & RsC.Fields("Simbolo"))
            
            CGlosaTransGratuita = CGlosaTransGratuita & Trim(Format(Val("" & RsC.Fields("TotalPVNeto")), "0.00")) & " - TRANSFERENCIA GRATUITA"
            
        End If
        
        RsC.Close: Set RsC = Nothing
        
    End If
    
    TraeGlosaTransGratuita = CGlosaTransGratuita
    
    Exit Function
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Function


Private Sub ImprimeDocVentasParte1(StrMsgError As String, strTD As String, strNumDoc As String, strSerie As String, StrTipoTicket As String, _
strRUCCliente As String, StrGlsCliente As String, StrTotalIGV As String, StrGlsVendedorCampo As String, StrDirecCliente As String, StrMotivoTraslado As String, _
StrCodigoCliente As String, IdTiendaCli As String, RsD As ADODB.Recordset, StrGlsMotivoTraslado As String)
On Error GoTo Err

    csql = "SELECT idtienda,GlsCliente,RUCCliente,TotalIGVVenta,idTipoTicket,GlsVendedorCampo,dirCliente,idMotivoTraslado,idPerCliente,idMotivoNCD " & _
           "FROM docventas " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'" & _
           "  AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    RsD.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not RsD.EOF Then
        StrTipoTicket = "" & RsD.Fields("idTipoTicket")
        strRUCCliente = "" & RsD.Fields("RUCCliente")
        StrGlsCliente = "" & RsD.Fields("GlsCliente")
        StrTotalIGV = CStr(Format("" & RsD.Fields("TotalIGVVenta"), "###,##0.00"))
        StrGlsVendedorCampo = "" & RsD.Fields("GlsVendedorCampo")
        StrDirecCliente = "" & RsD.Fields("dirCliente")
        StrMotivoTraslado = "" & RsD.Fields("idMotivoTraslado")
        StrCodigoCliente = "" & RsD.Fields("idPerCliente")
        IdTiendaCli = Trim("" & RsD.Fields("idtienda"))
    End If
    
    If Len(Trim(StrMotivoTraslado)) > 0 Then
        csql = "SELECT GlsMotivoTraslado FROM motivostraslados WHERE idMotivoTraslado = '" & StrMotivoTraslado & "'"
        If RsD.State = adStateOpen Then RsD.Close
        RsD.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not RsD.EOF Then
            StrGlsMotivoTraslado = Trim(RsD.Fields("GlsMotivoTraslado") & "")
        End If
    End If
    RsD.Close: Set RsD = Nothing
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub imprimeGuiaAbierta(strTD As String, strNumDoc As String, strSerie As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim wcont, wsum         As Integer
Dim wfila, wcolu        As Integer
Dim rst                 As New ADODB.Recordset
Dim rstObj              As New ADODB.Recordset
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strTipoFecDoc       As String
Dim strImprimeTicket    As String
Dim StrTipoTicket       As String
Dim strRUCCliente       As String
Dim StrGlsCliente       As String
Dim StrTotalIGV         As String
Dim StrGlsVendedorCampo As String
Dim StrDirecCliente     As String
Dim StrMotivoTraslado   As String
Dim StrCodigoCliente    As String
Dim cselect             As String
Dim nfontletra          As String
Dim numEntreLineasAdicional As Integer
Dim intRegistros            As Integer
Dim RsD                     As New ADODB.Recordset
Dim StrTexto                As String
Dim strTipoDato             As String
Dim intLong                 As Integer
Dim intX                    As Integer
Dim intY                    As Integer
Dim intDec                  As Integer
    
    '--- SELECCIONAMOS IMPRESORA
    indPrinter = False
    If strTD = "01" Then 'FACTURA
        For Each p In Printers
           If UCase(p.DeviceName) = "FACTURA" Then
              Set Printer = p
              indPrinter = True
              Exit For
           End If
        Next p
    ElseIf strTD = "03" Then 'BOLETA
            For Each p In Printers
                If UCase(p.DeviceName) = "BOLETA" Then
                    Set Printer = p
                    indPrinter = True
                    Exit For
                End If
            Next p
            
    ElseIf strTD = "86" Then 'GUIA
            For Each p In Printers
                If UCase(p.DeviceName) = "GUIA" Then
                    Set Printer = p
                    indPrinter = True
                    Exit For
                End If
            Next p
    End If
    
    If indPrinter = False Then
        For Each p In Printers
            If UCase(p.DeviceName) = "GENERAL" Then
                Set Printer = p
                indPrinter = True
                Exit For
            End If
        Next p
    End If
            
    If indPrinter = False Then
        For Each p In Printers
           If p.Port = "LPT1:" Then
              Set Printer = p
              Exit For
           End If
        Next p
    End If

    intScale = 6
    Printer.ScaleMode = intScale
    
    strTipoFecDoc = traerCampo("documentos", "TipoImpFecha", "idDocumento", strTD, False)
    strImprimeTicket = traerCampo("documentos", "indImprimeTicket", "idDocumento", strTD, False)
    numEntreLineasAdicional = Val("" & traerCampo("seriesdocumento", "espacioLineasImp", "idSerie", strSerie, True, " idsucursal = '" & glsSucursal & "' and iddocumento = '" & strTD & "'"))
    
    csql = "SELECT GlsCliente,RUCCliente,TotalIGVVenta,idTipoTicket,GlsVendedorCampo,dirCliente,idMotivoTraslado,idPerCliente " & _
           "FROM docventas " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'" & _
           "  AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
           
    RsD.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not RsD.EOF Then
        StrTipoTicket = "" & RsD.Fields("idTipoTicket")
        strRUCCliente = "" & RsD.Fields("RUCCliente")
        StrGlsCliente = "" & RsD.Fields("GlsCliente")
        StrTotalIGV = CStr(Format("" & RsD.Fields("TotalIGVVenta"), "###,##0.00"))
        StrGlsVendedorCampo = "" & RsD.Fields("GlsVendedorCampo")
        StrDirecCliente = "" & RsD.Fields("dirCliente")
        StrMotivoTraslado = "" & RsD.Fields("idMotivoTraslado")
        StrCodigoCliente = "" & RsD.Fields("idPerCliente")
    End If
    
    If Len(Trim(StrMotivoTraslado)) > 0 Then
        csql = "SELECT GlsMotivoTraslado FROM motivostraslados WHERE idMotivoTraslado = '" & StrMotivoTraslado & "'"
        If RsD.State = adStateOpen Then RsD.Close
        RsD.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not RsD.EOF Then
            StrGlsMotivoTraslado = Trim(RsD.Fields("GlsMotivoTraslado") & "")
        End If
    End If
    RsD.Close: Set RsD = Nothing
    Printer.FontSize = 8
    
    '--- ETIQUETAS
    '--- Traemos las etiquetas configuradas para el documento y la serie
    csql = "SELECT Etiqueta,impX,impY,idobjetiquetasdoc,indDirSucursal,indSerieEtiquetera,indSoloTicketFactura," & _
           "indRUCCliente,indRazonSocial,indIGVTotal,indHora,indDestinatarioGuia,indMotivoTrasladoGuia,indUsuario " & _
           "FROM objetiquetasdoc " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & strTD & "' AND idSerie = '" & _
           strSerie & "' AND tipoObj = 'C' ORDER BY IMPY,IMPX"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    Do While Not rst.EOF
        StrTexto = rst.Fields("Etiqueta") & ""
        
        If (rst.Fields("indDirSucursal") & "") = 1 Then StrTexto = StrTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)

        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                     StrTexto = StrTexto + strRUCCliente
                End If
            Else
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                     StrTexto = StrTexto + StrGlsCliente  '1
                End If
            Else
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                     StrTexto = StrTexto + StrTotalIGV
                End If
            Else
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indHora") & "") = 1 Then StrTexto = StrTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        '-------------------------------------------------------------------
        '------ BUSCA SI ES SUCURSAL, EXTRAE RAZON SOCIAL Y RUC DE LA EMPRESA
        If (rst.Fields("indDestinatarioGuia") & "") = 1 Then
            cselect = "SELECT idSucursal FROM sucursales WHERE idSucursal = '" & StrCodigoCliente & "' AND idEmpresa = '" & glsEmpresa & "'"
            If RsD.State = adStateOpen Then RsD.Close
            RsD.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
            If Not RsD.EOF Then
                RsD.Close: Set RsD = Nothing
                
                cselect = "SELECT GlsEmpresa, RUC FROM empresas WHERE idEmpresa = '" & glsEmpresa & "'"
                If RsD.State = adStateOpen Then RsD.Close
                RsD.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
                If Not RsD.EOF Then
                    StrTexto = StrTexto & Trim(RsD.Fields("GlsEmpresa") & "") & Space(5) & "RUC: " & Trim(RsD.Fields("RUC") & "")
                End If
                'rsd.Close: Set rsd = Nothing
            Else
                StrTexto = StrTexto + StrGlsCliente & Space(5) & "RUC/DNI: " & strRUCCliente
            End If
            RsD.Close: Set RsD = Nothing
        End If
        '-------------------------------------------------------------------
        If (rst.Fields("indMotivoTrasladoGuia") & "") = 1 Then
            StrTexto = StrTexto & StrGlsMotivoTraslado
        End If
        '-------------------------------------------------------------------
        ImprimeXY StrTexto, "T", Len(StrTexto & ""), rst.Fields("impY"), rst.Fields("impX"), 0, 0, StrMsgError
        If StrMsgError <> "" Then GoTo Err
        rst.MoveNext
    Loop
    rst.Close
    
    '--- CABECERA
    '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
    intRegistros = Val(traerCampo("objdocventas", "count(*)", "idDocumento", strTD, True, "tipoObj = 'C' and trim(GlsCampo) <> '' and indImprime = 1 and idSerie = '" & strSerie & "' "))
    If intRegistros = 0 Then
        csql = "SELECT GlsCampo " & _
               "FROM objdocventas " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'C' and trim(GlsCampo) <> '' and " & _
               "indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' " & _
               "ORDER BY IMPY,IMPX"
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    Else
        csql = "SELECT GlsCampo " & _
               "FROM objdocventas " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'C' and trim(GlsCampo) <> '' and " & _
               "indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "' " & _
               "ORDER BY IMPY,IMPX"
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    End If
    
    Do While Not rst.EOF
        strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
        rst.MoveNext
    Loop
    strCampos = left(strCampos, Len(strCampos) - 1)
    rst.Close
    
    '--- Traemos la data de lo campos seleccionados arriba
    csql = "SELECT " & strCampos & " FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        For i = 0 To rst.Fields.Count - 1
            '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
            If strTD = "86" And rst.Fields(i).Name = "idMotivoTraslado" Then
                csql = "SELECT 'X' AS valor,'T' AS tipoDato, 1 AS impLongitud, impX, impY,0 AS Decimales,0 as intNumFilas FROM impMotivosTraslados WHERE idEmpresa = '" & glsEmpresa & "' AND idSerie = '" & strSerie & "' AND idMotivoTraslado = '" & Trim(rst.Fields(i) & "") & "'"
            Else
                If intRegistros = 0 Then
                    csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales,intNumFilas FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' "
                Else
                    csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales,intNumFilas FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'"
                End If
            End If
            
            rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rstObj.EOF Then
                If rst.Fields(i).Name = "FecEmision" And strTipoFecDoc = "S" Then
                    'IMPRIME EL DIA
                    ImprimeXY Format(Day(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    
                    'IMPRIME EL MES
                    ImprimeXY strArregloMes(Val(Month(rstObj.Fields("valor")))) & "", "T", 10, rstObj.Fields("impY"), rstObj.Fields("impX") + 20, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    
                    'IMPRIME EL AÑO
                    ImprimeXY right(CStr(Year(rstObj.Fields("valor"))), 1) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 70, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                Else
                    ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                End If
            End If
            rstObj.Close
        Next
    End If
    rst.Close
    
    '--- DETALLE
    wcont = 1
    wsum = 0
    '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
    strCampos = ""
    If intRegistros = 0 Then
        csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' and (GlsCampo <> 'Cantidad' and GlsCampo <> 'Cantidad2') ORDER BY impY,impX"
    Else
        csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "' and (GlsCampo <> 'Cantidad' and GlsCampo <> 'Cantidad2') ORDER BY impY,impX"
    End If
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
        rst.MoveNext
    Loop
    strCampos = left(strCampos, Len(strCampos) - 1)
    rst.Close
    
    '--- Traemos la data de los campos seleccionados arriba
    csql = "SELECT " & strCampos & " FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    wsum = 0
    Do While Not rst.EOF
        For i = 0 To rst.Fields.Count - 1
            '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
            If intRegistros = 0 Then
                csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' "
            Else
                csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'"
            End If
            rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rstObj.EOF Then
                strTipoDato = rstObj.Fields("tipoDato")
                intLong = rstObj.Fields("impLongitud")
                intX = rstObj.Fields("impX")
                intY = rstObj.Fields("impY")
                intDec = Val("" & rstObj.Fields("Decimales"))
                    
                ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            rstObj.Close
        Next
        rst.MoveNext
        wsum = wsum + IIf(intScale = 6, 4, 1) + numEntreLineasAdicional
    Loop
    rst.Close
    
    If strImprimeTicket = "S" Then
        wsum = wsum + IIf(intScale = 6, 4, 1)
    Else
        wsum = 0
    End If
    
    strCampos = ""
    If intRegistros = 0 Then
        csql = "SELECT GlsCampo " & _
                "FROM objdocventas " & _
                "WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' and trim(GlsCampo) <> '' and " & _
                "indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' " & _
                " ORDER BY IMPY,IMPX"
    Else
        csql = "SELECT GlsCampo " & _
               "FROM objdocventas " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' and trim(GlsCampo) <> '' and " & _
               "indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "' " & _
               " ORDER BY IMPY,IMPX"
    End If
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
        rst.MoveNext
    Loop
    If Len(strCampos) > 0 Then
        strCampos = left(strCampos, Len(strCampos) - 1)
    End If
    rst.Close
    
    '--- Traemos la data de lo campos seleccionados arriba
    If Len(strCampos) > 0 Then
        csql = "SELECT " & strCampos & ",IdMoneda FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        
        If Not rst.EOF Then
            For i = 0 To rst.Fields.Count - 1
                '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
                '--- Imprime motivos de guias y nota credito
                If (strTD = "86" And rst.Fields(i).Name = "idMotivoTraslado") Or (strTD = "07" And rst.Fields(i).Name = "idMotivoNCD") Then
                    csql = "SELECT 'X' AS valor,'T' AS tipoDato, 1 AS impLongitud, impX, impY,0 AS Decimales FROM impMotivosTraslados WHERE idEmpresa = '" & glsEmpresa & "' AND idSerie = '" & strSerie & "' AND idMotivoTraslado = '" & Trim(rst.Fields(i) & "") & "'"
                Else
                    If intRegistros = 0 Then
                        csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999'"
                    Else
                        csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'"
                    End If
                End If
                
                rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rstObj.EOF Then
                    If strImprimeTicket = "S" Then
                        If rst.Fields(i).Name = "TotalPrecioVenta" Then
                            If rst.Fields("IdMoneda") = "PEN" Then
                                ImprimeXY "Total :  S/ ", "T", 20, intY + wsum, 1, 0, 0, StrMsgError
                            Else
                                ImprimeXY "Total : US$. ", "T", 20, intY + wsum, 1, 0, 0, StrMsgError
                            End If
                        End If
                    End If
                    
                    If strImprimeTicket = "S" Then
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), intY + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        wsum = wsum + IIf(intScale = 6, 4, 1)
                    Else
                        If strTD = "86" Or strTD = "07" Then
                            ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), 0, 0, StrMsgError
                        Else
                            ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        End If
                    End If
                    If StrMsgError <> "" Then GoTo Err
                End If
                rstObj.Close
            Next
        End If
        rst.Close
    End If

    '------------------------------------------------------------------------------------------------
    '--- IMPRIME ETIQUETAS FINAL
    If strImprimeTicket = "S" Then
        wsum = wsum + IIf(intScale = 6, 4, 1)
    Else
        wsum = 0
    End If
    
    csql = "SELECT Etiqueta,impX,impY,indDirSucursal,indSerieEtiquetera,indSoloTicketFactura,indRUCCliente,indRazonSocial,indIGVTotal,indHora,indVendedor,indDirecCliente,indUsuario " & _
           "FROM objetiquetasdoc " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & strTD & "' AND idSerie = '" & _
           strSerie & "'  AND tipoObj = 'T' ORDER BY impY,impX"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    Do While Not rst.EOF
         StrTexto = ""
        If (rst.Fields("indRUCCliente") & "") = 0 And (rst.Fields("indRazonSocial") & "") = 0 And (rst.Fields("indIGVTotal") & "") = 0 And (rst.Fields("indHora") & "") = 0 And (rst.Fields("indVendedor") & "") = 0 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = rst.Fields("Etiqueta") & ""
                End If
            End If
        End If
                        
        If (rst.Fields("indDirSucursal") & "") = 1 Then StrTexto = StrTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        '08001 FACTURA
        '08002 BOLETA
        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + strRUCCliente
                End If
            Else
                StrTexto = rst.Fields("Etiqueta") & ""
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrGlsCliente  '2
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrGlsCliente  '2
                End If
            End If
        End If
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                     StrTexto = StrTexto + StrTotalIGV
                End If
            Else
                StrTexto = rst.Fields("Etiqueta") & ""
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indHora") & "") = 1 Then
            StrTexto = rst.Fields("Etiqueta") & ""
            StrTexto = StrTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        End If
                           
        If (rst.Fields("indVendedor") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrGlsVendedorCampo
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrGlsVendedorCampo
                End If
            End If
        End If
        
        If (rst.Fields("indDirecCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrDirecCliente
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrDirecCliente
                End If
            End If
        End If
        
        If Len(Trim(StrTexto)) > 0 Then
            ImprimeXY StrTexto, "T", Len(StrTexto), intY + wsum, rst.Fields("impX"), 0, 0, StrMsgError
            wsum = wsum + IIf(intScale = 6, 4, 1)
        Else
        
        End If
        If StrMsgError <> "" Then GoTo Err
        
        rst.MoveNext
    Loop
    rst.Close
    '------------------------------------------------------------------------------------------------
    Set rst = Nothing
    Set rstObj = Nothing
                                   
    Printer.Print Chr$(149)
    Printer.Print ""
    Printer.Print ""
    Printer.EndDoc

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Printer.KillDoc
End Sub

Private Sub ImprimeXY(varData As Variant, strTipoDato As String, intTamanoCampo As Integer, intFila As Integer, intColu As Integer, intDecimales As Integer, intFilas As Integer, ByRef StrMsgError As String)
On Error GoTo Err
Dim i As Integer
Dim strDec  As String
Dim indFinWhile As Boolean
Dim intFilaImp As Integer
Dim intIndiceInicio As Integer
    
    Select Case strTipoDato
        Case "T"   'texto
             If (intFilas = 0 Or intFilas = 1) Or Len(varData) <= intTamanoCampo Then
                Printer.CurrentY = intFila
                Printer.CurrentX = intColu
                Printer.Print left(varData, intTamanoCampo)
             Else
                indFinWhile = True
                intFilaImp = 0
                intIndiceInicio = 1
                
                Do While (indFinWhile = True)
                    If intFilaImp < intFila Then
                        intFilaImp = intFilaImp + 1
                        
                        Printer.CurrentY = intFila
                        Printer.CurrentX = intColu
                        Printer.Print Mid(varData, intIndiceInicio, intTamanoCampo)
                        
                        intFila = intFila + 5
                        
                        intIndiceInicio = intIndiceInicio + intTamanoCampo
                    Else
                        indFinWhile = False
                    End If
                Loop
             End If
        Case "F"   'Fecha
             Printer.CurrentY = intFila
             Printer.CurrentX = intColu
             Printer.Print left(Format(varData, "dd/mm/yyyy"), intTamanoCampo)
        Case "H"   'Hora
             Printer.CurrentY = intFila
             Printer.CurrentX = intColu
             Printer.Print left(Format(varData, "hh:MM"), intTamanoCampo)
        Case "Y"   'Fecha y Hora
             Printer.CurrentY = intFila
             Printer.CurrentX = intColu
             Printer.Print left(Format(varData, "dd/mm/yyyy hh:MM"), intTamanoCampo)
        Case "N"     'numerico
            Printer.CurrentY = intFila
            Printer.CurrentX = intColu
                    
            '--- Asigna la cantidad de decimales
            For i = 1 To intDecimales
                strDec = strDec & "0"
            Next
            
            If intDecimales > 0 Then
                Printer.Print right((Space(intTamanoCampo) & Format(varData, "#,###,##0." & strDec)), intTamanoCampo)
            Else
                Printer.Print right((Space(intTamanoCampo) & Format(varData, "#,###,##0" & strDec)), intTamanoCampo)
            End If
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub imprimeReciboCaja(strNumMovCajaDet As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst             As New ADODB.Recordset
Dim rstObj          As New ADODB.Recordset
Dim p               As Object
Dim indPrinter      As Boolean
Dim intScale        As Integer

    '--- SELECCIONAMOS IMPRESORA
    indPrinter = False
    For Each p In Printers
       If UCase(p.DeviceName) = "RECIBO" Then
          Set Printer = p
          indPrinter = True
          Exit For
       End If
    Next p
    
    If indPrinter = False Then
        For Each p In Printers
            If UCase(p.DeviceName) = "GENERAL" Then
                Set Printer = p
                indPrinter = True
                Exit For
            End If
        Next p
    End If
            
    If indPrinter = False Then
        For Each p In Printers
           If p.Port = "LPT1:" Then
              Set Printer = p
              Exit For
           End If
        Next p
    End If

    intScale = 6
    Printer.ScaleMode = intScale
    Printer.FontName = "Roman 12cpi"
    Printer.FontBold = False
    Printer.FontSize = 8
    
    '--- CABECERA
    '--- Traemos la data a imprimir
    csql = "SELECT c.idMovCajaDet,c.FecRegistro,c.ValMonto,c.GlsObs," & _
                  "c.idTipoMovCaja,t.GlsTipoMovCaja," & _
                  "c.idMoneda, m.GlsMoneda,m.Simbolo " & _
           "FROM movcajasdet c,tiposmovcaja t,monedas m " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' " & _
             "AND c.idTipoMovCaja = t.idTipoMovCaja " & _
             "AND c.idMoneda = m.idMoneda " & _
             "AND c.idMovCajaDet = '" & strNumMovCajaDet & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        For i = 0 To rst.Fields.Count - 1
            '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
            csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales " & _
                   "FROM objimprecibos " & _
                   "WHERE idEmpresa = '" & glsEmpresa & "' " & _
                     "AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 "
            rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rstObj.EOF Then
                    ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
            End If
            rstObj.Close
        Next
    End If
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Set rstObj = Nothing
                           
    Printer.EndDoc
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Printer.KillDoc
End Sub

Public Sub ImprimeCodigoBarra(ByVal indTipo As Integer, ByVal codproducto As String, ByVal numVale As String, ByRef StrMsgError As String, Optional ByVal dblCantidad As Double = 0)
'On Error GoTo Err
'Dim objPrinter As New PrinterAPI.clsPrinter
'Dim rsp As New ADODB.Recordset
'Dim StrCodBarra As String
'Dim BlnFoundPrinter As Boolean
'Dim BlnFoundData As Boolean
'Dim intPar As Long
'Dim i As Integer
'Dim intParTotal As Long
'Dim indParTotal As Boolean
'Dim strPrecio As String
'Dim strTalla As String
'Dim intCantidad As Integer
'
'    If (objPrinter.SetPrinter("Generica / Solo Texto") = False) Then
'        If (objPrinter.SetPrinter("Generic / Text Only") = False) Then
'            MsgBox "No se Encuentra instalada la Impresora " & NombreImpresora_sp & "o " & NombreImpresora_us, vbInformation
'            Exit Sub
'        End If
'    End If
'
'    StrGlsEmpresa = traerCampo("empresas", "GlsEmpresa", "idEmpresa", glsEmpresa, False)
'    If indTipo = 0 Then
'        csql = "SELECT v.idProducto,p.GlsProducto,t.GlsTallaPeso,l.PVUnit,v.Cantidad " & _
'               "FROM valesdet v,productos p, tallapeso t, preciosventa l " & _
'               "WHERE v.idEmpresa = p.idEmpresa " & _
'                 "AND v.idProducto = p.idProducto " & _
'                 "AND p.idEmpresa = t.idEmpresa " & _
'                 "AND p.idTallaPeso = t.idTallaPeso " & _
'                 "AND p.idEmpresa = l.idEmpresa " & _
'                 "AND p.idProducto = l.idProducto " & _
'                 "AND p.idUMCompra = l.idUM " & _
'                 "AND v.idValesCab = '" & numVale & "' AND p.idProducto = '" & codproducto & "' AND l.idLista = '" & glsListaVentas & "'"
'        intCantidadTotal = 1
'
'    ElseIf indTipo = 1 Then
'        csql = "SELECT v.idProducto,p.GlsProducto,t.GlsTallaPeso,l.PVUnit,v.Cantidad " & _
'               "FROM valesdet v,productos p, tallapeso t, preciosventa l " & _
'               "WHERE v.idEmpresa = p.idEmpresa " & _
'                 "AND v.idProducto = p.idProducto " & _
'                 "AND p.idEmpresa = t.idEmpresa " & _
'                 "AND p.idTallaPeso = t.idTallaPeso " & _
'                 "AND p.idEmpresa = l.idEmpresa " & _
'                 "AND p.idProducto = l.idProducto " & _
'                 "AND p.idUMCompra = l.idUM " & _
'                 "AND v.idValesCab = '" & numVale & "' AND l.idLista = '" & glsListaVentas & "'"
'        intCantidadTotal = Val("" & traerCampo("valesdet", "SUM(Cantidad)", "idSucursal", glsSucursal, True, " idValesCab = '" & numVale & "'"))
'
'    ElseIf indTipo = 2 Then
''        csql = "SELECT p.idProducto,p.GlsProducto,t.GlsTallaPeso,l.PVUnit," & CStr(dblCantidad) & " AS Cantidad " & _
''               "FROM productos p, tallapeso t, preciosventa l " & _
''               "WHERE p.idEmpresa = t.idEmpresa " & _
''                 "AND p.idTallaPeso = t.idTallaPeso " & _
''                 "AND p.idEmpresa = l.idEmpresa " & _
''                 "AND p.idProducto = l.idProducto " & _
''                 "AND p.idUMCompra = l.idUM " & _
''                 "AND p.idProducto = '" & codproducto & "' AND l.idLista = '" & glsListaVentas & "'"
'
'        csql = "SELECT p.idProducto,p.GlsProducto,'' AS GlsTallaPeso, 0 AS PVUnit," & CStr(dblCantidad) & " AS Cantidad " & _
'               "FROM productos p " & _
'               "WHERE p.idProducto = '" & codproducto & "' AND p.idEmpresa = '" & glsEmpresa & "'"
'
'        intCantidadTotal = 1
'
'    End If
'    rsp.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
'
'    intPar = 0
'    If indTipo = 0 Or indTipo = 2 Then
'        If Not rsp.EOF Then
'            intParTotal = intCantidadTotal * Val("" & rsp.Fields("Cantidad"))
'        Else
'            StrMsgError = "El producto no tiene precio"
'            GoTo Err
'        End If
'    Else
'        intParTotal = intCantidadTotal
'    End If
'
'    indParTotal = True
'    If intParTotal Mod 2 Then
'        indParTotal = False
'    End If
'
'    Do While (Not rsp.EOF)
'        strPrecio = Format(rsp.Fields("PVUnit").Value, "##,##0.00")
'        strTalla = Trim$(rsp.Fields("GlsTallaPeso").Value)
'        intCantidad = Val("" & rsp.Fields("Cantidad"))
'
'        For i = 1 To intCantidad
'            intPar = intPar + 1
'            objPrinter.PrintDataLn Chr$(2) & "L"
'            objPrinter.PrintDataLn "A2"
'            objPrinter.PrintDataLn "D11"
'            objPrinter.PrintDataLn "z"
'            objPrinter.PrintDataLn "PN"
'            objPrinter.PrintDataLn "H10"
'
'            StrCodBarra = Trim$(rsp.Fields("idProducto").Value)
'
'            If intPar Mod 2 Then
'                objPrinter.PrintDataLn "191100300610140" & strPrecio 'PRECIO -30
'                objPrinter.PrintDataLn "191100100280010" & strTalla 'TALLA
'
'                objPrinter.PrintDataLn "191100100610070" & StrGlsEmpresa
'                objPrinter.PrintDataLn "191100100500010" & left(Trim$(rsp.Fields("GlsProducto").Value), 38)
'                objPrinter.PrintDataLn "191100300280133" & StrCodBarra
'                objPrinter.PrintDataLn "1e2201600010016B" & StrCodBarra
'
'                objPrinter.PrintDataLn "^01"     ' Numero de Copias
'                objPrinter.PrintDataLn "Q0001"   ' Numero de Etiquetas
'
'            Else
'                objPrinter.PrintDataLn "191100300610350" & strPrecio 'PRECIO
'                objPrinter.PrintDataLn "191100100280220" & strTalla 'TALLA
'
'                objPrinter.PrintDataLn "191100100610280" & StrGlsEmpresa
'                objPrinter.PrintDataLn "191100100500220" & left(Trim$(rsp.Fields("GlsProducto").Value), 38)
'                objPrinter.PrintDataLn "191100300280342" & StrCodBarra
'                objPrinter.PrintDataLn "1e2201600010228B" & StrCodBarra
'
'                objPrinter.PrintDataLn "^01"     ' Numero de Copias
'                objPrinter.PrintDataLn "Q0001"   ' Numero de Etiquetas
'                objPrinter.PrintDataLn "E"       ' Enviar la Impresion
'
'            End If
'
'            If indParTotal = False And intPar = intParTotal Then
'                objPrinter.PrintDataLn "E"       ' Enviar la Impresion
'            End If
'            BlnFoundData = True
'        Next
'        rsp.MoveNext
'    Loop
'    rsp.Close: Set rsp = Nothing
'
'    If (BlnFoundData) Then
'        Printer.EndDoc
'    Else
'        StrMsgError = "No se Realizó ninguna Impresión, No hay Datos"
'        GoTo Err
'    End If
'
'    Exit Sub
'
'Err:
'    If rsp.State = 1 Then rsp.Close: Set rsp = Nothing
'    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub imprimirVale(ByVal strNumVale As String)
On Error GoTo Err
Dim rsReporte       As New ADODB.Recordset
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim StrMsgError As String
                    
    Screen.MousePointer = 11
    gStrRutaRpts = App.Path + "\Reportes\"
    Set reporte = aplicacion.OpenReport(gStrRutaRpts & "rptImpVale.rpt")
    
    DoEvents
    
    Set rsReporte = DataProcedimiento("spu_ImpVales", StrMsgError, glsEmpresa, glsSucursal, strNumVale)
    If StrMsgError <> "" Then GoTo Err

    If Not rsReporte.EOF And Not rsReporte.BOF Then
        reporte.database.SetDataSource rsReporte, 3

        vistaPrevia.CRViewer91.ReportSource = reporte
        vistaPrevia.Caption = "Vale"
        vistaPrevia.CRViewer91.ViewReport
        vistaPrevia.CRViewer91.DisplayGroupTree = False
        Screen.MousePointer = 0
        vistaPrevia.WindowState = 2

        vistaPrevia.Show
    Else
        Screen.MousePointer = 0
        MsgBox "No existen Registros  Seleccionados", vbInformation, App.Title
    End If
    Screen.MousePointer = 0
    Set rsReporte = Nothing
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
    
    Exit Sub
    
Err:
    Screen.MousePointer = 0
    If StrMsgError = "" Then StrMsgError = Err.Description
    Set rsReporte = Nothing
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Function PropperWrap(ByVal sCadena As String, _
                            ByVal nCaracteres As Long, _
                            Optional ByVal DesdeDonde As ePropperWrapConstants = pwLeft) As String
    'Devuelve la cadena que habría que imprimir para mostrar los
    'caracteres indicados, sin cortar una palabra.
    'Esto es para los casos en los que se quiera usar:
    'Left$(sCadena,nCaracteres) o Mid$/Right$(sCadena,nCaracteres)
    'pero sin cortar una palabra
Dim i As Long
Dim sChar As String
    
    sSeparadores = Chr$(13) & " ,.;:_" & Chr$(34)
    i = InStr(sCadena, vbCrLf)
    If i > 0 And i < nCaracteres Then
        sCadena = left$(sCadena, i + 1)
    ElseIf nCaracteres > Len(sCadena) Then
        i = InStr(sCadena, vbCrLf)
        If i Then
            sCadena = left$(sCadena, i - 1)
        End If
    Else
        For i = nCaracteres To 1 Step -1
            If InStr(sSeparadores, Mid$(sCadena, i, 1)) Then
                '--- Si se especifica desde la izquierda
                If DesdeDonde = pwLeft Then
                    sCadena = left$(sCadena, i)
                Else
                '--- Lo mismo da desde el centro que desde la derecha
                    sCadena = Mid$(sCadena, i + 1)
                End If
                Exit For
            End If
        Next
    End If
    PropperWrap = sCadena
    
End Function

Public Function PropperRight(ByVal sCadena As String, ByVal nCaracteres As Long) As String
    
    PropperRight = PropperWrap(sCadena, nCaracteres, pwRight)

End Function

Public Function PropperMid(ByVal sCadena As String, ByVal nCaracteres As Long, Optional ByVal RestoNoUsado As Long) As String
    
    PropperMid = PropperWrap(sCadena, nCaracteres, pwMid)

End Function

Public Function PropperLeft(ByVal sCadena As String, ByVal nCaracteres As Long) As String
    
    PropperLeft = PropperWrap(sCadena, nCaracteres, pwLeft)

End Function

Private Sub Class_Initialize()
    
    sSeparadores = Chr$(13) & cSeparadores & Chr$(34)
    
End Sub

Public Property Get Separadores() As String
    
    Separadores = sSeparadores

End Property

Public Property Let Separadores(ByVal NewSeparadores As String)
    
    sSeparadores = NewSeparadores

End Property

Public Function LoopPropperWrap(Optional ByVal sCadena As String, _
                    Optional ByVal nCaracteres As Long = 60&, _
                    Optional ByVal DesdeDonde As ePropperWrapConstants = pwLeft) As String
    ' Repite la justificación hasta que la cadena esté vacia        (20/Ago/01)
    ' Devolviendo cada vez el número de caracteres indicados
    Static sCadenaCopia As String
    Static nCaracteresCopia As Long
    Static DesdeDondeCopia As ePropperWrapConstants

Dim s As String
    
    ' Si la cadena es una cadena vacía, es que se continua "partiendo"
    ' sino, es la primera llamada
    If Len(sCadena) Then
        sCadenaCopia = sCadena
        nCaracteresCopia = nCaracteres
        DesdeDondeCopia = DesdeDonde
    Else
        ' Asignar los valores que había antes
        sCadena = sCadenaCopia
        nCaracteres = nCaracteresCopia
        DesdeDonde = DesdeDondeCopia
    End If
    
    s = PropperWrap(sCadena, nCaracteres, DesdeDonde)
    sCadenaCopia = Mid$(sCadena, Len(s) + 1)
    If right$(s, 2) = vbCrLf Then
        s = left$(s, Len(s) - 2)
    End If
    LoopPropperWrap = s

End Function

Public Sub imprimeDocVentas_2(strTD As String, strNumDoc As String, strSerie As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim wcont, wsum         As Integer
Dim wfila, wcolu        As Integer
Dim rst                 As New ADODB.Recordset
Dim rstObj              As New ADODB.Recordset
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strTipoFecDoc       As String
Dim strImprimeTicket    As String
Dim StrTipoTicket       As String
Dim strRUCCliente       As String
Dim StrGlsCliente       As String
Dim StrTotalIGV         As String
Dim StrGlsVendedorCampo As String
Dim StrDirecCliente     As String
Dim StrMotivoTraslado   As String
Dim StrCodigoCliente    As String
Dim cselect             As String
Dim nfontletra          As String
Dim numEntreLineasAdicional As Integer
Dim intRegistros As Integer
Dim indventasterceros As String
Dim indterceros As Boolean
Dim nSumFilasEital      As String
Dim NTamanoLetra        As String
Dim RsD                 As New ADODB.Recordset
Dim StrTexto            As String
Dim nfiladetalle        As Integer, nvarfila As Integer, nvarfilatotal As Integer
Dim strTipoDato         As String
Dim intLong             As Integer
Dim intX                As Integer
Dim intY                As Integer
Dim intDec              As Integer
Dim rucEmpresa          As String
Dim iddocumentoRef      As String
Dim Long_total As Integer, Long_Acumu As Integer, contadorImp As Integer
Dim strGlosaDscto As String
    
    If Not (rucEmpresa = "20266578572" Or rucEmpresa = "20552185286" Or rucEmpresa = "20552184981") Then
        If strTD = "12" Then
            Printer.Print Chr$(149)
        End If
        Printer.EndDoc
        Exit Sub
    End If
    
    nSumFilasEital = 148
    indPrinter = True
    
    intScale = 6
    Printer.ScaleMode = intScale
    NTamanoLetra = Val(traerCampo("empresas", "NTamanoLetra", "idEmpresa", glsEmpresa, False) & "")
    Printer.FontSize = IIf(NTamanoLetra > 0, NTamanoLetra, Printer.FontSize)
    
    strTipoFecDoc = traerCampo("documentos", "TipoImpFecha", "idDocumento", strTD, False)
    strImprimeTicket = traerCampo("documentos", "indImprimeTicket", "idDocumento", strTD, False)
    numEntreLineasAdicional = Val("" & traerCampo("seriesdocumento", "espacioLineasImp", "idSerie", strSerie, True, " idsucursal = '" & glsSucursal & "' and iddocumento = '" & strTD & "'"))
    
    csql = "SELECT GlsCliente,RUCCliente,TotalIGVVenta,idTipoTicket,GlsVendedorCampo,dirCliente,idMotivoTraslado,idPerCliente " & _
           "FROM docventas " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'" & _
           "  AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
           
    RsD.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not RsD.EOF Then
        StrTipoTicket = "" & RsD.Fields("idTipoTicket")
        strRUCCliente = "" & RsD.Fields("RUCCliente")
        StrGlsCliente = "" & RsD.Fields("GlsCliente")
        StrTotalIGV = CStr(Format("" & RsD.Fields("TotalIGVVenta"), "###,##0.00"))
        StrGlsVendedorCampo = "" & RsD.Fields("GlsVendedorCampo")
        StrDirecCliente = "" & RsD.Fields("dirCliente")
        StrMotivoTraslado = "" & RsD.Fields("idMotivoTraslado")
        StrCodigoCliente = "" & RsD.Fields("idPerCliente")
    End If
    
    If Len(Trim(StrMotivoTraslado)) > 0 Then
        csql = "SELECT GlsMotivoTraslado FROM motivostraslados WHERE idMotivoTraslado = '" & StrMotivoTraslado & "'"
        If RsD.State = adStateOpen Then RsD.Close
        RsD.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not RsD.EOF Then
            StrGlsMotivoTraslado = Trim(RsD.Fields("GlsMotivoTraslado") & "")
        End If
    End If
    RsD.Close: Set RsD = Nothing
    
    If Len(Trim((traerCampo("sucursales", "TipoLetra", "idSucursal", glsSucursal, True)))) > 0 Then
        Printer.FontName = traerCampo("sucursales", "TipoLetra", "idSucursal", glsSucursal, True)
    End If
    
    '--- ETIQUETAS
    '--- Traemos las etiquetas configuradas para el documento y la serie
    csql = "SELECT Etiqueta,impX,impY,idobjetiquetasdoc,indDirSucursal,indSerieEtiquetera,indSoloTicketFactura," & _
           "indRUCCliente,indRazonSocial,indIGVTotal,indHora,indDestinatarioGuia,indMotivoTrasladoGuia,indUsuario " & _
           "FROM objetiquetasdoc " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & strTD & "' AND idSerie = '" & _
           strSerie & "' AND tipoObj = 'C' ORDER BY IMPY,IMPX"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    Do While Not rst.EOF
        StrTexto = rst.Fields("Etiqueta") & ""
        If (rst.Fields("indDirSucursal") & "") = 1 Then StrTexto = StrTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        '08001 FACTURA
        '08002 BOLETA
        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                     StrTexto = StrTexto + strRUCCliente
                End If
            Else
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                     StrTexto = StrTexto + StrGlsCliente  '1
                End If
            Else
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                     StrTexto = StrTexto + StrTotalIGV
                End If
            Else
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indHora") & "") = 1 Then StrTexto = StrTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        
        '-------------------------------------------------------------------
        '------ BUSCA SI ES SUCURSAL, EXTRAE RAZON SOCIAL Y RUC DE LA EMPRESA
        If (rst.Fields("indDestinatarioGuia") & "") = 1 Then
            cselect = "SELECT idSucursal FROM sucursales WHERE idSucursal = '" & StrCodigoCliente & "' AND idEmpresa = '" & glsEmpresa & "'"
            If RsD.State = adStateOpen Then RsD.Close
            RsD.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
            If Not RsD.EOF Then
                RsD.Close: Set RsD = Nothing
                cselect = "SELECT GlsEmpresa, RUC FROM empresas WHERE idEmpresa = '" & glsEmpresa & "'"
                If RsD.State = adStateOpen Then RsD.Close
                RsD.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
                If Not RsD.EOF Then
                    StrTexto = StrTexto & Trim(RsD.Fields("GlsEmpresa") & "") & Space(5) & "RUC: " & Trim(RsD.Fields("RUC") & "")
                End If
            Else
                StrTexto = StrTexto + StrGlsCliente & Space(5) & "RUC/DNI: " & strRUCCliente
            End If
            RsD.Close: Set RsD = Nothing
        End If
        '-------------------------------------------------------------------
        If (rst.Fields("indMotivoTrasladoGuia") & "") = 1 Then
            StrTexto = StrTexto & StrGlsMotivoTraslado
        End If
        '-------------------------------------------------------------------
        ImprimeXY StrTexto, "T", Len(StrTexto & ""), rst.Fields("impY") + nSumFilasEital, rst.Fields("impX"), 0, 0, StrMsgError
        If StrMsgError <> "" Then GoTo Err
        rst.MoveNext
    Loop
    rst.Close
    
    '--- CABECERA
    '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
    intRegistros = Val(traerCampo("objdocventas", "count(*)", "idDocumento", strTD, True, "tipoObj = 'C' and trim(GlsCampo) <> '' and indImprime = 1 and idSerie = '" & strSerie & "' "))
    If intRegistros = 0 Then
        csql = "SELECT GlsCampo " & _
               "FROM objdocventas " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'C' and trim(GlsCampo) <> '' and " & _
               "indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' " & _
               "ORDER BY IMPY,IMPX"
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    Else
        csql = "SELECT GlsCampo " & _
               "FROM objdocventas " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'C' and trim(GlsCampo) <> '' and " & _
               "indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "' " & _
               "ORDER BY IMPY,IMPX"
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    End If
    
    Do While Not rst.EOF
        strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
        rst.MoveNext
    Loop
    strCampos = left(strCampos, Len(strCampos) - 1)
    rst.Close
    
    '---- AGREGADO EL 06/05/10 VENTAS A TERCEROS DEPENDIENDO SI EL CAMPO IND VENTAS TERCEROS ESTA CON 1
    indterceros = False
    If strTD = "86" Then
        indventasterceros = traerCampo("clientes", "indventasterceros", "idcliente", StrCodigoCliente, True)
        If indventasterceros = "1" Then
            indterceros = True
        Else
            indterceros = False
        End If
    End If
    
    '--- Traemos la data de lo campos seleccionados arriba
    csql = "SELECT " & strCampos & " FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        For i = 0 To rst.Fields.Count - 1
            '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
            If strTD = "86" And rst.Fields(i).Name = "idMotivoTraslado" Then
                csql = "SELECT 'X' AS valor,'T' AS tipoDato, 1 AS impLongitud, impX, impY,0 AS Decimales,0 as intNumFilas FROM impMotivosTraslados WHERE idEmpresa = '" & glsEmpresa & "' AND idSerie = '" & strSerie & "' AND idMotivoTraslado = '" & Trim(rst.Fields(i) & "") & "'"
            Else
                If intRegistros = 0 Then
                    csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales,intNumFilas FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' "
                Else
                    csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales,intNumFilas FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'"
                End If
            End If
                            
            rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rstObj.EOF Then
                If rst.Fields(i).Name = "FecEmision" And strTipoFecDoc = "S" Then
                    'IMPRIME EL DIA
                    ImprimeXY Format(Day(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY") + nSumFilasEital, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    
                    'IMPRIME EL MES
                    ImprimeXY strArregloMes(Val(Month(rstObj.Fields("valor")))) & "", "T", 10, rstObj.Fields("impY") + nSumFilasEital, rstObj.Fields("impX") + 20, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    
                    'IMPRIME EL AÑO
                    ImprimeXY right((Year(rstObj.Fields("valor"))), 2) & "", "T", 4, rstObj.Fields("impY") + nSumFilasEital, rstObj.Fields("impX") + 60, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    
                ElseIf rst.Fields(i).Name = "ObsDocVentas" Then
                        xfila = Val(rstObj.Fields("impY"))
                        ncol = Val(rstObj.Fields("impX"))
                
                        Linea_l = rstObj.Fields("valor") & ""
                        margenIzq = 0
                        If margenIzq < 0 Or margenIzq > 40 Then
                            margenIzq = 0
                        End If
                        
                        longTotal = Val("" & rstObj.Fields("impLongitud"))
                        If longTotal < margenIzq Or longTotal > 136 Then
                            longTotal = Val("" & rstObj.Fields("impLongitud"))
                        End If
                        
                        sCabecera = Linea_l
                        s = LoopPropperWrap(sCabecera, longTotal, pwLeft)
                        sCabecera = ""
                        
                        ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila + nSumFilasEital, ncol, Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                        
                        xfila = xfila + 4
                        wsum = wsum + 4
                        nvarfila = nvarfila + 4
                        filas_detalle = 0
                        
                        If Len(Linea_l) <= Len(s) Then
                        Else
                            Do While Len(s) > 0
                                '--- Añadirle el margen izquierdo
                                sCabecera = sCabecera & Space$(margenIzq) & s & vbCrLf
                                s = LoopPropperWrap()
                                xMemo = s
                                ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila + nSumFilasEital, ncol, Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                                xfila = xfila + 4
                                wsum = wsum + 4
                                nvarfila = nvarfila + 4
                                filas_detalle = filas_detalle + 4
                            Loop
                        End If
                
                Else
                    If indterceros = True Then
                        If rst.Fields(i).Name = "GlsCliente" Or rst.Fields(i).Name = "RUCCliente" Or rst.Fields(i).Name = "llegada" Then
                        Else
                            ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + nSumFilasEital, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                        End If
                    Else
                        If rst.Fields(i).Name = "vtercerosCliente" Or rst.Fields(i).Name = "vterceroRuc" Or rst.Fields(i).Name = "vtercerosdireccion" Then
                        Else
                             ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + nSumFilasEital, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                             If StrMsgError <> "" Then GoTo Err
                        End If
                    End If
                End If
            End If
            rstObj.Close
        Next
    End If
    rst.Close
    
    '--- DETALLE
    wcont = 1
    wsum = 0
    
    '--- SI LA EMPRESA ES SINCHI O PIC HACE UNA IMPRESION PARTICULAR
    rucEmpresa = traerCampo("empresas", "ruc", "idEmpresa", glsEmpresa, False)
    If rucEmpresa = "20119041148" Or rucEmpresa = "20119040842" Or rucEmpresa = "20530611681" Then      ''' SINCHI      PIC     ACUICULTURA
        If glsSucursal <> "08090001" And (strTD = "01" Or strTD = "07") Then
            '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
            strCampos = ""
            If intRegistros = 0 Then
                csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999'  ORDER BY impY,impX"
            Else
                csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'  ORDER BY impY,impX"
            End If
            rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            Do While Not rst.EOF
                strCampos = strCampos & "d." & "" & rst.Fields("GlsCampo") & ","
                rst.MoveNext
            Loop
            strCampos = left(strCampos, Len(strCampos) - 1)
            rst.Close
            
            '--- Traemos la data de los campos seleccionados arriba
            csql = "SELECT " & strCampos & " FROM docventasdet d, productos p " & _
                   "WHERE d.idProducto = p.idProducto AND d.idEmpresa = p.idEmpresa " & _
                   "AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' " & _
                   "AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "' AND p.idNivel <> '10010062' "
            rst.Open csql, Cn, adOpenStatic, adLockReadOnly
            wsum = 0
            Do While Not rst.EOF
                For i = 0 To rst.Fields.Count - 1
                    '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
                    If intRegistros = 0 Then
                        csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' "
                    Else
                        csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'"
                    End If
                    rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                    
                    If rst.Fields(i).Name = "GlsProducto" Then
                        Printer.FontSize = 9
                    Else
                        Printer.FontSize = 9
                    End If
                    
                    If Not rstObj.EOF Then
                        strTipoDato = rstObj.Fields("tipoDato")
                        intLong = rstObj.Fields("impLongitud")
                        intX = rstObj.Fields("impX")
                        intY = rstObj.Fields("impY")
                        intDec = Val("" & rstObj.Fields("Decimales"))
                        ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                    rstObj.Close
                    Printer.FontSize = IIf(NTamanoLetra > 0, NTamanoLetra, Printer.FontSize)
                Next
                rst.MoveNext
                If glsEmpresa = "03" And strSerie = "001" Then
                    wsum = wsum + IIf(intScale = 6, 5, 1) + numEntreLineasAdicional
                Else
                    wsum = wsum + IIf(intScale = 6, 4, 1) + numEntreLineasAdicional
                End If
            Loop
            rst.Close
            
            csql = "SELECT TotalVVNeto FROM docventasdet d, productos p " & _
                   "WHERE d.idProducto = p.idProducto AND d.idEmpresa = p.idEmpresa " & _
                   "AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' " & _
                   "AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "' AND p.idNivel = '10010062' "
            rst.Open csql, Cn, adOpenStatic, adLockReadOnly
            
            Do While Not rst.EOF
                If glsEmpresa = "03" Then
                    ImprimeXY Format(Val(rst.Fields("TotalVVNeto") & ""), "0.00"), "N", 10, 115, 62, 2, 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                Else
                    ImprimeXY Format(Val(rst.Fields("TotalVVNeto") & ""), "0.00"), "N", 10, 95, 97, 2, 0, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                End If
                
                rst.MoveNext
            Loop
            rst.Close
        Else
            '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
            strCampos = ""
            If intRegistros = 0 Then
                csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999'  ORDER BY impY,impX"
            Else
                csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'  ORDER BY impY,impX"
            End If
            rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            Do While Not rst.EOF
                strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
                rst.MoveNext
            Loop
            strCampos = left(strCampos, Len(strCampos) - 1)
            rst.Close
            
            '--- Traemos la data de los campos seleccionados arriba
            csql = "SELECT " & strCampos & " FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
            rst.Open csql, Cn, adOpenStatic, adLockReadOnly
            wsum = 0
            Do While Not rst.EOF
                For i = 0 To rst.Fields.Count - 1
                    '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
                    If intRegistros = 0 Then
                        csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' "
                    Else
                        csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'"
                    End If
                    rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                    If Not rstObj.EOF Then
                        strTipoDato = rstObj.Fields("tipoDato")
                        intLong = rstObj.Fields("impLongitud")
                        intX = rstObj.Fields("impX")
                        intY = rstObj.Fields("impY")
                        intDec = Val("" & rstObj.Fields("Decimales"))
                        
                        ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                    rstObj.Close
                Next
                rst.MoveNext
                wsum = wsum + IIf(intScale = 6, 4, 1) + numEntreLineasAdicional
            Loop
            rst.Close
        End If
        
    Else
        '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
        strCampos = ""
        If intRegistros = 0 Then
            csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999'  ORDER BY impY,impX"
        Else
            csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'  ORDER BY impY,impX"
        End If
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
            rst.MoveNext
        Loop
        strCampos = left(strCampos, Len(strCampos) - 1)
        rst.Close
        
        '--- Traemos la data de los campos seleccionados arriba
        csql = "SELECT " & strCampos & " FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        wsum = 0
        Do While Not rst.EOF
            For i = 0 To rst.Fields.Count - 1
                '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
                If intRegistros = 0 Then
                    csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' "
                Else
                    csql = "SELECT tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'"
                End If
                rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rstObj.EOF Then
                    strTipoDato = rstObj.Fields("tipoDato")
                    intLong = rstObj.Fields("impLongitud")
                    intX = rstObj.Fields("impX")
                    intY = rstObj.Fields("impY")
                    intDec = Val("" & rstObj.Fields("Decimales"))
    
                    If rst.Fields(i).Name = "GlsProducto" Then
'                        nfiladetalle = 0: nvarfila = 0: nvarfilatotal = 0
'                        xfila = Val(rstObj.Fields("impY")) + wsum
'                        ncol = Val(rstObj.Fields("impX"))
'
'                        Linea_l = rst.Fields(i) & ""
'                        margenIzq = 0
'                        If margenIzq < 0 Or margenIzq > 40 Then
'                            margenIzq = 0
'                        End If
'
'                        longTotal = Val("" & rstObj.Fields("impLongitud"))
'                        If longTotal < margenIzq Or longTotal > 136 Then
'                            longTotal = Val("" & rstObj.Fields("impLongitud"))
'                        End If
'
'                        sCabecera = Linea_l
'                        s = LoopPropperWrap(sCabecera, longTotal, pwLeft)
'                        sCabecera = ""
'
'                        ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila + nSumFilasEital, ncol, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
'                        If StrMsgError <> "" Then GoTo Err
'
'                        xfila = xfila + 4
'                        nvarfila = nvarfila + 4
'                        filas_detalle = 0
'
'                        If Len(Linea_l) <= Len(s) Then
'                        Else
'                            Do While Len(s) > 0
'                                '--- Añadirle el margen izquierdo
'                                sCabecera = sCabecera & Space$(margenIzq) & s & vbCrLf
'                                s = LoopPropperWrap()
'                                xMemo = s
'                                ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila + nSumFilasEital, ncol, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
'                                If StrMsgError <> "" Then GoTo Err
'                                xfila = xfila + 4
'                                nvarfila = nvarfila + 4
'                                filas_detalle = filas_detalle + 4
'                            Loop
'                        End If
                        
                        Long_total = 0
                        Long_Acumu = 0
                        contadorImp = 0
                        
                        strGlosaDscto = ""
                        
                        Long_total = Len(Trim(rst.Fields(i))) + Len(Trim(strGlosaDscto))
                        nfiladetalle = 0: nvarfila = 0: nvarfilatotal = 0
                        xfila = Val(rstObj.Fields("impY")) + wsum
                        ncol = Val(rstObj.Fields("impX"))
                
                        Linea_l = rst.Fields(i) & "" + strGlosaDscto
                      
                        margenIzq = 0
                        If margenIzq < 0 Or margenIzq > 40 Then
                            margenIzq = 0
                        End If
                        
                        longTotal = Val("" & rstObj.Fields("impLongitud"))
                        If longTotal < margenIzq Or longTotal > 136 Then
                            longTotal = Val("" & rstObj.Fields("impLongitud"))
                        End If
                        
                        sCabecera = Linea_l
                        s = LoopPropperWrap(sCabecera, longTotal, pwLeft)
                        sCabecera = ""
                        
                        ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila + nSumFilasEital, ncol, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                        
                        xfila = xfila + 4
                        filas_detalle = 0
                        Long_Acumu = Long_Acumu + Len(s)
                        
                        'If Len(Linea_l) <= Len(s) Then
                        If Len(Linea_l) < Len(s) Then
                            nvarfila = nvarfila + 4
                            
                        Else
                            Do While Long_total > Long_Acumu
                               '--- Añadirle el margen izquierdo
                                nvarfila = nvarfila + 4
                                sCabecera = sCabecera & Space$(margenIzq) & s & vbCrLf
                                s = LoopPropperWrap()
                                xMemo = s
                                ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila + nSumFilasEital, ncol, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                If StrMsgError <> "" Then GoTo Err
                                xfila = xfila + 4
                                filas_detalle = filas_detalle + 4
                                Long_Acumu = Long_Acumu + Len(s) + Len(vbCrLf)
                                If Long_Acumu = Long_Acumu + Len(s) Then
                                    contadorImp = contadorImp + 1
                                End If
                                If contadorImp = 10 Then Exit Do
                            Loop
                            nvarfila = nvarfila + IIf(rucEmpresa = "20430471750" Or rucEmpresa = "20266578572" Or rucEmpresa = "20552185286" Or rucEmpresa = "20552184981", 0, 4)
                        End If
                        
                    Else
                        ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum + nSumFilasEital, intX, intDec, 0, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                End If
                rstObj.Close
            Next
            rst.MoveNext
            wsum = wsum + IIf(intScale = 6, 4, 1) + numEntreLineasAdicional + nvarfila
        Loop
        rst.Close
    End If
        
    If strImprimeTicket = "S" Then
        wsum = wsum + IIf(intScale = 6, 4, 1)
    Else
        wsum = 0
    End If
    
    strCampos = ""
    If intRegistros = 0 Then
        csql = "SELECT GlsCampo " & _
                "FROM objdocventas " & _
                "WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' and trim(GlsCampo) <> '' and " & _
                "indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' " & _
                " ORDER BY IMPY,IMPX"
    Else
        csql = "SELECT GlsCampo " & _
               "FROM objdocventas " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' and trim(GlsCampo) <> '' and " & _
               "indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "' " & _
               " ORDER BY IMPY,IMPX"
    End If
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
        rst.MoveNext
    Loop
    If Len(strCampos) > 0 Then
        strCampos = left(strCampos, Len(strCampos) - 1)
    End If
    rst.Close
    
    '--- Traemos la data de lo campos seleccionados arriba
    If Len(strCampos) > 0 Then
        csql = "SELECT " & strCampos & ",IdMoneda FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        
        If Not rst.EOF Then
            For i = 0 To rst.Fields.Count - 1
                '--- Traemos datos de impreison por en nombre del campo de la tabla objDocventas
                If strTD = "86" And rst.Fields(i).Name = "idMotivoTraslado" Then
                    csql = "SELECT 'X' AS valor,'T' AS tipoDato, 1 AS impLongitud, impX, impY,0 AS Decimales FROM impMotivosTraslados WHERE idEmpresa = '" & glsEmpresa & "' AND idSerie = '" & strSerie & "' AND idMotivoTraslado = '" & Trim(rst.Fields(i) & "") & "'"
                Else
                    If intRegistros = 0 Then
                        csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999'"
                    Else
                        csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, Decimales FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND GlsCampo = '" & rst.Fields(i).Name & "' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'"
                    End If
                End If
                
                rstObj.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rstObj.EOF Then
                    If strImprimeTicket = "S" Then
                        If rst.Fields(i).Name = "TotalPrecioVenta" Then
                            If rst.Fields("IdMoneda") = "PEN" Then
                                ImprimeXY "Total :  S/ ", "T", 20, intY + wsum + nSumFilasEital, 1, 0, 0, StrMsgError
                            Else
                                ImprimeXY "Total : US$. ", "T", 20, intY + wsum + nSumFilasEital, 1, 0, 0, StrMsgError
                            End If
                        End If
                    End If
                    
                    If strImprimeTicket = "S" Then
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), intY + wsum + nSumFilasEital, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        wsum = wsum + IIf(intScale = 6, 4, 1)
                    Else
                        If strTD = "86" Then
                            ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + nSumFilasEital, rstObj.Fields("impX"), 0, 0, StrMsgError
                        Else
                            If rucEmpresa = "20119041148" Or rucEmpresa = "20119040842" Or rucEmpresa = "20530611681" Then
                                If strTD = "07" And rst.Fields(i).Name = "GlsDocReferencia" Then
                                    If left(rstObj.Fields("valor"), 3) = "Fac" Then iddocumentoRef = "01"
                                    If left(rstObj.Fields("valor"), 3) = "Bov" Then iddocumentoRef = "03"
                                    
                                    iddocumentoRef = traerCampo("docventas", "FecEmision", "iddocumento", iddocumentoRef, True, "idSerie = '" & Mid(rstObj.Fields("valor"), 5, 3) & "' and iddocventas = '" & right(rstObj.Fields("valor"), 8) & "' ")
                                    
                                    ImprimeXY iddocumentoRef, "F", 10, 55, 184, 0, 0, StrMsgError
                                    ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                
                                Else
                                    ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                End If
                            Else
                                ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum + nSumFilasEital, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            End If
                            
                        End If
                    End If
                    If StrMsgError <> "" Then GoTo Err
                End If
                rstObj.Close
            Next
        End If
        rst.Close
    End If
    
    '------------------------------------------------------------------------------------------------
    '--- IMPRIME ETIQUETAS FINAL
    If strImprimeTicket = "S" Then
        wsum = wsum + IIf(intScale = 6, 4, 1)
    Else
        wsum = 0
    End If
    
    csql = "SELECT Etiqueta,impX,impY,indDirSucursal,indSerieEtiquetera,indSoloTicketFactura,indRUCCliente,indRazonSocial,indIGVTotal,indHora,indVendedor,indDirecCliente,indUsuario " & _
           "FROM objetiquetasdoc " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & strTD & "' AND idSerie = '" & _
           strSerie & "'  AND tipoObj = 'T' ORDER BY impY,impX"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    Do While Not rst.EOF
        StrTexto = ""
        If (rst.Fields("indRUCCliente") & "") = 0 And (rst.Fields("indRazonSocial") & "") = 0 And (rst.Fields("indIGVTotal") & "") = 0 And (rst.Fields("indHora") & "") = 0 And (rst.Fields("indVendedor") & "") = 0 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = rst.Fields("Etiqueta") & ""
                End If
            End If
        End If
                        
        If (rst.Fields("indDirSucursal") & "") = 1 Then StrTexto = StrTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        '08001 FACTURA
        '08002 BOLETA
        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + strRUCCliente
                End If
            Else
                StrTexto = rst.Fields("Etiqueta") & ""
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrGlsCliente  '2
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrGlsCliente  '2
                End If
            End If
        End If
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                     StrTexto = StrTexto + StrTotalIGV
                End If
            Else
                StrTexto = rst.Fields("Etiqueta") & ""
                StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indHora") & "") = 1 Then
            StrTexto = rst.Fields("Etiqueta") & ""
            StrTexto = StrTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        End If
                           
        If (rst.Fields("indVendedor") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrGlsVendedorCampo
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrGlsVendedorCampo
                End If
            End If
        End If
        
        If (rst.Fields("indDirecCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrDirecCliente
                End If
            Else
                If StrTipoTicket = "08002" Then 'BOLETA
                    StrTexto = rst.Fields("Etiqueta") & ""
                    StrTexto = StrTexto + StrDirecCliente
                End If
            End If
        End If
        
        If Len(Trim(StrTexto)) > 0 Then
            ImprimeXY StrTexto, "T", Len(StrTexto), Val(rst.Fields("impY").Value) + nSumFilasEital, rst.Fields("impX"), 0, 0, StrMsgError
            wsum = wsum + IIf(intScale = 6, 4, 1)
        End If
        If StrMsgError <> "" Then GoTo Err
        
        rst.MoveNext
    Loop
    rst.Close
    '------------------------------------------------------------------------------------------------
    Set rst = Nothing
    Set rstObj = Nothing
                                   
    Printer.Print Chr$(149)
    Printer.Print ""
    Printer.Print ""
    Printer.EndDoc
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Printer.KillDoc
End Sub

Private Sub SinchiImpSp(ByVal Fiel As String, ByVal ObjValor As String, ByVal ObjTipoDato As String, ByVal ObjLong As Integer, ByVal ObjImpy As Integer, ByVal ObjImpx As Integer, ByVal ObjImpDecim As Double, ByVal wsum As Double, ByVal iddocumentoRef As String, StrMsgError As String, strNumDoc As String, strSerie As String, strTD As String, idMoneda As String)
 
        If strTD = "07" And Fiel = "GlsDocReferencia" Then
        
            If left(ObjValor, 3) = "Fac" Then iddocumentoRef = "01"
            If left(ObjValor, 3) = "Bov" Then iddocumentoRef = "03"
            
            iddocumentoRef = traerCampo("docventas", "FecEmision", "iddocumento", iddocumentoRef, True, "idSerie = '" & Mid(ObjValor, 5, 3) & "' and iddocventas = '" & right(ObjValor, 8) & "' ")
            ImprimeXY iddocumentoRef, "F", 10, 55, 184, 0, 0, StrMsgError
            ImprimeXY ObjValor & "", ObjTipoDato, ObjLong, ObjImpy + wsum, ObjImpx, Val("" & ObjImpDecim), 0, StrMsgError
            
        ElseIf Fiel = "TotalPrecioVenta" And traerCampo("Docventas", "indVtaGratuita", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento='" & strTD & "'") = "1" Then
            ImprimeXY "0.00" & "", ObjTipoDato, ObjLong, ObjImpy + wsum, ObjImpx, Val("" & ObjImpDecim), 0, StrMsgError
            
        ElseIf Fiel = "totalLetras" And traerCampo("Docventas", "indVtaGratuita", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento='" & strTD & "'") = "1" Then
            ImprimeXY "SON: 0.00 " & IIf(idMoneda = "PEN", "SOLES", "DOLARES AMERICANOS") & "", ObjTipoDato, ObjLong, ObjImpy + wsum, ObjImpx, Val("" & ObjImpDecim), 0, StrMsgError
            
        Else
            ImprimeXY ObjValor & "", ObjTipoDato, ObjLong, ObjImpy + wsum, ObjImpx, Val("" & ObjImpDecim), 0, StrMsgError
        End If
        
End Sub

Private Sub MMGuiaRemison(rstObj As ADODB.Recordset, rst As ADODB.Recordset, i, StrMsgError As String)
On Error GoTo Err
Dim strDNITran As String
Dim strGlsTran As String
Dim strTIPDOC As String
 

    If rst.Fields(i).Name = "FecIniTraslado" Then
        'IMPRIME EL DIA
        ImprimeXY Format(Day(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
        If StrMsgError <> "" Then GoTo Err

        'IMPRIME EL MES
        ImprimeXY Format(Month(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY"), rstObj.Fields("impX") + 15, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
        If StrMsgError <> "" Then GoTo Err

        'IMPRIME EL AÑO
        ImprimeXY right((Year(rstObj.Fields("valor"))), 4) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 30, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
        If StrMsgError <> "" Then GoTo Err

    ElseIf rst.Fields(i).Name = "GlsVehiculo" Then
        ImprimeXY rstObj.Fields("valor") & " " & rst.Fields("Marca").Value & " " & rst.Fields("Placa").Value, rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
    ElseIf rst.Fields(i).Name = "idPerChofer" Then
        strDNITran = traerCampo("Personas", "RUC", "idPersona", rstObj.Fields("valor"), False)
        strGlsTran = traerCampo("Personas", "GlsPersona", "idPersona", rstObj.Fields("valor"), False)
    
        ImprimeXY strGlsTran & "", rstObj.Fields("tipoDato"), "80", "242", "17", Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        'Imprime el RUC y/o DNI
        ImprimeXY strDNITran & "", rstObj.Fields("tipoDato"), "12", "242", "118", Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
    ElseIf rst.Fields(i).Name = "GlsDocReferencia" Then
        strTIPDOC = traerCampo("Documentos", "GlsDocumento", "AbreDocumento", left(rstObj.Fields("valor"), 3), False)
        ImprimeXY strTIPDOC & "", rstObj.Fields("tipoDato"), "20", "242", "155", Val("" & rstObj.Fields("Decimales")), "1", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        ImprimeXY right(rstObj.Fields("valor"), 8) & "", rstObj.Fields("tipoDato"), "8", "242", "185", Val("" & rstObj.Fields("Decimales")), "1", StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
     
    Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub PredeterminaImpresora(indPrinter, strTD, p As Object, StrMsgError As String)
On Error GoTo Err
' , indPrinter, strTD, p, StrMsgError
indPrinter = False
    If strTD = "01" Then '--- FACTURA
        If leeParametro("IMPRESORA_FACTURA_PREDETERMINADA") = "1" Then
            indPrinter = True
        Else
            For Each p In Printers
               If InStr(UCase(p.DeviceName), "FACTURA") > 0 Then
                  Set Printer = p
                  indPrinter = True
                  Exit For
               End If
            Next p
        End If
    ElseIf strTD = "03" Then '--- BOLETA
            For Each p In Printers
                If InStr(UCase(p.DeviceName), "BOLETA") > 0 Then
                    Set Printer = p
                    indPrinter = True
                    Exit For
                End If
            Next p
    ElseIf strTD = "86" Then '--- GUIA
            For Each p In Printers
                If InStr(UCase(p.DeviceName), "GUIA") > 0 Then
                    Set Printer = p
                    indPrinter = True
                    Exit For
                End If
            Next p
    End If
    
    If indPrinter = False Then
        For Each p In Printers
            If InStr(UCase(p.DeviceName), "GENERAL") > 0 Then
                Set Printer = p
                indPrinter = True
                Exit For
            End If
        Next p
    End If
            
    If indPrinter = False Then
        For Each p In Printers
           ''If p.Port = "Ne03:" Then
           If p.Port = "LPT1:" Then
              Set Printer = p
              Exit For
           End If
        Next p
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub ImprimeAbreviaturaVen(StrMsgError As String, PIdentificador As Long, PValor As String)
On Error GoTo Err
Dim CSqlC                       As String
Dim RsC                         As New ADODB.Recordset

    CSqlC = "Select *,'" & PValor & "' Valor From ObjDocVentas Where Identificador = " & PIdentificador & ""
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not RsC.EOF Then
    
        ImprimeXY traerCampo("Vendedores", "Abreviatura", "IdVendedor", RsC.Fields("valor"), True), RsC.Fields("tipoDato"), RsC.Fields("impLongitud"), RsC.Fields("impY"), RsC.Fields("impX"), Val("" & RsC.Fields("Decimales")), Val("" & RsC.Fields("intNumFilas")), StrMsgError
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    RsC.Close: Set RsC = Nothing
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub ImprimeEtiquetas(StrMsgError As String, strRUCCliente As String, strTD As String, strSerie As String, StrGlsCliente As String, StrTotalIGV As String, StrGlsMotivoTraslado As String, strNumDoc As String)
On Error GoTo Err
Dim rst                         As New ADODB.Recordset
Dim StrTexto                    As String
Dim CSqlC                       As String
Dim StrGlsMG                    As String

     '--- Evalua la etiqueta de muestra gratuita - SOLVET
    If traerCampo("Docventas", "indVtaGratuita", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento='" & strTD & "'") = "1" Then
        StrGlsMG = "And indMuestraGratuita In('0','1') "
    Else
        StrGlsMG = "And indMuestraGratuita Not In('1') "
    End If
    
    '-------------------------------------------------------------------
    '--- ETIQUETAS
    '--- Traemos las etiquetas configuradas para el documento y la serie
    csql = "SELECT Etiqueta,impX,impY,idobjetiquetasdoc,indDirSucursal,indSerieEtiquetera,indSoloTicketFactura," & _
           "indRUCCliente,indRazonSocial,indIGVTotal,indHora,indDestinatarioGuia,indMotivoTrasladoGuia,indUsuario " & _
           "FROM objetiquetasdoc " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & strTD & "' " & StrGlsMG & _
           "AND idSerie = '" & strSerie & "' " & _
           "AND tipoObj = 'C' ORDER BY IMPY,IMPX"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        
    Do While Not rst.EOF
        StrTexto = rst.Fields("Etiqueta") & ""
        If (rst.Fields("indDirSucursal") & "") = 1 Then StrTexto = StrTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then StrTexto = StrTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        CSqlC = traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        
        '--- 08001 FACTURA  '--- 08002 BOLETA
        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                     StrTexto = StrTexto + strRUCCliente
                End If
            Else
                StrTexto = StrTexto + CSqlC
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                     StrTexto = StrTexto + StrGlsCliente  '1
                End If
            Else
                StrTexto = StrTexto + CSqlC
            End If
        End If
                        
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If StrTipoTicket = "08001" Then 'FACTURA
                     StrTexto = StrTexto + StrTotalIGV
                End If
            Else
                StrTexto = StrTexto + CSqlC
            End If
        End If
            
        If (rst.Fields("indHora") & "") = 1 Then StrTexto = StrTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        '-------------------------------------------------------------------
        '------ BUSCA SI ES SUCURSAL, EXTRAE RAZON SOCIAL Y RUC DE LA EMPRESA
        If (rst.Fields("indDestinatarioGuia") & "") = 1 Then
            cselect = "SELECT idSucursal FROM sucursales WHERE idSucursal = '" & StrCodigoCliente & "' AND idEmpresa = '" & glsEmpresa & "'"
            If RsD.State = adStateOpen Then RsD.Close
            RsD.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
            If Not RsD.EOF Then
                RsD.Close: Set RsD = Nothing
                
                cselect = "SELECT GlsEmpresa, RUC FROM empresas WHERE idEmpresa = '" & glsEmpresa & "'"
                If RsD.State = adStateOpen Then RsD.Close
                RsD.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
                If Not RsD.EOF Then
                    StrTexto = StrTexto & Trim(RsD.Fields("GlsEmpresa") & "") & Space(5) & "RUC: " & Trim(RsD.Fields("RUC") & "")
                End If
            Else
                StrTexto = StrTexto + StrGlsCliente & Space(5) & "RUC/DNI: " & strRUCCliente
            End If
            RsD.Close: Set RsD = Nothing
        End If
        '-------------------------------------------------------------------
        If (rst.Fields("indMotivoTrasladoGuia") & "") = 1 Then
            StrTexto = StrTexto & StrGlsMotivoTraslado
        End If
        '-------------------------------------------------------------------
        ImprimeXY StrTexto, "T", Len(StrTexto & ""), rst.Fields("impY"), rst.Fields("impX"), 0, 0, StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        rst.MoveNext
    Loop
    rst.Close
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub DireccionRecojo(strTD, strSerie, strNumDoc, StrIndDirRecojox, Gls_Distritoxx, Gls_Provxx, Gls_Depaxx, StrMsgError As String)
Dim rstienda        As New ADODB.Recordset
On Error GoTo Err

    Cad_Mysql = "select b.idpais,b.iddistrito,b.GlsDirRecojo,c.glsubigeo glsubigeoDis,d.glsubigeo glsubigeoProv,e.glsubigeo glsubigeoDep from docventas a " & _
                "inner join dirrecojos b on a.iddirrecojo = b.iddirrecojo and a.idempresa  = b.idempresa " & _
                "inner join ubigeo c on b.iddistrito  = c.iddistrito and c.idpais  = b.idpais " & _
                "inner join ubigeo d  on  d.idProv = c.idProv and d.idDpto = c.idDpto and d.idProv <> '00' and d.idDist = '00' " & _
                "inner join ubigeo e  on  e.idDpto = c.idDpto and e.idProv = '00' " & _
                "where a.iddocumento = '" & strTD & "' and  a.idserie = '" & strSerie & "' and a.iddocventas = '" & strNumDoc & "' "
    
    If rstienda.State = 1 Then rstienda.Close
    rstienda.Open Cad_Mysql, Cn, adOpenStatic, adLockReadOnly
    If Not rstienda.EOF Then
        StrIndDirRecojox = Trim("" & rstienda.Fields("GlsDirRecojo"))
        Gls_Distritoxx = Trim("" & rstienda.Fields("glsubigeoDis"))
        Gls_Provxx = Trim("" & rstienda.Fields("glsubigeoProv"))
        Gls_Depaxx = Trim("" & rstienda.Fields("glsubigeoDep"))
    End If
                        
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
