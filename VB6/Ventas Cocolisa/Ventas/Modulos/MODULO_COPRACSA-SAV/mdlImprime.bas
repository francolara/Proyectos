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

Public Sub imprimeDocVentas(strTD As String, strNumDoc As String, strSerie As String, ByRef StrMsgError As String)
On Error GoTo ERR
Dim wcont, wsum             As Integer
Dim wfila, wcolu            As Integer
Dim rst                     As New ADODB.Recordset
Dim rstObj                  As New ADODB.Recordset
Dim p                       As Object
Dim indPrinter              As Boolean
Dim strCampos               As String
Dim intScale                As Integer
Dim strTipoFecDoc           As String
Dim strImprimeTicket        As String
Dim strTipoTicket           As String
Dim strRUCCliente           As String
Dim strGlsCliente           As String
Dim strTotalIGV             As String
Dim strGlsVendedorCampo     As String
Dim strDirecCliente         As String
Dim strMotivoTraslado       As String
Dim strCodigoCliente        As String
Dim cselect                 As String
Dim nfontletra              As String
Dim numEntreLineasAdicional As Integer
Dim intRegistros            As Integer
Dim indventasterceros       As String
Dim indterceros             As Boolean
Dim TbConsultaRef           As New ADODB.Recordset
Dim cSqlRef                 As String
Dim nFilRef                 As Integer
Dim strFormatoImpFec        As String
Dim NTamanoLetra            As Integer
Dim idtiendacli             As String
Dim rstienda                As New ADODB.Recordset
Dim CodDistrito             As String
Dim codPais                 As String
Dim Gls_Pais                As String
Dim Gls_Depa                As String
Dim Gls_Prov                As String
Dim Gls_Distrito            As String
Dim Cad_Mysql               As String
Dim impxsp                  As Integer
Dim impysp                  As Integer
Dim LongSp                  As Integer
Dim CSqlC                   As String
Dim RsC                     As New ADODB.Recordset
Dim strIdDocumento          As String
Dim strIdDocventas          As String
Dim strIdSerie              As String
Dim stridEmpresa            As String
Dim stridSucursal           As String
Dim stridPersona            As String
Dim strglspersona           As String
Dim strRUCPersona           As String
Dim stridTd                 As String
Dim strcadll                As String
Dim rsrecorset              As New ADODB.Recordset
Dim rucEmpresa              As String
Dim rsd                     As New ADODB.Recordset
Dim strTexto                As String
Dim strimpNF                As String
Dim strimpSF                As String
Dim StrimpFecf              As String
Dim FecRefNC                As String
Dim numDocorigenNC          As String
Dim serieDocorigenNC        As String
Dim tipodocorigenNC         As String
Dim rsref                   As New ADODB.Recordset
Dim strFechaDR              As String
Dim nfiladetalle            As Integer, nvarfila As Integer, nvarfilatotal As Integer
Dim strTipoDato             As String
Dim intLong                 As Integer
Dim intX                    As Integer
Dim intY                    As Integer
Dim intDec                  As Integer
Dim iddocumentoRef          As String
Dim Long_total              As Integer
Dim Long_Acumu              As Integer
Dim contadorImp             As Integer
Dim GlsMarca                As String
Dim TotFlete                As Double
Dim strcaddref              As String
Dim strGlsMG                As String
Dim ccadenafecha            As String
Dim strGlosaDscto           As String
Dim strMtoSinDsc            As Double
Dim strMO                   As String
Dim dblPorcDsc              As Double
Dim dblMtoTotal             As Double
Dim dblMtoTotEnt            As String
Dim dblMtoTotVuelto         As String
Dim boolEst                 As Boolean

    rucEmpresa = traerCampo("empresas", "ruc", "idEmpresa", glsEmpresa, False)
    intScale = 6
    Printer.ScaleMode = intScale
     
    '--- SELECCIONAMOS IMPRESORA
    PredeterminaImpresora indPrinter, strTD, p, StrMsgError
    If StrMsgError <> "" Then GoTo ERR
        
    NTamanoLetra = Val(traerCampo("empresas", "NTamanoLetra", "idEmpresa", glsEmpresa, False) & "")
    Printer.FontSize = IIf(NTamanoLetra > 0, NTamanoLetra, Printer.FontSize)
        
    strTipoFecDoc = traerCampo("documentos", "TipoImpFecha", "idDocumento", strTD, False)
    strImprimeTicket = traerCampo("documentos", "indImprimeTicket", "idDocumento", strTD, False)
    strFormatoImpFec = "" & traerCampo("documentos", "FormatoImpFecha", "idDocumento", strTD, False)
    
    numEntreLineasAdicional = Val("" & traerCampo("seriesdocumento", "espacioLineasImp", "idSerie", strSerie, True, " idsucursal = '" & glsSucursal & "' and iddocumento = '" & strTD & "'"))
        
    csql = "SELECT idtienda,GlsCliente,RUCCliente,TotalIGVVenta,idTipoTicket,GlsVendedorCampo,dirCliente,idMotivoTraslado,idPerCliente,idMotivoNCD " & _
           "FROM docventas " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'" & _
           "  AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    rsd.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsd.EOF Then
        strTipoTicket = "" & rsd.Fields("idTipoTicket")
        strRUCCliente = "" & rsd.Fields("RUCCliente")
        strGlsCliente = "" & rsd.Fields("GlsCliente")
        strTotalIGV = CStr(Format("" & rsd.Fields("TotalIGVVenta"), "###,##0.00"))
        strGlsVendedorCampo = "" & rsd.Fields("GlsVendedorCampo")
        strDirecCliente = "" & rsd.Fields("dirCliente")
        strMotivoTraslado = "" & rsd.Fields("idMotivoTraslado")
        strCodigoCliente = "" & rsd.Fields("idPerCliente")
        idtiendacli = Trim("" & rsd.Fields("idtienda"))
    End If
    
    If Len(Trim(strMotivoTraslado)) > 0 Then
        csql = "SELECT GlsMotivoTraslado FROM motivostraslados WHERE idMotivoTraslado = '" & strMotivoTraslado & "'"
        If rsd.State = adStateOpen Then rsd.Close
        rsd.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rsd.EOF Then
            strGlsMotivoTraslado = Trim(rsd.Fields("GlsMotivoTraslado") & "")
        End If
    End If
    rsd.Close: Set rsd = Nothing
        
    If Len(Trim((traerCampo("sucursales", "TipoLetra", "idSucursal", glsSucursal, True)))) > 0 Then
        Printer.FontName = traerCampo("sucursales", "TipoLetra", "idSucursal", glsSucursal, True)
    End If
    
     '--- Evalua la etiqueta de muestra gratuita - SOLVET
    If traerCampo("Docventas", "indVtaGratuita", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento='" & strTD & "'") = "1" Then
        strGlsMG = "And indMuestraGratuita In('0','1') "
    Else
        strGlsMG = "And indMuestraGratuita Not In('1') "
    End If
     
    '-------------------------------------------------------------------
    '--- ETIQUETAS
    '--- Traemos las etiquetas configuradas para el documento y la serie
    csql = "SELECT Etiqueta,impX,impY,idobjetiquetasdoc,indDirSucursal,indSerieEtiquetera,indSoloTicketFactura," & _
           "indRUCCliente,indRazonSocial,indIGVTotal,indHora,indDestinatarioGuia,indMotivoTrasladoGuia,indUsuario " & _
           "FROM objetiquetasdoc " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & strTD & "' " & strGlsMG & _
           "AND idSerie = '" & strSerie & "' " & _
           "AND tipoObj = 'C' ORDER BY IMPY,IMPX"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        
    Do While Not rst.EOF
        strTexto = rst.Fields("Etiqueta") & ""
        If (rst.Fields("indDirSucursal") & "") = 1 Then strTexto = strTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        '--- 08001 FACTURA  '--- 08002 BOLETA
        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                     strTexto = strTexto + strRUCCliente
                End If
            Else
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                     strTexto = strTexto + strGlsCliente  '1
                End If
            Else
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
                        
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                     strTexto = strTexto + strTotalIGV
                End If
            Else
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
            
        If (rst.Fields("indHora") & "") = 1 Then strTexto = strTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        '-------------------------------------------------------------------
        '------ BUSCA SI ES SUCURSAL, EXTRAE RAZON SOCIAL Y RUC DE LA EMPRESA
        If (rst.Fields("indDestinatarioGuia") & "") = 1 Then
            cselect = "SELECT idSucursal FROM sucursales WHERE idSucursal = '" & strCodigoCliente & "' AND idEmpresa = '" & glsEmpresa & "'"
            If rsd.State = adStateOpen Then rsd.Close
            rsd.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rsd.EOF Then
                rsd.Close: Set rsd = Nothing
                
                cselect = "SELECT GlsEmpresa, RUC FROM empresas WHERE idEmpresa = '" & glsEmpresa & "'"
                If rsd.State = adStateOpen Then rsd.Close
                rsd.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rsd.EOF Then
                    strTexto = strTexto & Trim(rsd.Fields("GlsEmpresa") & "") & Space(5) & "RUC: " & Trim(rsd.Fields("RUC") & "")
                End If
            Else
                strTexto = strTexto + strGlsCliente & Space(5) & "RUC/DNI: " & strRUCCliente
            End If
            rsd.Close: Set rsd = Nothing
        End If
        '-------------------------------------------------------------------
        If (rst.Fields("indMotivoTrasladoGuia") & "") = 1 Then
            strTexto = strTexto & strGlsMotivoTraslado
        End If
        '-------------------------------------------------------------------
        ImprimeXY strTexto, "T", Len(strTexto & ""), rst.Fields("impY"), rst.Fields("impX"), 0, 0, StrMsgError
        If StrMsgError <> "" Then GoTo ERR
        
        rst.MoveNext
    Loop
    rst.Close

    '-------------------------------------------------------------------
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
    
    '--- SERVICIOS MEDIO MUNDO
    If strTD = "01" And rucEmpresa = "20542001969" Then
        If strSerie = "003" Then '--- SERIE BOSH CARD SERVICE
            ImprimeXY traerCampo("UnidadProduccion A Inner Join DocVentas B On A.IdEmpresa = B.IdEmpresa And A.CodUnidProd = B.IdUPP", "A.DescUnidad", "B.IdEmpresa", glsEmpresa, False, "B.IdSucursal = '" & glsSucursal & "' And B.IdDocumento = '" & strTD & "' And B.IdSerie = '" & strSerie & "' And B.IdDocVentas = '" & strNumDoc & "'"), "T", 55, 262, 6, 0, 0, StrMsgError
            If StrMsgError <> "" Then GoTo ERR
        Else
            ImprimeXY traerCampo("UnidadProduccion A Inner Join DocVentas B On A.IdEmpresa = B.IdEmpresa And A.CodUnidProd = B.IdUPP", "A.DescUnidad", "B.IdEmpresa", glsEmpresa, False, "B.IdSucursal = '" & glsSucursal & "' And B.IdDocumento = '" & strTD & "' And B.IdSerie = '" & strSerie & "' And B.IdDocVentas = '" & strNumDoc & "'"), "T", 55, 125, 1, 0, 0, StrMsgError
            If StrMsgError <> "" Then GoTo ERR
        End If
    End If
    
    '--- SERVICIOS MEDIO MUNDO
    '--- Imprime directo para que no se corra la hoja
    If strTD = "03" And rucEmpresa = "20542001969" Then
        If strSerie = "004" Then '--- SERIE BOSH CARD SERVICE
            ImprimeXY traerCampo("UnidadProduccion A Inner Join DocVentas B On A.IdEmpresa = B.IdEmpresa And A.CodUnidProd = B.IdUPP", "A.DescUnidad", "B.IdEmpresa", glsEmpresa, False, "B.IdSucursal = '" & glsSucursal & "' And B.IdDocumento = '" & strTD & "' And B.IdSerie = '" & strSerie & "' And B.IdDocVentas = '" & strNumDoc & "'"), "T", 55, 262, 6, 0, 0, StrMsgError
            If StrMsgError <> "" Then GoTo ERR
            
        Else
            ImprimeXY traerCampo("DocVentas", "Format(TotalPrecioVenta,2)", "idSerie", strSerie, True, "IdSucursal = '" & glsSucursal & "' And IdDocumento = '" & strTD & "' And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strNumDoc & "'"), "T", 15, 118, 181, 2, 0, StrMsgError
            If StrMsgError <> "" Then GoTo ERR
            
            ImprimeXY traerCampo("DocVentas", "IF(idMoneda='PEN','S/.','US$')", "idSerie", strSerie, True, "IdSucursal = '" & glsSucursal & "' And IdDocumento = '" & strTD & "' And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strNumDoc & "'"), "T", 15, 118, 170, 0, 0, StrMsgError
            If StrMsgError <> "" Then GoTo ERR
        End If
    End If
    
    '---- AGREGADO EL 06/05/10 VENTAS A TERCEROS DEPENDIENDO SI EL CAMPO IND VENTAS TERCEROS ESTA CON 1
    indterceros = False
    If strTD = "86" Then
        indventasterceros = traerCampo("clientes", "indventasterceros", "idcliente", strCodigoCliente, True)
        If indventasterceros = "1" Then
            indterceros = True
        Else
            indterceros = False
        End If
    End If
    
    '-----------------------------------------------------------------------------------------------------------------------------------
    '--- Traemos la data de lo campos seleccionados arriba
    '-----------------------------------------------------------------------------------------------------------------------------------
    csql = "SELECT " & strCampos & " , '" & strGlsMotivoTraslado & "' As GlsMotivoNCD FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF Then
        For i = 0 To rst.Fields.Count - 1
            '--- Traemos datos de impresion por en nombre del campo de la tabla objDocventas
            If (strTD = "86" And rst.Fields(i).Name = "idMotivoTraslado") Or (strTD = "07" And rst.Fields(i).Name = "idMotivoTraslado") Then
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
                    '--- IMPRIME EL DIA
                    ImprimeXY Format(Day(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                    
                    '--- IMPRIME EL MES
                    ImprimeXY strArregloMes(Val(Month(rstObj.Fields("valor")))) & "", "T", 10, rstObj.Fields("impY"), rstObj.Fields("impX") + 20, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                    
                    '--- IMPRIME EL AÑO
                    ImprimeXY right((Year(rstObj.Fields("valor"))), 2) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 60, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                    
                ElseIf rst.Fields(i).Name = "FecEmision" And strTipoFecDoc = "C" Then   '---- PARA CADILLO
                    ccadenafecha = Format(Day(rstObj.Fields("valor")), "00") & " de " & strArregloMes(Val(Month(rstObj.Fields("valor")))) & " del " & right((Year(rstObj.Fields("valor"))), 4)
                    ImprimeXY ccadenafecha, "T", 30, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                    
                ElseIf rst.Fields(i).Name = "FecEmision" And strFormatoImpFec = "S" Then
                    '--- IMPRIME EL DIA
                    ImprimeXY Format(Day(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                    
                    '--- IMPRIME EL MES
                    ImprimeXY Format(Month(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY"), rstObj.Fields("impX") + 15, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
        
                    '--- IMPRIME EL AÑO
                    ImprimeXY right((Year(rstObj.Fields("valor"))), 4) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 30, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                
               ElseIf (rst.Fields(i).Name = "GlsVehiculo" Or rst.Fields(i).Name = "FecIniTraslado" Or rst.Fields(i).Name = "idPerChofer") And rucEmpresa = "20542001969" Then
                    MMGuiaRemison rstObj, rst, i, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                    
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
                                CodDistrito = Trim("" & traerCampo("tiendascliente", "idDistrito", "idtdacli", stridTd, True))
                                codPais = Trim("" & traerCampo("tiendascliente", "idPais", "idtdacli", stridTd, True))
                            End If
                        End If
                          
                        Cad_Mysql = " select idDpto, idProv " & _
                                      " FROM ubigeo " & _
                                      " where iddistrito = '" & CodDistrito & "' and idPais = '" & codPais & "' "
                        If rsrecorset.State = 1 Then rsrecorset.Close
                        rsrecorset.Open Cad_Mysql, Cn, adOpenStatic, adLockReadOnly
                                  
                        If Not rsrecorset.EOF Then
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DISTRITO' "))
                                
                                Gls_Distrito = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", CodDistrito, False)
                                
                                ImprimeXY Gls_Distrito & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                            End If
                              
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                              
                                Gls_Prov = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Trim("" & rsrecorset.Fields("idProv")), False, " idDpto = '" & Trim("" & rsrecorset.Fields("idDpto")) & "' and idProv <> '00' and idDist = '00'")
                              
                                ImprimeXY Gls_Prov & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                            End If
                              
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                            
                                Gls_Depa = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Trim("" & rsrecorset.Fields("idDpto")), False, " idProv = '00'")
                                
                                ImprimeXY Gls_Depa & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                                  
                                '--- Imprime glosa cliente
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", "GlsPersona", True, " iddocumento = '" & strTD & "' and Campo = 'GlsPersona' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", "GlsPersona", True, " iddocumento = '" & strTD & "' and Campo = 'GlsPersona' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", "GlsPersona", True, " iddocumento = '" & strTD & "' and Campo = 'GlsPersona' "))
                              
                                ImprimeXY strglspersona & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                                  
                                '--- Imprim ruc cliente
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", "RUCPersona", True, " iddocumento = '" & strTD & "' and Campo = 'RUCPersona' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", "RUCPersona", True, " iddocumento = '" & strTD & "' and Campo = 'RUCPersona' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", "RUCPersona", True, " iddocumento = '" & strTD & "' and Campo = 'RUCPersona' "))
                              
                                ImprimeXY strRUCPersona & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                            End If
                        End If
                        rsrecorset.Close: Set rstienda = Nothing
                          
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo ERR
                    
                    Else
                        ' ----------------------------------------------------------------------------------------------------------------------------
                        If Len(Trim("" & traerCampo("objimpespeciales", "iddocumento", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' "))) > 0 Then
                            Gls_Pais = "": Gls_Depa = "": Gls_Prov = "": Gls_Distrito = "": impxsp = 0: impysp = 0
                            CodDistrito = Trim("" & traerCampo("tiendascliente", "idDistrito", "idtdacli", idtiendacli, True))
                            codPais = Trim("" & traerCampo("tiendascliente", "idPais", "idtdacli", idtiendacli, True))
                        
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
                                    
                                    Gls_Distrito = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", CodDistrito, False)
                                    
                                    ImprimeXY Gls_Distrito & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                    If StrMsgError <> "" Then GoTo ERR
                            End If
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                            
                                Gls_Prov = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Trim("" & rstienda.Fields("idProv")), False, " idDpto = '" & Trim("" & rstienda.Fields("idDpto")) & "' and idProv <> '00' and idDist = '00'")
                            
                                ImprimeXY Gls_Prov & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                            End If
                            
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                            
                                Gls_Depa = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Trim("" & rstienda.Fields("idDpto")), False, " idProv = '00'")
                                
                                ImprimeXY Gls_Depa & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                            End If
                        End If
                        rstienda.Close: Set rstienda = Nothing
                        
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo ERR
                                                
                    ElseIf rucEmpresa = "20430471750" Then
                        strcadll = rstObj.Fields("valor") & ""
                        ImprimeXY left(rstObj.Fields("valor"), 54) & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        ImprimeXY Mid(strcadll, 55, Len(strcadll)) & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + 4, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                    Else
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo ERR
                    End If
                End If
                
            ElseIf rst.Fields(i).Name = "Partida" Then
                    If Len(Trim("" & traerCampo("objimpespeciales", "count(*)", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' "))) > 0 Then
                        Gls_Pais = ""
                        Gls_Depa = ""
                        Gls_Prov = ""
                        Gls_Distrito = ""
                        CodDistrito = Trim("" & traerCampo("Personas", "idDistrito", "idPersona", glsSucursal, False))
                        codPais = Trim("" & traerCampo("Personas", "idPais", "idPersona", glsSucursal, False))
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
                           
                                Gls_Distrito = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", CodDistrito, False)
                                
                                ImprimeXY Gls_Distrito & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                            End If
                            
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'PROVINCIA' "))
                            
                                Gls_Prov = traerCampo("Ubigeo", "GlsUbigeo", "idProv", Trim("" & rstienda.Fields("idProv")), False, " idDpto = '" & Trim("" & rstienda.Fields("idDpto")) & "' and idProv <> '00' and idDist = '00'")
                            
                                ImprimeXY Gls_Prov & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                            End If
                            
                            If Len(Trim("" & traerCampo("objimpespeciales", "Campo", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))) > 0 Then
                                impxsp = Trim("" & traerCampo("objimpespeciales", "impX", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                impysp = Trim("" & traerCampo("objimpespeciales", "impY", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                                LongSp = Trim("" & traerCampo("objimpespeciales", "Longitud", "Variable", rst.Fields(i).Name, True, " iddocumento = '" & strTD & "' and Campo = 'DEPARTAMENTO' "))
                            
                                Gls_Depa = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", Trim("" & rstienda.Fields("idDpto")), False, " idProv = '00'")
                                
                                ImprimeXY Gls_Depa & "", rstObj.Fields("tipoDato"), LongSp, impysp, impxsp, 0, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                            End If
                        End If
                        rstienda.Close: Set rstienda = Nothing
                        
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo ERR
                                                                                            
                    Else
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo ERR
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
                            If StrMsgError <> "" Then GoTo ERR
                        End If
                        If rsref.State = 1 Then rsref.Close: Set rsref = Nothing
                        
                    Else
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                        If StrMsgError <> "" Then GoTo ERR
                    End If
                    
                ElseIf rst.Fields(i).Name = "FecPago" And strTD = "07" And leeParametro("IMPRIME_FECHA_FACTURA") = "1" Then 'imprime la fecha del documento de referencia
                        
                    strFechaDR = traerCampo("Docventas", "FecEmision", "idDocventas", "" & right(rst.Fields("GlsDocReferencia").Value, 8), True, "idSerie= '" & "" & Mid(rst.Fields("GlsDocReferencia").Value, 5, 3) & "' And idSucursal='" & glsSucursal & "' ")
                    ImprimeXY strFechaDR & "", "T", 100, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                
                    '--- Recuperamos el RUC para la empresa Apimas solo para clientes del Pais Bolivia Concatenamos al Ruc N.I.T.
                ElseIf strTD = "01" And rucEmpresa = "20305948277" And rst.Fields(i).Name = "RUCCliente" And traerCampo("Personas", "idPais", "Ruc", rstObj.Fields("valor"), False) = "02005" Then
                    ImprimeXY "N.I.T." & rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                
                '--- Se está utilizando el campo vehiculo para imprimir la fecha del documento de referencia  de NC para ITS
                ElseIf strTD = "07" And rucEmpresa = "20257354041" And rst.Fields(i).Name = "idVehiculo" Then
                    numDocorigenNC = traerCampo("DocReferencia", "numDocReferencia", "numDocOrigen", strNumDoc, True, "  serieDocOrigen= '" & strSerie & "' And tipoDocOrigen = '" & strTD & "' And idSucursal = '" & glsSucursal & "'  ")
                    serieDocorigenNC = traerCampo("DocReferencia", "serieDocReferencia", "numDocOrigen", strNumDoc, True, "  serieDocOrigen= '" & strSerie & "' And tipoDocOrigen = '" & strTD & "' And idSucursal = '" & glsSucursal & "'  ")
                    tipodocorigenNC = traerCampo("DocReferencia", "tipoDocReferencia", "numDocOrigen", strNumDoc, True, "  serieDocOrigen= '" & strSerie & "' And tipoDocOrigen = '" & strTD & "' And idSucursal = '" & glsSucursal & "'  ")
                                                              
                    FecRefNC = traerCampo("Docventas", "FecEmision", "idDocventas", numDocorigenNC, True, "idSerie = '" & serieDocorigenNC & "'  And idSucursal ='" & glsSucursal & "' And idDocumento = '" & tipodocorigenNC & "'  ")
                    ImprimeXY FecRefNC & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                
                '-----------------------------------------------------------------------------
                '---- AQUI TENER EN CUENTA PARA LOS OTROS CLIENTES
                '-----------------------------------------------------------------------------
                ElseIf rst.Fields(i).Name = "ObsDocVentas" Or rst.Fields(i).Name = "ObsDocVentas2" Or rst.Fields(i).Name = "dirCliente" Then
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
                        If StrMsgError <> "" Then GoTo ERR
                            
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
                                If StrMsgError <> "" Then GoTo ERR
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
                            If StrMsgError <> "" Then GoTo ERR
                        End If
                    Else
                        If rst.Fields(i).Name = "vtercerosCliente" Or rst.Fields(i).Name = "vterceroRuc" Or rst.Fields(i).Name = "vtercerosdireccion" Then
                        Else
                             ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                             If StrMsgError <> "" Then GoTo ERR
                        End If
                    End If
                End If
            End If
            rstObj.Close
        Next
    End If
    rst.Close
    '-----------------------------------------------------------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------------------------------------------------------------
        
        
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
                        If StrMsgError <> "" Then GoTo ERR
                            
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
                    ImprimeXY Format(Val(rst.Fields("TotalVVNeto") & ""), "0.00"), "N", 10, 95, 97, 2, 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                Else
                    ImprimeXY Format(Val(rst.Fields("TotalVVNeto") & ""), "0.00"), "N", 10, 95, 97, 2, 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
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
                        If StrMsgError <> "" Then GoTo ERR
                            
                    End If
                    rstObj.Close
                Next
                rst.MoveNext
                wsum = wsum + IIf(intScale = 6, 4, 1) + numEntreLineasAdicional
            Loop
            rst.Close
        End If
        
    Else
        'Medio Mundo Venta directa de productos imprime el código
        If (rucEmpresa = "20542001969" And strTD = "01") Then
            If traerCampo("DocventasDet", "idCentroCosto", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento = '" & strTD & "' And idSucursal = '" & glsSucursal & "' ") = "" Then
                csql = "Update objdocventas Set indImprime = '1' WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and idDocumento = '" & strTD & "' and idSerie = '999'   And GlsCampo = 'CodigoRapido'"
                Cn.Execute (csql)
                
            Else
                Cn.Execute "Update objdocventas Set indImprime = '0' WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999'   And GlsCampo = 'CodigoRapido'"
            End If
        End If
        
        '--- Seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
        strCampos = ""
        If intRegistros = 0 Then
            csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999'  ORDER BY impY,impX"
        Else
            csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "'  ORDER BY impY,impX"
        End If
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        
        If rucEmpresa = "20504973990" Then
            Do While Not rst.EOF
                strCampos = strCampos & "d." & "" & rst.Fields("GlsCampo") & ","
                rst.MoveNext
            Loop
        Else
            Do While Not rst.EOF
                strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
                rst.MoveNext
            Loop
        End If
        
        strCampos = left(strCampos, Len(strCampos) - 1)
        rst.Close
        
        If rucEmpresa = "20504973990" Then
            '--- Traemos la data de los campos seleccionados arriba
            csql = "SELECT " & strCampos & " FROM docventasdet d, productos p WHERE d.idProducto = p.idProducto AND d.idEmpresa = p.idEmpresa AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "' AND p.idNivel <> '11080037' "
        
        Else
            '--- Traemos la data de los campos seleccionados arriba
            csql = "SELECT " & strCampos & " FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
        End If
        
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
                        Long_total = 0
                        Long_Acumu = 0
                        contadorImp = 0
                        
                        strGlosaDscto = ""
                        
                        'Para MM si el item del Pro. tiene descuento imprime glosa
                        If rucEmpresa = "20542001969" And strTD = "01" Then
                            strMtoSinDsc = traerCampo("DocventasDet", "Round(TotalPVBruto,2)", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' And idProducto = '" & rst.Fields("idProducto").Value & "' ")
                            strMO = IIf(traerCampo("Docventas", "idMoneda", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' ") = "PEN", "S/.", "U$$")
                            dblPorcDsc = traerCampo("DocventasDet", "PorDcto", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' And idProducto = '" & rst.Fields("idProducto").Value & "' ")
                            dblMtoTotal = traerCampo("DocventasDet", "Round(TotalVVNeto,2)", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' And idProducto = '" & rst.Fields("idProducto").Value & "' ")
                        
                            If traerCampo("DocventasDet", "PorDcto", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento  ='" & strTD & "' And idSucursal = '" & glsSucursal & "' And idProducto = '" & rst.Fields("idProducto").Value & "' ") > 0 Then
                                strGlosaDscto = "Al monto " & strMO & "" & strMtoSinDsc & " aplicar descuento del " & dblPorcDsc & "%, total " & strMO & "" & dblMtoTotal & " "
                            Else
                                strGlosaDscto = ""
                            End If
                        End If
                        
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
                        
                        ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila, ncol, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        If StrMsgError <> "" Then GoTo ERR
                        
                        xfila = xfila + 4
                        filas_detalle = 0
                        Long_Acumu = Long_Acumu + Len(s)
                        
                        If Len(Linea_l) <= Len(s) Then
                            nvarfila = nvarfila + 4
                            
                        Else
                            Do While Long_total > Long_Acumu
                               '--- Añadirle el margen izquierdo
                                nvarfila = nvarfila + 4
                                sCabecera = sCabecera & Space$(margenIzq) & s & vbCrLf
                                s = LoopPropperWrap()
                                xMemo = s
                                ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila, ncol, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                                xfila = xfila + 4
                                filas_detalle = filas_detalle + 4
                                Long_Acumu = Long_Acumu + Len(s) + Len(vbCrLf)
                                If Long_Acumu = Long_Acumu + Len(s) Then
                                    contadorImp = contadorImp + 1
                                End If
                                If contadorImp = 10 Then Exit Do
                            Loop
                            nvarfila = nvarfila + IIf(rucEmpresa = "20430471750", 0, 4)
                        End If
                 
                    
                    ElseIf rst.Fields(i).Name = "idMarca" Then
                        If rucEmpresa = "20305948277" Then
                            GlsMarca = Trim("" & traerCampo("Marcas", "GlsMarca", "idmarca", rst.Fields(i), True))
                            ImprimeXY GlsMarca, strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                            If StrMsgError <> "" Then GoTo ERR
                            
                        Else
                            ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                            If StrMsgError <> "" Then GoTo ERR
                        End If
                        
                    Else
                        If rucEmpresa = "20305948277" Then
                            If rst.Fields(i).Name = "idProducto" Then
                                If Len(Trim(traerCampo("productosclientes", "Codigo", "idproducto", Trim("" & rst.Fields(i)), True, " idclIente = '" & strCodigoCliente & "' "))) = 0 Then
                                    If Len(Trim("" & traerCampo("Productos", "CodigoRapido", "idproducto", rst.Fields(i), True))) = 0 Then
                                        ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                                        If StrMsgError <> "" Then GoTo ERR
                                    Else
                                        ImprimeXY right(Trim("" & traerCampo("Productos", "CodigoRapido", "idproducto", rst.Fields(i), True)), 6) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                                        If StrMsgError <> "" Then GoTo ERR
                                    End If
                                Else
                                    ImprimeXY Trim(traerCampo("productosclientes", "Codigo", "idproducto", Trim("" & rst.Fields(i)), True, " idclIente = '" & strCodigoCliente & "' ")) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                                    If StrMsgError <> "" Then GoTo ERR
                                End If
                            Else
                                ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                            End If
                        Else
                            ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, StrMsgError
                            If StrMsgError <> "" Then GoTo ERR
                        End If
                    End If
                 End If
                rstObj.Close
            Next
            
            If Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "GENERA_VALE_FORMULA", True) & "") = "S" Then
                
                CSqlC = "Select B.IdFabricante,B.GlsProducto,A.Cantidad " & _
                        "From DocVentasDetFormula A " & _
                        "Inner Join Productos B " & _
                            "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                        "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdDocumento = '" & strTD & "' " & _
                        "And A.IdSerie = '" & strSerie & "' And A.IdDocVentas = '" & strNumDoc & "'"
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
            wsum = wsum + IIf(intScale = 6, IIf(rucEmpresa = "20504973990" Or rucEmpresa = "20305948277" Or rucEmpresa = "20257354041" Or rucEmpresa = "20430471750" Or rucEmpresa = "20511137137" Or rucEmpresa = "20509571792" Or rucEmpresa = "20542001969" Or rucEmpresa = "20544632192" Or rucEmpresa = "20388197804", 0, 4), 1) + numEntreLineasAdicional + nvarfila
        Loop
        rst.Close
        
        If rucEmpresa = "20504973990" Then
            TotFlete = 0#
            csql = "SELECT TotalVVNeto FROM docventasdet d, productos p " & _
                   "WHERE d.idProducto = p.idProducto AND d.idEmpresa = p.idEmpresa " & _
                   "AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' " & _
                   "AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "' AND p.idNivel = '11080037' "
            rst.Open csql, Cn, adOpenStatic, adLockReadOnly
            
            Do While Not rst.EOF
                TotFlete = Format(Val(rst.Fields("TotalVVNeto") & ""), "0.00")
                
                If strTD = "01" Or strTD = "03" Then
                    ImprimeXY Format(Val(rst.Fields("TotalVVNeto") & ""), "0.00"), "N", 10, 160, 70, 2, 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                End If
                rst.MoveNext
            Loop
            rst.Close
            
            csql = "SELECT TotalValorVenta FROM docventas d " & _
                   "WHERE d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' " & _
                   "AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "' "
            rst.Open csql, Cn, adOpenStatic, adLockReadOnly
            
            Do While Not rst.EOF
            
                If strTD = "01" Or strTD = "03" Then
                    ImprimeXY Format(Val(rst.Fields("TotalValorVenta") & "") - Val(TotFlete & ""), "0.00"), "N", 10, 160, 10, 2, 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                End If
                
                rst.MoveNext
            Loop
            rst.Close
            
        End If
    End If
        
    '------------------------------------------------------------------------------------------------
    If strImprimeTicket = "S" Then
        wsum = wsum + IIf(intScale = 6, 4, 1)
    Else
        wsum = 0
    End If
    
    boolEst = False
    strCampos = ""
    If intRegistros = 0 Then
        csql = "SELECT GlsCampo FROM objdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '999' ORDER BY IMPY,IMPX "
    Else
        csql = "SELECT GlsCampo FROM objdocventas  WHERE idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' and trim(GlsCampo) <> '' and indImprime = 1 and idDocumento = '" & strTD & "' and idSerie = '" & strSerie & "' ORDER BY IMPY,IMPX "
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
                                ImprimeXY "Total :  S/. ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
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
                            
                            If rst.Fields("IdMoneda") = "PEN" Then
                                ImprimeXY "Efectivo :  S/. ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
                            Else
                                ImprimeXY "Efectivo : US$. ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
                            End If
                            
                            ImprimeXY Format(IIf(dblMtoTotEnt = "", "0.00", dblMtoTotEnt), "0.00") & "", "N", "10", intY + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            wsum = wsum + IIf(intScale = 6, 4, 1)
                            
                            If rst.Fields("IdMoneda") = "PEN" Then
                                ImprimeXY "Vuelto :  S/. ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
                            Else
                                ImprimeXY "Vuelto : US$. ", "T", 20, intY + wsum, 1, 2, 0, StrMsgError
                            End If
                            
                            dblMtoTotVuelto = Trim("" & traerCampo("movcajasdet m INNER JOIN monedas o ON m.idMoneda = o.idMoneda", "Sum(m.ValMonto)", "m.idDocumento", strTD, True, "m.idDocVentas = '" & strNumDoc & "' AND m.idSerie = '" & strSerie & "' AND m.idTipoMovCaja = '99990003' And m.idSucursal = '" & glsSucursal & "'"))
                            ImprimeXY Format(IIf(dblMtoTotVuelto = "", "0.00", dblMtoTotVuelto), "0.00") & "", "N", "10", intY + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            wsum = wsum + IIf(intScale = 6, 4, 1)
                            
                            boolEst = True
                        End If
                    Else
                        If strTD = "86" Or strTD = "07" Then
                        If strTD = "86" And rucEmpresa = "20542001969" Then ' Para MM
                            MMGuiaRemison rstObj, rst, i, StrMsgError
                            If StrMsgError <> "" Then GoTo ERR
                        Else
                            ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                        End If
                        
                        Else
                            If rucEmpresa = "20119041148" Or rucEmpresa = "20119040842" Or rucEmpresa = "20530611681" Then
                                
                                SinchiImpSp rst.Fields(i).Name, rstObj.Fields("valor"), rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), rstObj.Fields("Decimales"), wsum, iddocumentoRef, StrMsgError, strNumDoc, strSerie, strTD, rst.Fields("idMoneda").Value
                                If StrMsgError <> "" Then GoTo ERR
                                                                    
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
                            
                            ElseIf rucEmpresa = "20509571792" And strTD = "01" And rst.Fields(i).Name = "GlsDocReferencia" Then
                                cSqlRef = "Select SerieDocReferencia,NumDocReferencia " & _
                                            "From DocReferencia " & _
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
                                                ImprimeXY Val("" & .Fields("NumDocReferencia")) & " / ", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), nFilRef, rstObj.Fields("impX") + (NVueltas * 11), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                            Else
                                                .MovePrevious
                                                ImprimeXY Val("" & .Fields("NumDocReferencia")), rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), nFilRef, rstObj.Fields("impX") + (NVueltas * 11), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
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
                                If StrMsgError <> "" Then GoTo ERR
                            ElseIf rucEmpresa = "20504973990" And rst.Fields(i).Name = "TotalPrecioVenta" And traerCampo("Docventas", "indVtaGratuita", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento='" & strTD & "'") = "1" Then   'Solvet Evalua Campo para las muestras gratuitas
                                ImprimeXY "0.00" & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            ElseIf rucEmpresa = "20504973990" And rst.Fields(i).Name = "totalLetras" And traerCampo("Docventas", "indVtaGratuita", "idDocventas", strNumDoc, True, "idSerie = '" & strSerie & "' And idDocumento='" & strTD & "'") = "1" Then  'Solvet Evalua Campo para las muestras gratuitas
                                ImprimeXY "SON: 0.00 " & IIf(rst.Fields("idMoneda").Value = "PEN", "NUEVOS SOLES", "DOLARES AMERICANOS") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            Else
                                ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + wsum, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                            End If
                        End If
                    End If
                    
                    If StrMsgError <> "" Then GoTo ERR
                End If
                rstObj.Close
            Next
        End If
        rst.Close
    End If
    
    '-----------------------------------------------------------------------------------------------------------------------------------
    '--- IMPRIME ETIQUETAS FINAL
    '-----------------------------------------------------------------------------------------------------------------------------------
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
         strTexto = ""
        If (rst.Fields("indRUCCliente") & "") = 0 And (rst.Fields("indRazonSocial") & "") = 0 And (rst.Fields("indIGVTotal") & "") = 0 And (rst.Fields("indHora") & "") = 0 And (rst.Fields("indVendedor") & "") = 0 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                End If
            End If
        End If
                        
        If (rst.Fields("indDirSucursal") & "") = 1 Then strTexto = strTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        '08001 FACTURA -- '08002 BOLETA
        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strRUCCliente
                End If
            Else
                strTexto = rst.Fields("Etiqueta") & ""
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsCliente  '2
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsCliente  '2
                End If
            End If
        End If
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                     strTexto = strTexto + strTotalIGV
                End If
            Else
                strTexto = rst.Fields("Etiqueta") & ""
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        
        If (rst.Fields("indHora") & "") = 1 Then
            strTexto = rst.Fields("Etiqueta") & ""
            strTexto = strTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        End If
                           
        If (rst.Fields("indVendedor") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsVendedorCampo
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsVendedorCampo
                End If
            End If
        End If
        
        If (rst.Fields("indDirecCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strDirecCliente
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strDirecCliente
                End If
            End If
        End If
        
        If Len(Trim(strTexto)) > 0 Then
            If strImprimeTicket = "S" And strTD = "12" Then
                ImprimeXY strTexto, "T", Len(strTexto), intY + wsum, rst.Fields("impX"), 0, 0, StrMsgError
                wsum = wsum + IIf(intScale = 6, 4, 1)
            Else
                ImprimeXY strTexto, "T", Len(strTexto), rst.Fields("impY"), rst.Fields("impX"), 0, 0, StrMsgError
                wsum = wsum + IIf(intScale = 6, 4, 1)
            End If
        End If
        If StrMsgError <> "" Then GoTo ERR
        
        rst.MoveNext
    Loop
    rst.Close
    '------------------------------------------------------------------------------------------------
        
    Set rst = Nothing
    Set rstObj = Nothing
                               
    If (rucEmpresa = "20266578572" Or rucEmpresa = "20552185286" Or rucEmpresa = "20552184981") Then
        imprimeDocVentas_2 strTD, strNumDoc, strSerie, StrMsgError
        If Len(Trim(StrMsgError)) > 0 Then GoTo ERR
    Else
        If strTD = "12" Then
            Printer.Print Chr$(149)
            'Printer.Print ""
            'Printer.Print ""
        End If
        Printer.EndDoc
    End If
        
    Exit Sub
    
ERR:
    If StrMsgError = "" Then StrMsgError = ERR.Description
    Printer.KillDoc
End Sub
Public Sub imprimeGuiaAbierta(strTD As String, strNumDoc As String, strSerie As String, ByRef StrMsgError As String)
On Error GoTo ERR
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
Dim strTipoTicket       As String
Dim strRUCCliente       As String
Dim strGlsCliente       As String
Dim strTotalIGV         As String
Dim strGlsVendedorCampo As String
Dim strDirecCliente     As String
Dim strMotivoTraslado   As String
Dim strCodigoCliente    As String
Dim cselect             As String
Dim nfontletra          As String
Dim numEntreLineasAdicional As Integer
Dim intRegistros            As Integer
Dim rsd                     As New ADODB.Recordset
Dim strTexto                As String
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
           
    rsd.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsd.EOF Then
        strTipoTicket = "" & rsd.Fields("idTipoTicket")
        strRUCCliente = "" & rsd.Fields("RUCCliente")
        strGlsCliente = "" & rsd.Fields("GlsCliente")
        strTotalIGV = CStr(Format("" & rsd.Fields("TotalIGVVenta"), "###,##0.00"))
        strGlsVendedorCampo = "" & rsd.Fields("GlsVendedorCampo")
        strDirecCliente = "" & rsd.Fields("dirCliente")
        strMotivoTraslado = "" & rsd.Fields("idMotivoTraslado")
        strCodigoCliente = "" & rsd.Fields("idPerCliente")
    End If
    
    If Len(Trim(strMotivoTraslado)) > 0 Then
        csql = "SELECT GlsMotivoTraslado FROM motivostraslados WHERE idMotivoTraslado = '" & strMotivoTraslado & "'"
        If rsd.State = adStateOpen Then rsd.Close
        rsd.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rsd.EOF Then
            strGlsMotivoTraslado = Trim(rsd.Fields("GlsMotivoTraslado") & "")
        End If
    End If
    rsd.Close: Set rsd = Nothing
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
        strTexto = rst.Fields("Etiqueta") & ""
        
        If (rst.Fields("indDirSucursal") & "") = 1 Then strTexto = strTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)

        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                     strTexto = strTexto + strRUCCliente
                End If
            Else
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                     strTexto = strTexto + strGlsCliente  '1
                End If
            Else
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                     strTexto = strTexto + strTotalIGV
                End If
            Else
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indHora") & "") = 1 Then strTexto = strTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        '-------------------------------------------------------------------
        '------ BUSCA SI ES SUCURSAL, EXTRAE RAZON SOCIAL Y RUC DE LA EMPRESA
        If (rst.Fields("indDestinatarioGuia") & "") = 1 Then
            cselect = "SELECT idSucursal FROM sucursales WHERE idSucursal = '" & strCodigoCliente & "' AND idEmpresa = '" & glsEmpresa & "'"
            If rsd.State = adStateOpen Then rsd.Close
            rsd.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rsd.EOF Then
                rsd.Close: Set rsd = Nothing
                
                cselect = "SELECT GlsEmpresa, RUC FROM empresas WHERE idEmpresa = '" & glsEmpresa & "'"
                If rsd.State = adStateOpen Then rsd.Close
                rsd.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rsd.EOF Then
                    strTexto = strTexto & Trim(rsd.Fields("GlsEmpresa") & "") & Space(5) & "RUC: " & Trim(rsd.Fields("RUC") & "")
                End If
                'rsd.Close: Set rsd = Nothing
            Else
                strTexto = strTexto + strGlsCliente & Space(5) & "RUC/DNI: " & strRUCCliente
            End If
            rsd.Close: Set rsd = Nothing
        End If
        '-------------------------------------------------------------------
        If (rst.Fields("indMotivoTrasladoGuia") & "") = 1 Then
            strTexto = strTexto & strGlsMotivoTraslado
        End If
        '-------------------------------------------------------------------
        ImprimeXY strTexto, "T", Len(strTexto & ""), rst.Fields("impY"), rst.Fields("impX"), 0, 0, StrMsgError
        If StrMsgError <> "" Then GoTo ERR
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
                    If StrMsgError <> "" Then GoTo ERR
                    
                    'IMPRIME EL MES
                    ImprimeXY strArregloMes(Val(Month(rstObj.Fields("valor")))) & "", "T", 10, rstObj.Fields("impY"), rstObj.Fields("impX") + 20, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                    
                    'IMPRIME EL AÑO
                    ImprimeXY right(CStr(Year(rstObj.Fields("valor"))), 1) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 70, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                Else
                    ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
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
                If StrMsgError <> "" Then GoTo ERR
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
                                ImprimeXY "Total :  S/. ", "T", 20, intY + wsum, 1, 0, 0, StrMsgError
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
                    If StrMsgError <> "" Then GoTo ERR
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
         strTexto = ""
        If (rst.Fields("indRUCCliente") & "") = 0 And (rst.Fields("indRazonSocial") & "") = 0 And (rst.Fields("indIGVTotal") & "") = 0 And (rst.Fields("indHora") & "") = 0 And (rst.Fields("indVendedor") & "") = 0 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                End If
            End If
        End If
                        
        If (rst.Fields("indDirSucursal") & "") = 1 Then strTexto = strTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        '08001 FACTURA
        '08002 BOLETA
        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strRUCCliente
                End If
            Else
                strTexto = rst.Fields("Etiqueta") & ""
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsCliente  '2
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsCliente  '2
                End If
            End If
        End If
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                     strTexto = strTexto + strTotalIGV
                End If
            Else
                strTexto = rst.Fields("Etiqueta") & ""
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indHora") & "") = 1 Then
            strTexto = rst.Fields("Etiqueta") & ""
            strTexto = strTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        End If
                           
        If (rst.Fields("indVendedor") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsVendedorCampo
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsVendedorCampo
                End If
            End If
        End If
        
        If (rst.Fields("indDirecCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strDirecCliente
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strDirecCliente
                End If
            End If
        End If
        
        If Len(Trim(strTexto)) > 0 Then
            ImprimeXY strTexto, "T", Len(strTexto), intY + wsum, rst.Fields("impX"), 0, 0, StrMsgError
            wsum = wsum + IIf(intScale = 6, 4, 1)
        Else
        
        End If
        If StrMsgError <> "" Then GoTo ERR
        
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
    
ERR:
    If StrMsgError = "" Then StrMsgError = ERR.Description
    Printer.KillDoc
End Sub

Private Sub ImprimeXY(varData As Variant, strTipoDato As String, intTamanoCampo As Integer, intFila As Integer, intColu As Integer, intDecimales As Integer, intFilas As Integer, ByRef StrMsgError As String)
On Error GoTo ERR
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
    
ERR:
    If StrMsgError = "" Then StrMsgError = ERR.Description
End Sub

Public Sub imprimeReciboCaja(strNumMovCajaDet As String, ByRef StrMsgError As String)
On Error GoTo ERR
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
                    If StrMsgError <> "" Then GoTo ERR
            End If
            rstObj.Close
        Next
    End If
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Set rstObj = Nothing
                           
    Printer.EndDoc
    
    Exit Sub

ERR:
    If StrMsgError = "" Then StrMsgError = ERR.Description
    Printer.KillDoc
End Sub

Public Sub ImprimeCodigoBarra(ByVal indTipo As Integer, ByVal codproducto As String, ByVal numVale As String, ByRef StrMsgError As String, Optional ByVal dblCantidad As Double = 0)
On Error GoTo ERR
Dim objPrinter As New PrinterAPI.clsPrinter
Dim rsp As New ADODB.Recordset
Dim StrCodBarra As String
Dim BlnFoundPrinter As Boolean
Dim BlnFoundData As Boolean
Dim intPar As Long
Dim i As Integer
Dim intParTotal As Long
Dim indParTotal As Boolean
Dim strPrecio As String
Dim strTalla As String
Dim intCantidad As Integer

    If (objPrinter.SetPrinter("Generica / Solo Texto") = False) Then
        If (objPrinter.SetPrinter("Generic / Text Only") = False) Then
            MsgBox "No se Encuentra instalada la Impresora " & NombreImpresora_sp & "o " & NombreImpresora_us, vbInformation
            Exit Sub
        End If
    End If

    StrGlsEmpresa = traerCampo("empresas", "GlsEmpresa", "idEmpresa", glsEmpresa, False)
    If indTipo = 0 Then
        csql = "SELECT v.idProducto,p.GlsProducto,t.GlsTallaPeso,l.PVUnit,v.Cantidad " & _
               "FROM valesdet v,productos p, tallapeso t, preciosventa l " & _
               "WHERE v.idEmpresa = p.idEmpresa " & _
                 "AND v.idProducto = p.idProducto " & _
                 "AND p.idEmpresa = t.idEmpresa " & _
                 "AND p.idTallaPeso = t.idTallaPeso " & _
                 "AND p.idEmpresa = l.idEmpresa " & _
                 "AND p.idProducto = l.idProducto " & _
                 "AND p.idUMCompra = l.idUM " & _
                 "AND v.idValesCab = '" & numVale & "' AND p.idProducto = '" & codproducto & "' AND l.idLista = '" & glsListaVentas & "'"
        intCantidadTotal = 1

    ElseIf indTipo = 1 Then
        csql = "SELECT v.idProducto,p.GlsProducto,t.GlsTallaPeso,l.PVUnit,v.Cantidad " & _
               "FROM valesdet v,productos p, tallapeso t, preciosventa l " & _
               "WHERE v.idEmpresa = p.idEmpresa " & _
                 "AND v.idProducto = p.idProducto " & _
                 "AND p.idEmpresa = t.idEmpresa " & _
                 "AND p.idTallaPeso = t.idTallaPeso " & _
                 "AND p.idEmpresa = l.idEmpresa " & _
                 "AND p.idProducto = l.idProducto " & _
                 "AND p.idUMCompra = l.idUM " & _
                 "AND v.idValesCab = '" & numVale & "' AND l.idLista = '" & glsListaVentas & "'"
        intCantidadTotal = Val("" & traerCampo("valesdet", "SUM(Cantidad)", "idSucursal", glsSucursal, True, " idValesCab = '" & numVale & "'"))

    ElseIf indTipo = 2 Then
        csql = "SELECT p.idProducto,p.GlsProducto,t.GlsTallaPeso,l.PVUnit," & CStr(dblCantidad) & " AS Cantidad " & _
               "FROM productos p, tallapeso t, preciosventa l " & _
               "WHERE p.idEmpresa = t.idEmpresa " & _
                 "AND p.idTallaPeso = t.idTallaPeso " & _
                 "AND p.idEmpresa = l.idEmpresa " & _
                 "AND p.idProducto = l.idProducto " & _
                 "AND p.idUMCompra = l.idUM " & _
                 "AND p.idProducto = '" & codproducto & "' AND l.idLista = '" & glsListaVentas & "'"
        intCantidadTotal = 1

    End If
    rsp.Open csql, Cn, adOpenForwardOnly, adLockReadOnly

    intPar = 0
    If indTipo = 0 Or indTipo = 2 Then
        If Not rsp.EOF Then
            intParTotal = intCantidadTotal * Val("" & rsp.Fields("Cantidad"))
        Else
            StrMsgError = "El producto no tiene precio"
            GoTo ERR
        End If
    Else
        intParTotal = intCantidadTotal
    End If

    indParTotal = True
    If intParTotal Mod 2 Then
        indParTotal = False
    End If

    Do While (Not rsp.EOF)
        strPrecio = Format(rsp.Fields("PVUnit").Value, "##,##0.00")
        strTalla = Trim$(rsp.Fields("GlsTallaPeso").Value)
        intCantidad = Val("" & rsp.Fields("Cantidad"))

        For i = 1 To intCantidad
            intPar = intPar + 1
            objPrinter.PrintDataLn Chr$(2) & "L"
            objPrinter.PrintDataLn "A2"
            objPrinter.PrintDataLn "D11"
            objPrinter.PrintDataLn "z"
            objPrinter.PrintDataLn "PN"
            objPrinter.PrintDataLn "H10"

            StrCodBarra = Trim$(rsp.Fields("idProducto").Value)

            If intPar Mod 2 Then
                objPrinter.PrintDataLn "191100300610140" & strPrecio 'PRECIO -30
                objPrinter.PrintDataLn "191100100280010" & strTalla 'TALLA

                objPrinter.PrintDataLn "191100100610070" & StrGlsEmpresa
                objPrinter.PrintDataLn "191100100500010" & left(Trim$(rsp.Fields("GlsProducto").Value), 38)
                objPrinter.PrintDataLn "191100300280133" & StrCodBarra
                objPrinter.PrintDataLn "1e2201600010016B" & StrCodBarra

                objPrinter.PrintDataLn "^01"     ' Numero de Copias
                objPrinter.PrintDataLn "Q0001"   ' Numero de Etiquetas

            Else
                objPrinter.PrintDataLn "191100300610350" & strPrecio 'PRECIO
                objPrinter.PrintDataLn "191100100280220" & strTalla 'TALLA

                objPrinter.PrintDataLn "191100100610280" & StrGlsEmpresa
                objPrinter.PrintDataLn "191100100500220" & left(Trim$(rsp.Fields("GlsProducto").Value), 38)
                objPrinter.PrintDataLn "191100300280342" & StrCodBarra
                objPrinter.PrintDataLn "1e2201600010228B" & StrCodBarra

                objPrinter.PrintDataLn "^01"     ' Numero de Copias
                objPrinter.PrintDataLn "Q0001"   ' Numero de Etiquetas
                objPrinter.PrintDataLn "E"       ' Enviar la Impresion

            End If

            If indParTotal = False And intPar = intParTotal Then
                objPrinter.PrintDataLn "E"       ' Enviar la Impresion
            End If
            BlnFoundData = True
        Next
        rsp.MoveNext
    Loop
    rsp.Close: Set rsp = Nothing

    If (BlnFoundData) Then
        Printer.EndDoc
    Else
        StrMsgError = "No se Realizó ninguna Impresión, No hay Datos"
        GoTo ERR
    End If

    Exit Sub
    
ERR:
    If rsp.State = 1 Then rsp.Close: Set rsp = Nothing
    If StrMsgError = "" Then StrMsgError = ERR.Description
End Sub

Public Sub imprimirVale(ByVal strNumVale As String)
On Error GoTo ERR
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
    If StrMsgError <> "" Then GoTo ERR

    If Not rsReporte.EOF And Not rsReporte.BOF Then
        reporte.Database.SetDataSource rsReporte, 3

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
    
ERR:
    Screen.MousePointer = 0
    If StrMsgError = "" Then StrMsgError = ERR.Description
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
On Error GoTo ERR
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
Dim strTipoTicket       As String
Dim strRUCCliente       As String
Dim strGlsCliente       As String
Dim strTotalIGV         As String
Dim strGlsVendedorCampo As String
Dim strDirecCliente     As String
Dim strMotivoTraslado   As String
Dim strCodigoCliente    As String
Dim cselect             As String
Dim nfontletra          As String
Dim numEntreLineasAdicional As Integer
Dim intRegistros As Integer
Dim indventasterceros As String
Dim indterceros As Boolean
Dim nSumFilasEital      As String
Dim NTamanoLetra        As String
Dim rsd                 As New ADODB.Recordset
Dim strTexto            As String
Dim nfiladetalle        As Integer, nvarfila As Integer, nvarfilatotal As Integer
Dim strTipoDato         As String
Dim intLong             As Integer
Dim intX                As Integer
Dim intY                As Integer
Dim intDec              As Integer
Dim rucEmpresa          As String
Dim iddocumentoRef      As String

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
           
    rsd.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsd.EOF Then
        strTipoTicket = "" & rsd.Fields("idTipoTicket")
        strRUCCliente = "" & rsd.Fields("RUCCliente")
        strGlsCliente = "" & rsd.Fields("GlsCliente")
        strTotalIGV = CStr(Format("" & rsd.Fields("TotalIGVVenta"), "###,##0.00"))
        strGlsVendedorCampo = "" & rsd.Fields("GlsVendedorCampo")
        strDirecCliente = "" & rsd.Fields("dirCliente")
        strMotivoTraslado = "" & rsd.Fields("idMotivoTraslado")
        strCodigoCliente = "" & rsd.Fields("idPerCliente")
    End If
    
    If Len(Trim(strMotivoTraslado)) > 0 Then
        csql = "SELECT GlsMotivoTraslado FROM motivostraslados WHERE idMotivoTraslado = '" & strMotivoTraslado & "'"
        If rsd.State = adStateOpen Then rsd.Close
        rsd.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rsd.EOF Then
            strGlsMotivoTraslado = Trim(rsd.Fields("GlsMotivoTraslado") & "")
        End If
    End If
    rsd.Close: Set rsd = Nothing
    
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
        strTexto = rst.Fields("Etiqueta") & ""
        If (rst.Fields("indDirSucursal") & "") = 1 Then strTexto = strTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        '08001 FACTURA
        '08002 BOLETA
        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                     strTexto = strTexto + strRUCCliente
                End If
            Else
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                     strTexto = strTexto + strGlsCliente  '1
                End If
            Else
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                     strTexto = strTexto + strTotalIGV
                End If
            Else
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indHora") & "") = 1 Then strTexto = strTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        
        '-------------------------------------------------------------------
        '------ BUSCA SI ES SUCURSAL, EXTRAE RAZON SOCIAL Y RUC DE LA EMPRESA
        If (rst.Fields("indDestinatarioGuia") & "") = 1 Then
            cselect = "SELECT idSucursal FROM sucursales WHERE idSucursal = '" & strCodigoCliente & "' AND idEmpresa = '" & glsEmpresa & "'"
            If rsd.State = adStateOpen Then rsd.Close
            rsd.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rsd.EOF Then
                rsd.Close: Set rsd = Nothing
                cselect = "SELECT GlsEmpresa, RUC FROM empresas WHERE idEmpresa = '" & glsEmpresa & "'"
                If rsd.State = adStateOpen Then rsd.Close
                rsd.Open cselect, Cn, adOpenForwardOnly, adLockReadOnly
                If Not rsd.EOF Then
                    strTexto = strTexto & Trim(rsd.Fields("GlsEmpresa") & "") & Space(5) & "RUC: " & Trim(rsd.Fields("RUC") & "")
                End If
            Else
                strTexto = strTexto + strGlsCliente & Space(5) & "RUC/DNI: " & strRUCCliente
            End If
            rsd.Close: Set rsd = Nothing
        End If
        '-------------------------------------------------------------------
        If (rst.Fields("indMotivoTrasladoGuia") & "") = 1 Then
            strTexto = strTexto & strGlsMotivoTraslado
        End If
        '-------------------------------------------------------------------
        ImprimeXY strTexto, "T", Len(strTexto & ""), rst.Fields("impY") + nSumFilasEital, rst.Fields("impX"), 0, 0, StrMsgError
        If StrMsgError <> "" Then GoTo ERR
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
        indventasterceros = traerCampo("clientes", "indventasterceros", "idcliente", strCodigoCliente, True)
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
                    If StrMsgError <> "" Then GoTo ERR
                    
                    'IMPRIME EL MES
                    ImprimeXY strArregloMes(Val(Month(rstObj.Fields("valor")))) & "", "T", 10, rstObj.Fields("impY") + nSumFilasEital, rstObj.Fields("impX") + 20, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                    
                    'IMPRIME EL AÑO
                    ImprimeXY right((Year(rstObj.Fields("valor"))), 2) & "", "T", 4, rstObj.Fields("impY") + nSumFilasEital, rstObj.Fields("impX") + 60, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
                    
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
                        If StrMsgError <> "" Then GoTo ERR
                        
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
                                If StrMsgError <> "" Then GoTo ERR
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
                            If StrMsgError <> "" Then GoTo ERR
                        End If
                    Else
                        If rst.Fields(i).Name = "vtercerosCliente" Or rst.Fields(i).Name = "vterceroRuc" Or rst.Fields(i).Name = "vtercerosdireccion" Then
                        Else
                             ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY") + nSumFilasEital, rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
                             If StrMsgError <> "" Then GoTo ERR
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
                        If StrMsgError <> "" Then GoTo ERR
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
                    If StrMsgError <> "" Then GoTo ERR
                Else
                    ImprimeXY Format(Val(rst.Fields("TotalVVNeto") & ""), "0.00"), "N", 10, 95, 97, 2, 0, StrMsgError
                    If StrMsgError <> "" Then GoTo ERR
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
                        If StrMsgError <> "" Then GoTo ERR
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
                        nfiladetalle = 0: nvarfila = 0: nvarfilatotal = 0
                        xfila = Val(rstObj.Fields("impY")) + wsum
                        ncol = Val(rstObj.Fields("impX"))
                
                        Linea_l = rst.Fields(i) & ""
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
                        If StrMsgError <> "" Then GoTo ERR
                        
                        xfila = xfila + 4
                        nvarfila = nvarfila + 4
                        filas_detalle = 0
                        
                        If Len(Linea_l) <= Len(s) Then
                        Else
                            Do While Len(s) > 0
                                '--- Añadirle el margen izquierdo
                                sCabecera = sCabecera & Space$(margenIzq) & s & vbCrLf
                                s = LoopPropperWrap()
                                xMemo = s
                                ImprimeXY s, rstObj.Fields("tipoDato"), Val("" & rstObj.Fields("impLongitud")), xfila + nSumFilasEital, ncol, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
                                If StrMsgError <> "" Then GoTo ERR
                                xfila = xfila + 4
                                nvarfila = nvarfila + 4
                                filas_detalle = filas_detalle + 4
                            Loop
                        End If
                    Else
                        ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum + nSumFilasEital, intX, intDec, 0, StrMsgError
                        If StrMsgError <> "" Then GoTo ERR
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
                                ImprimeXY "Total :  S/. ", "T", 20, intY + wsum + nSumFilasEital, 1, 0, 0, StrMsgError
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
                    If StrMsgError <> "" Then GoTo ERR
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
        strTexto = ""
        If (rst.Fields("indRUCCliente") & "") = 0 And (rst.Fields("indRazonSocial") & "") = 0 And (rst.Fields("indIGVTotal") & "") = 0 And (rst.Fields("indHora") & "") = 0 And (rst.Fields("indVendedor") & "") = 0 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                End If
            End If
        End If
                        
        If (rst.Fields("indDirSucursal") & "") = 1 Then strTexto = strTexto + traerDireccionSucursal
        If (rst.Fields("indSerieEtiquetera") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
        If (rst.Fields("indUsuario") & "") = 1 Then strTexto = strTexto + traerCampo("usuarios", "varUsuario", "idusuario", glsUser, True)
        
        '08001 FACTURA
        '08002 BOLETA
        If (rst.Fields("indRUCCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strRUCCliente
                End If
            Else
                strTexto = rst.Fields("Etiqueta") & ""
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indRazonSocial") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsCliente  '2
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsCliente  '2
                End If
            End If
        End If
                    
        If (rst.Fields("indIGVTotal") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                     strTexto = strTexto + strTotalIGV
                End If
            Else
                strTexto = rst.Fields("Etiqueta") & ""
                strTexto = strTexto + traerCampo("usuarios", "serieetiquetera", "idusuario", glsUser, True)
            End If
        End If
        
        If (rst.Fields("indHora") & "") = 1 Then
            strTexto = rst.Fields("Etiqueta") & ""
            strTexto = strTexto + CStr(Format(getFechaHoraSistema, "HH:mm:SS"))
        End If
                           
        If (rst.Fields("indVendedor") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsVendedorCampo
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strGlsVendedorCampo
                End If
            End If
        End If
        
        If (rst.Fields("indDirecCliente") & "") = 1 Then
            If (rst.Fields("indSoloTicketFactura") & "") = 1 Then
                If strTipoTicket = "08001" Then 'FACTURA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strDirecCliente
                End If
            Else
                If strTipoTicket = "08002" Then 'BOLETA
                    strTexto = rst.Fields("Etiqueta") & ""
                    strTexto = strTexto + strDirecCliente
                End If
            End If
        End If
        
        If Len(Trim(strTexto)) > 0 Then
            ImprimeXY strTexto, "T", Len(strTexto), Val(rst.Fields("impY").Value) + nSumFilasEital, rst.Fields("impX"), 0, 0, StrMsgError
            wsum = wsum + IIf(intScale = 6, 4, 1)
        End If
        If StrMsgError <> "" Then GoTo ERR
        
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

ERR:
    If StrMsgError = "" Then StrMsgError = ERR.Description
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
            ImprimeXY "SON: 0.00 " & IIf(idMoneda = "PEN", "NUEVOS SOLES", "DOLARES AMERICANOS") & "", ObjTipoDato, ObjLong, ObjImpy + wsum, ObjImpx, Val("" & ObjImpDecim), 0, StrMsgError
            
        Else
            ImprimeXY ObjValor & "", ObjTipoDato, ObjLong, ObjImpy + wsum, ObjImpx, Val("" & ObjImpDecim), 0, StrMsgError
        End If
        
End Sub

Private Sub MMGuiaRemison(rstObj As ADODB.Recordset, rst As ADODB.Recordset, i, StrMsgError As String)
On Error GoTo ERR
Dim strDNITran As String
Dim strGlsTran As String
Dim strTIPDOC As String
 

    If rst.Fields(i).Name = "FecIniTraslado" Then
        'IMPRIME EL DIA
        ImprimeXY Format(Day(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
        If StrMsgError <> "" Then GoTo ERR

        'IMPRIME EL MES
        ImprimeXY Format(Month(rstObj.Fields("valor")), "00") & "", "T", 2, rstObj.Fields("impY"), rstObj.Fields("impX") + 15, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
        If StrMsgError <> "" Then GoTo ERR

        'IMPRIME EL AÑO
        ImprimeXY right((Year(rstObj.Fields("valor"))), 4) & "", "T", 4, rstObj.Fields("impY"), rstObj.Fields("impX") + 30, Val("" & rstObj.Fields("Decimales")), 0, StrMsgError
        If StrMsgError <> "" Then GoTo ERR

    ElseIf rst.Fields(i).Name = "GlsVehiculo" Then
        ImprimeXY rstObj.Fields("valor") & " " & rst.Fields("Marca").Value & " " & rst.Fields("Placa").Value, rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
        If StrMsgError <> "" Then GoTo ERR
        
    ElseIf rst.Fields(i).Name = "idPerChofer" Then
        strDNITran = traerCampo("Personas", "RUC", "idPersona", rstObj.Fields("valor"), False)
        strGlsTran = traerCampo("Personas", "GlsPersona", "idPersona", rstObj.Fields("valor"), False)
    
        ImprimeXY strGlsTran & "", rstObj.Fields("tipoDato"), "80", "242", "17", Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
        If StrMsgError <> "" Then GoTo ERR
        
        'Imprime el RUC y/o DNI
        ImprimeXY strDNITran & "", rstObj.Fields("tipoDato"), "12", "242", "118", Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), StrMsgError
        If StrMsgError <> "" Then GoTo ERR
        
    ElseIf rst.Fields(i).Name = "GlsDocReferencia" Then
        strTIPDOC = traerCampo("Documentos", "GlsDocumento", "AbreDocumento", left(rstObj.Fields("valor"), 3), False)
        ImprimeXY strTIPDOC & "", rstObj.Fields("tipoDato"), "20", "242", "155", Val("" & rstObj.Fields("Decimales")), "1", StrMsgError
        If StrMsgError <> "" Then GoTo ERR
        
        ImprimeXY right(rstObj.Fields("valor"), 8) & "", rstObj.Fields("tipoDato"), "8", "242", "185", Val("" & rstObj.Fields("Decimales")), "1", StrMsgError
        If StrMsgError <> "" Then GoTo ERR
    End If
     
    Exit Sub
ERR:
If StrMsgError = "" Then StrMsgError = ERR.Description
End Sub

Private Sub PredeterminaImpresora(indPrinter, strTD, p As Object, StrMsgError As String)
On Error GoTo ERR
' , indPrinter, strTD, p, StrMsgError
indPrinter = False
    If strTD = "01" Then '--- FACTURA
        For Each p In Printers
           If left(UCase(p.DeviceName), 7) = "FACTURA" Then
              Set Printer = p
              indPrinter = True
              Exit For
           End If
        Next p
    ElseIf strTD = "03" Then '--- BOLETA
            For Each p In Printers
                If left(UCase(p.DeviceName), 6) = "BOLETA" Then
                    Set Printer = p
                    indPrinter = True
                    Exit For
                End If
            Next p
    ElseIf strTD = "86" Then '--- GUIA
            For Each p In Printers
                If left(UCase(p.DeviceName), 4) = "GUIA" Then
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
           If p.Port = "Ne03:" Then
           ''If p.Port = "LPT1:" Then
              Set Printer = p
              Exit For
           End If
        Next p
    End If
    
    Exit Sub
    
ERR:
    If StrMsgError = "" Then StrMsgError = ERR.Description
End Sub
