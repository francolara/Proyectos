VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function generaCorrelativoDocVentas(strTD As String, strSerie As String) As String
Dim rst As New ADODB.Recordset
Dim strCond As String
Dim csql As String

    csql = "SELECT MAX(idDocVentas) FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & strTD & "' AND idSerie = '" & strSerie & "'"
    'csql = "SELECT MAX(idDocVentas) FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idSerie = '" & strSerie & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        generaCorrelativoDocVentas = Format(Val("" & rst.Fields(0)) + 1, "00000000")
    Else
        generaCorrelativoDocVentas = "00000001"
    End If
    rst.Close: Set rst = Nothing
    
End Function


Public Sub EjecutaSQLFormDocVentas_OC(F As Form, ByVal tipoOperacion As Integer, ByRef StrMsgError As String, ByVal strTD As String, ByVal strSerie As String, ByVal g As dxDBGrid, ByVal D As dxDBGrid, ByVal indGeneraVale As Boolean)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsCorrela As New ADODB.Recordset
Dim C           As Object
Dim csql        As String
Dim strCampo    As String
Dim strTipoDato As String
Dim strCampos   As String
Dim strValores  As String
Dim strValCod   As String
Dim strCodMotivoTranslado As String
Dim strCodCliente         As String
Dim strCodVendedorCampo   As String
Dim strCod                As String
Dim strMoneda             As String
Dim strtipocambio         As String
Dim swNuevoProducto       As String
Dim valValorVenta         As Double
Dim valIGVVenta           As Double
Dim valPrecioVenta        As Double
Dim i                     As Integer
Dim StrItem               As Integer
Dim indTrans              As Boolean
Dim strcant               As Integer
Dim cantReg               As Integer
Dim rstCGP                As New ADODB.Recordset
Dim nPC                   As String
Dim strNomTbl             As String
Dim strVarAux             As String
Dim strVarCC              As String
Dim rsEli                 As New ADODB.Recordset
Dim strCadMtoGP           As String
Dim rstCtrlPre            As New ADODB.Recordset

    '--- Control de consumo Presupuestos
    If traerCampo("Parametros", "ValParametro", "GlsParametro", "VALIDA_MONTO_PRESUPUESTADO_POR_RUBRO", True) = "S" Then
        nPC = ComputerName
        nPC = Replace(nPC, "-", "")
        nPC = Replace(nPC, "", Trim(nPC))
        nPC = Trim(nPC)
        strVarAux = ""
        strVarCC = ""
        strNomTbl = ""
        
        rstCtrlPre.Fields.Append ("item"), adInteger, 2, adFldRowID
        rstCtrlPre.Fields.Append ("idCentroCosto"), adChar, 8, adFldIsNullable
        rstCtrlPre.Open
        
        With g
            .Dataset.First
            If Not .Dataset.EOF Then
                Do While Not .Dataset.EOF
                    rstCtrlPre.AddNew
                    rstCtrlPre.Fields("item") = "" & .Columns.ColumnByFieldName("item").Value
                    rstCtrlPre.Fields("IdCentroCosto") = "" & .Columns.ColumnByFieldName("IdCentroCosto").Value
                .Dataset.Next
                Loop
            End If
        End With
        
        If Not rstCtrlPre.EOF Then
            rstCtrlPre.MoveFirst
            rstCtrlPre.Sort = "IdCentroCosto"
            Do While Not rstCtrlPre.EOF
                strVarAux = rstCtrlPre.Fields("IdCentroCosto").Value
                If strVarAux <> "" Then
                    Do While strVarAux = rstCtrlPre.Fields("IdCentroCosto").Value
                        strVarCC = rstCtrlPre.Fields("IdCentroCosto").Value
                        rstCtrlPre.MoveNext
                        If rstCtrlPre.EOF Then Exit Do
                        If strVarAux <> rstCtrlPre.Fields("IdCentroCosto").Value & "" Then Exit Do
                    Loop
                End If
                       
                If strVarCC <> "" Then
'''''                    If traerPEx(strVarCC, StrMsgError) <> "" Then
'''''                        If strNomTbl = "" Then
'''''                            strNomTbl = "W" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_COMPRAS")
'''''                        End If
'''''                        csql = "Call Spu_ControlaGatosPresupuestados('" & glsEmpresa & "','" & strVarCC & "','" & strNomTbl & "','" & tipoOperacion & "','CS','" & glsSucursal & "','" & strTD & "','" & F.txt_NumDoc.Text & "','" & strSerie & "','','','','',' ','')"
'''''                        Cn.Execute (csql)
'''''                        strVarCC = ""
'''''
'''''                    End If
                    If StrMsgError <> "" Then GoTo Err
                End If
            Loop
        End If
        If rstCtrlPre.State = 1 Then rstCtrlPre.Close: Set rstCtrlPre = Nothing
    End If
            
    indTrans = False
    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Or TypeOf C Is ComboBox Then
            If C.Tag <> "" Then 'And C.Visible = True Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                
                If UCase(strCampo) = UCase("idMotivoTraslado") Then
                    strCodMotivoTranslado = C.Value
                End If
                
                If UCase(strCampo) = UCase("idPerVendedorCampo") Then
                    strCodVendedorCampo = C.Value
                End If
                
                If UCase(strCampo) = UCase("idPerCliente") Then
                    strCodCliente = C.Value
                End If
                
                If UCase(strCampo) = UCase("idMoneda") Then
                    strMoneda = C.Value
                End If
                
                If UCase(strCampo) = UCase("TipoCambio") Then
                    strtipocambio = Val(C.Value)
                End If
                
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        If UCase(strCampo) = UCase("idDocVentas") Then
                            If Trim(C.Value) = "" Then
                                csql = "SELECT MAX(idDocVentas) FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & strTD & "' "
                                rsCorrela.Open csql, Cn, adOpenStatic, adLockReadOnly
                                If Not rsCorrela.EOF Then
                                    strCod = Format(Val("" & rsCorrela.Fields(0)) + 1, "00000000")
                                Else
                                    strCod = "00000001"
                                End If
                                C.Text = strCod
                            Else
                                If traerCampo("docventas", "idDocVentas", "idDocumento", strTD, True, "idSucursal = '" & glsSucursal & "' AND idDocVentas = '" & Trim(C.Value) & "' AND idSerie = '" & strSerie & "'") <> "" Then
                                    StrMsgError = "El Documento numero " & C.Value & " con serie " & strSerie & " ya existe"
                                    GoTo Err
                                End If
                            End If
                            strValCod = Trim(C.Value)
                        End If
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & Val(C.Value) & ","
                            Case "T"
                                '''strValores = strValores & "'" & Trim(C.Value) & "',"
                                If (TypeOf C Is ComboBox) Then
                                    strValores = strValores & "'" & Trim(C.Text) & "',"
                                Else
                                    strValores = strValores & "'" & Trim(C.Value) & "',"
                                End If
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    
                    Case 1
                        If UCase(strCampo) <> UCase("idDocVentas") Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = Val(C.Value)
                                Case "T"
                                    If (TypeOf C Is ComboBox) Then
                                        strValores = "'" & C.Text & "'"
                                    Else
                                        strValores = "'" & C.Value & "'"
                                    End If
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
            If Trim(traerCampo("Docventas", "idDocventas", "idDocventas", strCod, True, "idSerie = '" & strSerie & "' And idSucursal = '" & glsSucursal & "' And idDocumento = '" & strTD & "' ")) <> "" Then
                strCod = generaCorrelativoDocVentas(strTD, strSerie)
                F.txt_NumDoc.Text = strCod
            End If
        
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            csql = "INSERT INTO docventas(" & strCampos & ",idDocumento,idEmpresa,idSucursal,FecRegistro,idUsuarioReg) VALUES(" & strValores & ",'" & strTD & "','" & glsEmpresa & "','" & glsSucursal & "',Getdate(),'" & glsUser & "')"
        Case 1
            csql = "UPDATE docventas SET " & strCampos & " WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strValCod & "' AND idSerie = '" & strSerie & "'"
    End Select
    
    indTrans = True
    Cn.BeginTrans
    
    'Graba controles
    Cn.Execute csql
    
    'Grabando Grilla detalle
    If TypeName(g) <> "Nothing" Then
        Cn.Execute "DELETE FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strValCod & "' AND idSerie = '" & strSerie & "'"
        
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
                            strValores = strValores & Val(g.Columns(i).Value) & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO docventasdet(" & strCampos & ",idDocumento,idDocVentas,idSerie,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strTD & "','" & strValCod & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "')"
            Cn.Execute csql
            g.Dataset.Next
        Loop
    End If
    
    '--- Grabando Grilla DocRef
    If TypeName(D) <> "Nothing" Then
        Cn.Execute "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = '" & strTD & "' AND numDocOrigen = '" & strValCod & "' AND serieDocOrigen = '" & strSerie & "'"
        
        D.Dataset.First
        Do While Not D.Dataset.EOF
            If Trim(D.Columns.ColumnByFieldName("idDocumento").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idSerie").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idNumDoc").Value) <> "" Then
                strCampos = ""
                strValores = ""
                For i = 0 To D.Columns.Count - 1
                    If UCase(left(D.Columns(i).ObjectName, 1)) = "W" Then
                        strTipoDato = Mid(D.Columns(i).ObjectName, 2, 1)
                        strCampo = Mid(D.Columns(i).ObjectName, 3)
                        strCampos = strCampos & strCampo & ","
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & D.Columns(i).Value & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(D.Columns(i).Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(D.Columns(i).Value, "yyyy-mm-dd") & "',"
                        End Select
                    End If
                Next
                
                If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
                If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
                
                csql = "INSERT INTO docreferencia(" & strCampos & ",tipoDocOrigen,numDocOrigen,serieDocOrigen,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strTD & "','" & strValCod & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "')"
                        
                Cn.Execute csql
            End If
            D.Dataset.Next
            If strTD = "87" Then
                Cn.Execute "Update DocVentas Set IndReq = '1' Where IdEmpresa = '" & glsEmpresa & "' And IdDocumento = '" & D.Columns.ColumnByFieldName("IdDocumento").Value & "' And IdSerie = '" & D.Columns.ColumnByFieldName("IdSerie").Value & "' And IdDocVentas = '" & D.Columns.ColumnByFieldName("IdDocVentas").Value & "'"
            End If
        Loop
    End If
    
    If strTD <> "87" Then
        csql = "SELECT d.idDocumentoImp,d.idDocVentasImp,d.idSerieImp,d.idProducto,d.idUM,d.Cantidad " & _
                "FROM docventasdet d " & _
                "WHERE d.idEmpresa = '" & glsEmpresa & "'" & _
                "AND d.idSucursal = '" & glsSucursal & "'" & _
                "AND d.idDocumento = '" & strTD & "'" & _
                "AND d.idDocVentas = '" & strValCod & "'" & _
                "AND d.idSerie = '" & strSerie & "'" & _
                "AND d.idDocVentasImp <> ''"
            
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            '--- ACTUALIZAMOS CANTIDAD IMPORTADA
            csql = "UPDATE docventasdet dd  SET dd.CantidadImp = dd.CantidadImp + " & CStr(rst.Fields("Cantidad")) & " " & _
                    "WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idSucursal = '" & glsSucursal & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                    "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "'"
            Cn.Execute csql
        
            csql = "UPDATE docventasdet dd  SET dd.estDocImportado = 'S' " & _
                    "WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idSucursal = '" & glsSucursal & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                    "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "' AND dd.CantidadImp >= dd.Cantidad"
            Cn.Execute csql
                
            rst.MoveNext
        Loop
    
        '--- Marcamos la Cabecera
        csql = "SELECT DISTINCT idDocumentoImp,idDocVentasImp,idSerieImp  FROM docventasdet " & _
                "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strValCod & "' AND idSerie = '" & strSerie & "' AND idDocVentasImp <> '' AND idDocVentasImp is not null"
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            csql = "UPDATE docVentas c,docVentasDet d  SET c.estDocImportado = 'S' " & _
                   "WHERE c.idEmpresa = d.idEmpresa AND c.idSucursal = d.idSucursal AND c.idDocumento = d.idDocumento AND c.idDocVentas = d.idDocVentas AND c.idSerie = d.idSerie " & _
                   "AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND d.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND d.idSerie = '" & rst.Fields("idSerieImp") & "'" & _
                   "AND (SELECT count(x.idDocVentas) FROM docVentasDet x WHERE x.idEmpresa = d.idEmpresa AND x.idSucursal = d.idSucursal AND x.idDocumento = d.idDocumento AND x.idDocVentas = d.idDocVentas AND x.idSerie = d.idSerie AND x.estDocImportado <> 'S') = 0"
            Cn.Execute csql
            
            rst.MoveNext
        Loop
    End If
    
    '--- Si la empresa es ALSISAC UPDATE a docventas  el estado PAR(Parcial),CER(Cerrado)
    If traerCampo("Empresas", "RUC", "idEmpresa", glsEmpresa, False) = "20430471750" Then
        If D.Columns.ColumnByFieldName("idDocumento").Value = "87" Then
            strcant = 0
            cantReg = 0
            If TypeName(D) <> "Nothing" Then
                D.Dataset.First
                Do While Not D.Dataset.EOF
                 strcant = 0
                 cantReg = 0
                    If Trim(D.Columns.ColumnByFieldName("idDocumento").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idSerie").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idNumDoc").Value) <> "" Then
               
                        csql = "SELECT  d.idDocumentoImp,d.idDocVentasImp,d.idSerieImp,d.idProducto,d.idUM,d.Cantidad,d.estDocImportado,d.CantidadImp " & _
                                "FROM  docVentas c Inner Join docVentasDet d " & _
                                "On c.idempresa = d.idempresa AND c.idSucursal = d.idSucursal AND c.idDocumento = d.idDocumento AND c.idDocVentas = d.idDocVentas " & _
                                "WHERE d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & D.Columns.ColumnByFieldName("idDocumento").Value & "'" & _
                                "AND d.idDocVentas = '" & D.Columns.ColumnByFieldName("idDocventas").Value & "'" & _
                                "AND d.idSerie = '" & D.Columns.ColumnByFieldName("idSerie").Value & "' "
                                
                        If rst.State = 1 Then rst.Close
                        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                        If Not rst.EOF Then
                            rst.MoveFirst
                            
                            Do While Not rst.EOF
                            cantReg = cantReg + 1
                                If (rst.Fields("Cantidad") = rst.Fields("CantidadImp") And rst.Fields("estDocImportado") = "S") Then
                                    strcant = strcant + 1
                                End If
                                rst.MoveNext
                            Loop
                            
                            If cantReg = strcant Then
                                csql = "UPDATE docVentas d Set estDocVentas='CER' WHERE d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "'" & _
                                       "AND d.idDocumento = '" & D.Columns.ColumnByFieldName("idDocumento").Value & "' AND d.idDocVentas = '" & D.Columns.ColumnByFieldName("idDocventas").Value & "' AND d.idSerie = '" & D.Columns.ColumnByFieldName("idSerie").Value & "' "
                            Else
                               csql = "UPDATE docVentas d Set estDocVentas='PAR' WHERE d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & D.Columns.ColumnByFieldName("idDocumento").Value & "'" & _
                                      "AND d.idDocVentas = '" & D.Columns.ColumnByFieldName("idDocventas").Value & "' AND d.idSerie = '" & D.Columns.ColumnByFieldName("idSerie").Value & "' "
                            End If
                            Cn.Execute csql
                         End If
                     End If
                     D.Dataset.Next
                  Loop
            End If
        End If
    End If
    '--- Actualizamos la tabla productosproveedor
    StrItem = 0
    StrItem = Val(traerCampo("ProductosProveedor", "Max(Item)", "IdProveedor", strCodCliente, True))
    g.Dataset.First
    Do While Not g.Dataset.EOF
        '--- Calculamos monto a actualizar
        valValorVenta = 0
        valIGVVenta = 0
        valPrecioVenta = 0
        
        'If strMoneda = "USD" Then
        '    valValorVenta = Format(Val(G.Columns.ColumnByFieldName("VVUnit").Value) * strTipoCambio, "0.00")
        '    valIGVVenta = Format((Val(G.Columns.ColumnByFieldName("VVUnit").Value) * glsIGV) * strTipoCambio, "0.00")
        '    valPrecioVenta = Format((Val(G.Columns.ColumnByFieldName("VVUnit").Value) * (glsIGV + 1)) * strTipoCambio, "0.00")
        'Else
        
        valValorVenta = Format(Val(g.Columns.ColumnByFieldName("VVUnit").Value), "0.00000")
        valIGVVenta = Format(Val(g.Columns.ColumnByFieldName("VVUnit").Value) * glsIGV, "0.00000")
        valPrecioVenta = Format(Val(g.Columns.ColumnByFieldName("VVUnit").Value) * (glsIGV + 1), "0.00000")
        
        'End If
        StrItem = StrItem + 1
        swNuevoProducto = ""
        swNuevoProducto = traerCampo("ProductosProveedor", "IdProducto", "IdProducto", g.Columns.ColumnByFieldName("IdProducto").Value, True, "IdProveedor = '" & strCodCliente & "' And IdDocumento = '" & strTD & "' And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strValCod & "'")
        
        If swNuevoProducto = "" Then
            
            csql = "Insert Into ProductosProveedor(IdEmpresa,IdProveedor,IdProducto,Fecha,IdDocumento,IdSerie,IdDocVentas,IdMoneda,ValorVenta," & _
                   "IGVVenta,PrecioVenta,Item)Values" & _
                   "('" & glsEmpresa & "','" & strCodCliente & "','" & g.Columns.ColumnByFieldName("IdProducto").Value & "'," & _
                   "'" & Format(F.dtp_Emision.Value, "yyyy-mm-dd") & "','" & strTD & "','" & strSerie & "','" & strValCod & "','" & F.txtCod_Moneda.Text & "'," & _
                   "" & valValorVenta & "," & valIGVVenta & ", " & valPrecioVenta & ", " & StrItem & ")"
            Cn.Execute csql
            
        Else
            csql = "Update ProductosProveedor " & _
                   "Set Fecha = '" & Format(F.dtp_Emision.Value, "yyyy-mm-dd") & "',ValorVenta = " & valValorVenta & ",IGVVenta = " & valIGVVenta & "," & _
                   "PrecioVenta = " & valPrecioVenta & " " & _
                   "where IdEmpresa = '" & glsEmpresa & "' And IdProveedor = '" & strCodCliente & "' " & _
                   "And IdProducto = '" & g.Columns.ColumnByFieldName("IdProducto").Value & "' And IdDocumento = '" & strTD & "' " & _
                   "And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strValCod & "'"
            Cn.Execute csql
        End If
        
        g.Dataset.Next
    Loop
    
    '--- Control de consumo Presupuestos
    If traerCampo("Parametros", "ValParametro", "GlsParametro", "VALIDA_MONTO_PRESUPUESTADO_POR_RUBRO", True) = "S" And (strTD = "OS" Or strTD = "69") And strNomTbl <> "" Then
    strCadMtoGP = IIf(tipoOperacion = "0", "IFNULL((A.TotalConsumido + A.TotalVVNeto),0) As TotalCP,0 As Diferencia", "IfNull((TotalConsumido + (A.TotalVVNeto - TotalVVNetoAnt)),IFNULL((A.TotalConsumido + A.TotalVVNeto),0)) As TotalCP,IFNULL((A.TotalVVNeto - TotalVVNetoAnt),0) As Diferencia")
    
        csql = "Select A.idCentroCosto,A.idGrupo,A.TotalRubro,A.TotalConsumido,A.TotalVVNeto,A.GlsGrupo," & strCadMtoGP & _
                " From( " & _
                "Select A.idCentroCosto,A.idGrupo,A.TotalRubro,A.TotalConsumido,TotalVVNeto,B.GlsGrupo,TotalVVNetoAnt " & _
                "From " & strNomTbl & "  A " & _
                "Left Join( " & _
                "Select GP.idGrupo,DD.idCentroCosto,Sum(If(C.IdMoneda = D.IdMoneda,DD.TotalVVNeto,If(C.IdMoneda = 'PEN',DD.TotalVVNeto * D.TipoCambio,DD.TotalVVNeto / D.TipoCambio))) As TotalVVNeto,GlsGrupo,D.TipoCambio,D.idMoneda " & _
                "From Docventas D " & _
                "Inner Join DocventasDet DD " & _
                    "On D.idEmpresa = DD.idEmpresa And D.idSucursal = DD.idSucursal " & _
                    "And D.idSerie =  DD.idSerie And D.idDocventas = DD.idDocventas " & _
                    "And D.idDocumento = DD.idDocumento  " & _
                "Inner Join Productos P  " & _
                    "On P.idEmpresa = DD.idEmpresa " & _
                    "And P.idProducto = DD.idProducto  " & _
                "Inner Join GruposProducto GP  " & _
                    "On P.idEmpresa = GP.idEmpresa " & _
                    "And P.idGrupo =  GP.idGrupo " & _
                "Inner Join CentrosCosto C " & _
                    "On DD.idCentroCosto = C.idCentroCosto " & _
                    "And DD.IdEmpresa = C.idEmpresa " & _
                "Where D.idEmpresa = '" & glsEmpresa & "' And D.idDocventas = '" & strValCod & "' And D.idSerie = '" & strSerie & "' " & _
                "And D.idDocumento  = '" & strTD & "' And D.idSucursal = '" & glsSucursal & "' " & _
                "Group By D.IdEmpresa,DD.IdCentroCosto,GP.idGrupo) B " & _
                "On  A.idGrupo = B.idGrupo And A.idCentroCosto = B.idCentroCosto ) A Order By A.idCentroCosto,A.idGrupo "
        
        rstCGP.Open csql, Cn, adOpenStatic, adLockReadOnly
        rstCGP.MoveFirst
        Do While Not rstCGP.EOF
                If rstCGP.Fields("TotalCP").Value > rstCGP.Fields("TotalRubro").Value Then
                        StrMsgError = StrMsgError & "El rubro " & rstCGP.Fields("GlsGrupo").Value & " de la H.C. " & rstCGP.Fields("idCentroCosto").Value & " excede lo presupuestado. " & vbCrLf & " Monto Presupuestado " & Format(rstCGP.Fields("TotalRubro").Value, "0.00") & vbCrLf & " Monto Gastado " & Format(rstCGP.Fields("TotalConsumido").Value, "0.00") & vbCrLf & " Monto a Consumir " & IIf(tipoOperacion = "0", Format(rstCGP.Fields("TotalVVNeto").Value, "0.00"), IIf(Format(rstCGP.Fields("Diferencia").Value, "0.00") = 0, Format(rstCGP.Fields("TotalVVNeto").Value, "0.00"), Format(rstCGP.Fields("Diferencia").Value, "0.00"))) & Chr(13) & Chr(10)
                End If
            rstCGP.MoveNext
        Loop
        
        If tipoOperacion = 0 Then
            If StrMsgError <> "" Then
                F.txt_NumDoc.Text = ""
            End If
        End If
        strNomTbl = IIf(strNomTbl = "", "0a", strNomTbl)
        
        Set rsEli = DataProcedimiento("Spu_EliminaTemporales", StrMsgError, "0a", "0a", "0a", "0a", "0a", strNomTbl)
        strNomTbl = ""
        
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Cn.CommitTrans
    
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If rstCGP.State = 1 Then rstCGP.Close: Set rstCGP = Nothing
    If rsEli.State = 1 Then rsEli.Close: Set rsEli = Nothing
    
    Exit Sub

Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If rsEli.State = 1 Then rsEli.Close: Set rsEli = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
End Sub





Public Sub EjecutaSQLFormDocVentas(F As Form, ByVal tipoOperacion As Integer, ByRef StrMsgError As String, ByVal strTD As String, ByVal strSerie As String, ByVal g As dxDBGrid, ByVal D As dxDBGrid, ByVal indGeneraVale As Boolean, ByRef rsTempDocVlote As ADODB.Recordset, ByRef RsDetProductos As ADODB.Recordset, ByRef rsTempLiquidacion As ADODB.Recordset, NArrTotales() As Double)
On Error GoTo Err
Dim C               As Object
Dim csql            As String
Dim strCampo        As String
Dim strTipoDato     As String
Dim strCampos       As String
Dim strValores      As String
Dim strValCod       As String
Dim strCodMotivoTranslado       As String
Dim strCodCliente               As String
Dim strCodVendedorCampo         As String
Dim rst                         As New ADODB.Recordset
Dim strCod                      As String
Dim strCodFacRef                As String
Dim obsdocventas                As String
Dim RsTempVale                  As New ADODB.Recordset
Dim rstsepar                    As New ADODB.Recordset
Dim rstcab                      As New ADODB.Recordset
Dim i                           As Integer
Dim cadUpd                      As String
Dim rstUpd                      As New ADODB.Recordset
Dim strParamItemPro             As String
Dim cantidadtot                 As Double
Dim RstDet                      As New ADODB.Recordset
Dim item                        As Integer
Dim rstdocref                   As New ADODB.Recordset
Dim Cadmysql                    As String
Dim rstdetref                   As New ADODB.Recordset
Dim cantidad                    As Double
Dim producto1                   As String
Dim producto2                   As String
Dim cantidad1                   As Double
Dim cantidad2                   As Double
Dim rsh                         As New ADODB.Recordset
Dim rsliquidacion       As New ADODB.Recordset
Dim CadMysqlLiq         As String
Dim CadMysqlLiqAct      As String
Dim idValeDocventas  As String
Dim strTipoVale      As String
Dim strSigno         As String
Dim strCodAlmacen    As String

    '--- Parametro que evaluará la liberación del detalle de un documento de  referencia por el item del producto
    strParamItemPro = Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "EVALUA_ITEM_DETALLE_PRODUCTO", True))
    
    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                
                If UCase(strCampo) = UCase("idMotivoTraslado") Then
                    strCodMotivoTranslado = C.Value
                End If
                
                If UCase(strCampo) = UCase("idPerVendedorCampo") Then
                    strCodVendedorCampo = C.Value
                End If
                
                If UCase(strCampo) = UCase("idPerCliente") Then
                    strCodCliente = C.Value
                End If
    
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        If UCase(strCampo) = UCase("idDocVentas") Then
                            If Trim(C.Value) = "" Then
                                strCod = generaCorrelativoDocVentas(strTD, strSerie)
                                C.Text = strCod
                            Else
                                If traerCampo("docventas", "idDocVentas", "idDocumento", strTD, True, "idSucursal = '" & glsSucursal & "' AND idDocVentas = '" & Trim(C.Value) & "' AND idSerie = '" & strSerie & "'") <> "" Then
                                    StrMsgError = "El Documento numero " & C.Value & " con serie " & strSerie & " ya existe"
                                    GoTo Err
                                End If
                            End If
                            strValCod = Trim(C.Value)
                        End If
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & Val(C.Value) & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    Case 1
                        If UCase(strCampo) <> UCase("idDocVentas") Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = Val(C.Value)
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
            csql = "INSERT INTO docventas(" & strCampos & ",idDocumento,idEmpresa,idSucursal,FecRegistro,idUsuarioReg,TotalValorVenta,TotalIGVVenta,TotalIvap," & _
                   "TotalDsctoVV,TotalDsctoPV,TotalBaseImponible,TotalExonerado)VALUES(" & _
                   "" & strValores & ",'" & strTD & "','" & glsEmpresa & "','" & glsSucursal & "',GETDATE(),'" & glsUser & "'," & NArrTotales(0) & "," & _
                   "" & NArrTotales(1) & "," & NArrTotales(2) & "," & NArrTotales(3) & "," & NArrTotales(4) & "," & NArrTotales(5) & "," & NArrTotales(6) & ")"
            
        Case 1
            csql = "UPDATE docventas SET " & strCampos & ",TotalValorVenta = " & NArrTotales(0) & ",TotalIGVVenta = " & NArrTotales(1) & "," & _
                   "TotalIvap = " & NArrTotales(2) & ",TotalDsctoVV = " & NArrTotales(3) & ",TotalDsctoPV = " & NArrTotales(4) & "," & _
                   "TotalBaseImponible = " & NArrTotales(5) & ",TotalExonerado = " & NArrTotales(6) & " " & _
                   "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strValCod & "' AND idSerie = '" & strSerie & "'"
                   
    End Select

    '--- Graba controles
    Cn.Execute csql
    
    '--- Grabando Grilla detalle
    If TypeName(g) <> "Nothing" Then
        Cn.Execute "DELETE FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strValCod & "' AND idSerie = '" & strSerie & "'"
        
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
                            strValores = strValores & Val(g.Columns(i).Value) & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO docventasdet(" & strCampos & ",idDocumento,idDocVentas,idSerie,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strTD & "','" & strValCod & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "')"
            Cn.Execute csql
            
            g.Dataset.Next
        Loop
    End If
    
    '--- Grabando Grilla DocRef
    If TypeName(D) <> "Nothing" Then
        Cn.Execute "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = '" & strTD & "' AND numDocOrigen = '" & strValCod & "' AND serieDocOrigen = '" & strSerie & "'"
        
        If sw_graba_factura = True And strTD = "86" Then
            Cn.Execute "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = '01' AND tipoDocReferencia = '" & strTD & "' AND numDocReferencia = '" & strValCod & "' AND serieDocReferencia = '" & strSerie & "'"
        End If
        
        D.Dataset.First
        Do While Not D.Dataset.EOF
            If Trim(D.Columns.ColumnByFieldName("idDocumento").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idSerie").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idNumDoc").Value) <> "" Then
                strCampos = ""
                strValores = ""
                strCodFacRef = ""
                For i = 0 To D.Columns.Count - 1
                    If UCase(left(D.Columns(i).ObjectName, 1)) = "W" Then
                        strTipoDato = Mid(D.Columns(i).ObjectName, 2, 1)
                        strCampo = Mid(D.Columns(i).ObjectName, 3)
                        
                        strCampos = strCampos & strCampo & ","
                        If strCampo = "tipoDocReferencia" Then strCodFacRef = D.Columns(i).Value
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & D.Columns(i).Value & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(D.Columns(i).Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(D.Columns(i).Value, "yyyy-mm-dd") & "',"
                        End Select
                    End If
                Next
                
                If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
                If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
                
                If sw_graba_factura = True And strTD = "86" And strCodFacRef = "01" Then
                    csql = "INSERT INTO docreferencia(Item,tipoDocOrigen,serieDocOrigen,IndImportado,numDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strTD & "','" & strValCod & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "')"
                    Cn.Execute csql
                End If
                
                csql = "INSERT INTO docreferencia(" & strCampos & ",tipoDocOrigen,numDocOrigen,serieDocOrigen,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strTD & "','" & strValCod & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "')"
                Cn.Execute csql
            End If
            D.Dataset.Next
        Loop
    End If
    
    '--- Actualizamos el codigo del vendedor de campo al cliente
    If strCodVendedorCampo <> "" Then
        csql = "UPDATE clientes SET idVendedorCampo = '" & strCodVendedorCampo & "' WHERE idCliente = '" & strCodCliente & "' AND idEmpresa = '" & glsEmpresa & "'"
        Cn.Execute csql
    End If
    
    If Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "GENERA_VALE_FORMULA", True) & "") = "S" Then
        csql = "Delete from DocventasDetFormula where iddocventas = '" & strValCod & "' and iddocumento = '" & strTD & "' and idserie = '" & strSerie & "' and idempresa = '" & glsEmpresa & "' and idsucursal = '" & glsSucursal & "'"
        Cn.Execute (csql)
        
        With RsDetProductos
            If Not .EOF Then
                .MoveFirst
                Do While Not .EOF
                    csql = "Insert Into DocventasDetFormula(IdDocumento,IdSerie,IdDocVentas,ItemDocventas,IdProducto,GlsProducto,CantidadStock," & _
                            "Cantidad,Item,IdComboCab,IdEmpresa,IdSucursal,VVUnit,PorcDcto,TotalVVenta)Values(" & _
                            "'" & strTD & "','" & strSerie & "','" & strValCod & "'," & .Fields("ItemDocVentas") & ",'" & .Fields("IdProducto") & "'," & _
                            "'" & .Fields("GlsProducto") & "',0," & .Fields("Cantidad") & "," & .Fields("Item") & ",'" & .Fields("IdComboCab") & "'," & _
                            "'" & glsEmpresa & "','" & glsSucursal & "'," & .Fields("VVUnit") & "," & .Fields("PorcDcto") & "," & .Fields("TotalVVenta") & ")"
                    Cn.Execute (csql)
                    
                    .MoveNext
                Loop
            End If
        End With
    
    End If
    
    If Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "LIQUIDACIONES", True) & "") = "S" Then
        If strTD = "86" Or strTD = "03" Then
            CadMysqlLiq = "Select * from docventasdetliquidacion " & _
                          "where iddocventas = '" & strValCod & "' and iddocumento = '" & strTD & "' and idserie = '" & strSerie & "' " & _
                          "and idempresa = '" & glsEmpresa & "' and idsucursal = '" & glsSucursal & "' "
            If rsliquidacion.State = 1 Then rsliquidacion.Close
            rsliquidacion.Open CadMysqlLiq, strcn, adOpenStatic, adLockOptimistic
            If Not rsliquidacion.EOF Then
                rsliquidacion.MoveFirst
                Do While Not rsliquidacion.EOF
                
                    CadMysqlLiqAct = "Update docventasliqcab set CantidadImp = (CantidadImp - " & rsliquidacion.Fields("Unidad") & ") ,PesoImp = (PesoImp - " & (rsliquidacion.Fields("Kg") + rsliquidacion.Fields("ReKg")) & ") " & _
                                     "where IdEmpresa = '" & glsEmpresa & "' and  IdSucursal = '" & glsSucursal & "' and  IdLiquidacion = '" & Trim("" & rsliquidacion.Fields("IdLiquidacion")) & "' "
                    Cn.Execute (CadMysqlLiqAct)
                    
                    rsliquidacion.MoveNext
                Loop
            End If
            
            csql = "Delete from docventasdetliquidacion where iddocventas = '" & strValCod & "' and iddocumento = '" & strTD & "' and idserie = '" & strSerie & "' and idempresa = '" & glsEmpresa & "' and idsucursal = '" & glsSucursal & "' "
            Cn.Execute (csql)
            
            If rsTempLiquidacion.State = 1 Then
                With rsTempLiquidacion
                    If Not .EOF Then
                        .MoveFirst
                        Do While Not .EOF
                            csql = "Insert Into docventasdetliquidacion(IdDocumento, IdSerie, IdDocventas, IdProducto, Item, IdLiquidacion, IdUPP," & _
                                   "idCamal , FechaLiq, UnidadSaldo, Unidad, KgSaldo, ReKg,Kg, KgVivo, IdEmpresa, Idsucursal,KgEnv)Values(" & _
                                   "'" & strTD & "','" & strSerie & "','" & strValCod & "','" & .Fields("IdProducto") & "'," & .Fields("ItemDocVentas") & "," & _
                                   "'" & .Fields("IdLiquidacion") & "','" & .Fields("IdUPP") & "','" & .Fields("idCamal") & "','" & Format(.Fields("FechaLiq"), "yyyy-mm-dd") & "'," & .Fields("UnidadSaldo") & "," & .Fields("Unidad") & "," & .Fields("KgSaldo") & "," & .Fields("ReKg") & "," & .Fields("Kg") & "," & .Fields("KgVivo") & ", " & _
                                   "'" & glsEmpresa & "','" & glsSucursal & "'," & .Fields("KgEnv") & ")"
                            Cn.Execute (csql)
                            
                            CadMysqlLiqAct = "Update docventasliqcab set CantidadImp = (CantidadImp + " & .Fields("Unidad") & ") ,PesoImp = (PesoImp + " & (.Fields("Kg") + .Fields("ReKg")) & ") " & _
                                             "where IdEmpresa = '" & glsEmpresa & "' and  IdSucursal = '" & glsSucursal & "' and  IdLiquidacion = '" & Trim("" & .Fields("IdLiquidacion")) & "' "
                            Cn.Execute (CadMysqlLiqAct)
                            
                            .MoveNext
                        Loop
                        
                    End If
                End With
            End If
        End If
    End If
        
    '--- Actualizando Vales
    If indGeneraVale Then
        idValeDocventas = Trim("" & traerCampo("docventas", "idValesCab", "iddocventas", strValCod, True, " idserie = '" & strSerie & "' and iddocumento = '" & strTD & "' and idsucursal = '" & glsSucursal & "' "))
        strTipoVale = Trim("" & traerCampo("docventas", "tipoVale", "iddocventas", strValCod, True, " idserie = '" & strSerie & "' and iddocumento = '" & strTD & "' and idsucursal = '" & glsSucursal & "' "))
        
        csql = "Select idvalescab,Cantidad,idEmpresa,idSucursal,idProducto,idUM,idLote from Valesdet " & _
               "Where idvalescab = '" & idValeDocventas & "' " & _
               "and idempresa = '" & glsEmpresa & "' " & _
               "and idsucursal = '" & glsSucursal & "' " & _
               "and tipoVale = '" & strTipoVale & "' "
        If RsTempVale.State = 1 Then RsTempVale.Close
        RsTempVale.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        
        If tipoOperacion = "1" Then
            If Not RsTempVale.EOF Then
                RsTempVale.MoveFirst
                Do While Not RsTempVale.EOF
                    strSigno = "-"
                    If strTipoVale = "S" Then strSigno = "+"
                    strCodAlmacen = traerCampo("valescab", "idalmacen", "idvalescab", Trim("" & RsTempVale.Fields("idvalescab")), True, " idsucursal = '" & glsSucursal & "' and tipoVale = '" & Trim("" & strTipoVale) & "' ")
                        
                    csql = "UPDATE productosalmacen " & _
                            "SET CantidadStock = CantidadStock " & strSigno & " " & RsTempVale.Fields("Cantidad") & " " & _
                            "WHERE idEmpresa = '" & RsTempVale.Fields("idEmpresa") & "' " & _
                              "AND idSucursal = '" & RsTempVale.Fields("idSucursal") & "' " & _
                              "AND idAlmacen = '" & strCodAlmacen & "' " & _
                              "AND idProducto = '" & RsTempVale.Fields("idProducto") & "' " & _
                              "AND idUMCompra = '" & RsTempVale.Fields("idUM") & "' "
                    Cn.Execute csql
                    
                    csql = "UPDATE productosalmacenporlote " & _
                            "SET CantidadStock = CantidadStock " & strSigno & " " & RsTempVale.Fields("Cantidad") & " " & _
                            "WHERE idEmpresa = '" & RsTempVale.Fields("idEmpresa") & "' " & _
                              "AND idSucursal = '" & RsTempVale.Fields("idSucursal") & "' " & _
                              "AND idAlmacen = '" & strCodAlmacen & "' " & _
                              "AND idProducto = '" & RsTempVale.Fields("idProducto") & "' " & _
                              "AND idUMCompra = '" & RsTempVale.Fields("idUM") & "' " & _
                              "AND idLote = '" & RsTempVale.Fields("idLote") & "' "
                    Cn.Execute csql
                                                
                    RsTempVale.MoveNext
                Loop
            End If
        End If
        
        csql = "Delete from docventasdetlote where iddocventas = '" & strValCod & "' and iddocumento = '" & strTD & "' and idserie = '" & strSerie & "' and idempresa = '" & glsEmpresa & "' and idsucursal = '" & glsSucursal & "'"
        Cn.Execute (csql)
        
        If rsTempDocVlote.State = 1 Then
            If rsTempDocVlote.RecordCount > 0 Then
                rsTempDocVlote.MoveFirst
                If Not rsTempDocVlote.EOF Then
                    Do While Not rsTempDocVlote.EOF
                    
                        csql = "Insert Into docventasdetlote(idDocventas, Item, idLote, GlsLote, CantidadStock, Cantidad, idProducto, ItemDocVentas, idSerie, idDocumento,idempresa,idsucursal) " & _
                                " vALUES('" & strValCod & "'," & rsTempDocVlote.Fields("Item") & ",'" & rsTempDocVlote.Fields("idLote") & "','" & rsTempDocVlote.Fields("GlsLote") & "', " & _
                                "" & rsTempDocVlote.Fields("CantidadStock") & "," & rsTempDocVlote.Fields("Cantidad") & ",'" & rsTempDocVlote.Fields("idProducto") & "', " & _
                                "'" & rsTempDocVlote.Fields("ItemDocVentas") & "','" & strSerie & "','" & strTD & "','" & glsEmpresa & "','" & glsSucursal & "') "
                        Cn.Execute (csql)
                    
                        rsTempDocVlote.MoveNext
                    Loop
                End If
            End If
        End If
        
        If Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "STOCK_POR_LOTE", True) & "") = "S" Then
            generaValeDocVentasLote strTD, strSerie, Trim("" & strTipoVale), strValCod, strCodMotivoTranslado, F.dtp_Emision, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        ElseIf Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "GENERA_VALE_FORMULA", True) & "") = "S" Then
            If traerCampo("CentrosCosto", "IndGenerado", "IdCentroCosto", traerCampo("DocVentas", "IdCentroCosto", "IdDocumento", strTD, True, "IdSucursal = '" & glsSucursal & "' And IdSerie = '" & strSerie & "' And IdDocVentas = '" & strValCod & "'"), True) = "A" Then
                generaValeDocVentas strTD, strSerie, Trim("" & strTipoVale), strValCod, strCodMotivoTranslado, F.dtp_Emision, F.txtCod_CentroCosto.Text, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            
            Else
                generaValeDocVentasFormula strTD, strSerie, Trim("" & strTipoVale), strValCod, strCodMotivoTranslado, F.dtp_Emision, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            
        Else
            If traerCampo("seriesdocumento", "IndTransferencia", "idDocumento", strTD, True, " idserie = '" & strSerie & "' ") = "S" Then
                generaValeDocVentas strTD, strSerie, Trim("" & strTipoVale), strValCod, strCodMotivoTranslado, F.dtp_Emision, F.txtCod_CentroCosto.Text, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                '--- GRABA EL VALE DE INGRESO, PARA LAS TRANSFERENCIAS ENTRE SUCURSALES
                generaValeDocVentasTransfIngreso strTD, strSerie, "I", strValCod, strCodMotivoTranslado, F.dtp_Emision, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                csql = ("Call Spu_InsertValesTrans('" & glsEmpresa & "','" & strTD & "','" & F.txt_Serie.Text & "','" & F.txt_NumDoc.Text & "','" & glsSucursal & "','" & GeneraCorrelativoAnoMes("valestrans", "idValesTrans") & "') ")
                Cn.Execute ("Call Spu_InsertValesTrans('" & glsEmpresa & "','" & strTD & "','" & F.txt_Serie.Text & "','" & F.txt_NumDoc.Text & "','" & glsSucursal & "','" & GeneraCorrelativoAnoMes("valestrans", "idValesTrans") & "') ")
            Else
                If strTD = "07" Then
                    strCodMotivoTranslado = traerCampo("DocVentas", "IdMotivoNCD", "IdDocumento", strTD, True, "IdSerie = '" & strSerie & "' And IdDocVentas = '" & strValCod & "'")
                End If
                
                generaValeDocVentas strTD, strSerie, Trim("" & strTipoVale), strValCod, strCodMotivoTranslado, F.dtp_Emision, F.txtCod_CentroCosto.Text, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
        End If
        
    End If
    
    '--- Si es guia verificamos el motivo y si esta con la marca de recepcion de mercaderia auomatico
    If strTD = "86" And glsRecepcionAuto Then
        If strCodMotivoTranslado = "06090006" Or strCodMotivoTranslado = "08050001" Then
            generaValeRecMercaderiaAuto strTD, strSerie, strValCod, strCodMotivoTranslado, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
    End If
    
    If traerCampo("Parametros", "ValParametro", "GlsParametro", "AGRUPAPRODUCTOS", True) = "N" Then
        '--- Cambiamos el estado a Importado
        csql = "SELECT d.idDocumentoImp,d.idDocVentasImp,d.idSerieImp,d.idProducto,d.idUM,d.Cantidad,d.iTemPro " & _
                "FROM docventasdet d " & _
                "WHERE d.idEmpresa = '" & glsEmpresa & "'" & _
                "AND d.idSucursal = '" & glsSucursal & "'" & _
                "AND d.idDocumento = '" & strTD & "'" & _
                "AND d.idDocVentas = '" & strValCod & "'" & _
                "AND d.idSerie = '" & strSerie & "'" & _
                "AND d.idDocVentasImp <> ''"
                
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rst.EOF Then
            rst.MoveFirst
            rst.Sort = "idProducto"
                Do While Not rst.EOF
                    '--- ACTUALIZAMOS CANTIDAD IMPORTADA
                    If strParamItemPro = "S" Then
                    
                        csql = "UPDATE  dd  SET dd.CantidadImp = dd.CantidadImp + " & CStr(rst.Fields("Cantidad")) & " FROM docventasdet dd " & _
                               "WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idSucursal = '" & glsSucursal & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                                 "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "' AND dd.ItemPro ='" & rst.Fields("ItemPro") & "' "
                        Cn.Execute csql
                        
                        csql = "UPDATE  dd  SET dd.estDocImportado = 'S' FROM docventasdet dd " & _
                               "WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idSucursal = '" & glsSucursal & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                                 "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "' AND dd.CantidadImp >= dd.Cantidad   AND dd.ItemPro ='" & rst.Fields("ItemPro") & "' "
                        Cn.Execute csql
                        
                    Else
                        csql = "UPDATE dd SET dd.CantidadImp = dd.CantidadImp + " & CStr(rst.Fields("Cantidad")) & " FROM docventasdet dd  " & _
                               "WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idSucursal = '" & glsSucursal & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                                 "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "' "
                        Cn.Execute csql
                        
                        csql = "UPDATE dd  SET dd.estDocImportado = 'S' FROM docventasdet dd " & _
                               "WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idSucursal = '" & glsSucursal & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                                 "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "' AND dd.CantidadImp >= dd.Cantidad "
                        Cn.Execute csql
                        
                    End If
                    
                    rst.MoveNext
                Loop
            End If
        
            '--- Marcamos la Cabecera
            csql = "SELECT DISTINCT idDocumentoImp,idDocVentasImp,idSerieImp  FROM docventasdet " & _
                   "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strValCod & "' AND idSerie = '" & strSerie & "' AND idDocVentasImp <> '' AND idDocVentasImp is not null"
            
            If rstcab.State = 1 Then rstcab.Close
            rstcab.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rstcab.EOF Then
                rstcab.MoveFirst
                Do While Not rstcab.EOF
                    csql = "UPDATE C  SET c.estDocImportado = 'S' FROM docVentas c,docVentasDet d " & _
                           "WHERE c.idEmpresa = d.idEmpresa AND c.idSucursal = d.idSucursal AND c.idDocumento = d.idDocumento AND c.idDocVentas = d.idDocVentas AND c.idSerie = d.idSerie " & _
                           "AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & rstcab.Fields("idDocumentoImp") & "' AND d.idDocVentas = '" & rstcab.Fields("idDocVentasImp") & "' AND d.idSerie = '" & rstcab.Fields("idSerieImp") & "'" & _
                           "AND (SELECT count(x.idDocVentas) FROM docVentasDet x WHERE x.idEmpresa = d.idEmpresa AND x.idSucursal = d.idSucursal AND x.idDocumento = d.idDocumento AND x.idDocVentas = d.idDocVentas AND x.idSerie = d.idSerie AND x.estDocImportado <> 'S') = 0"
                    Cn.Execute csql
                    
                    rstcab.MoveNext
                Loop
            End If
        
        Else
            '--- Graba el documento(origen) y el detalle de los documentos de referencia
            If strTD <> "07" Then
                cantidadtot = 0
                csql = "select idempresa,idsucursal,iddocumento,idserie,iddocventas,idproducto,cantidad from docventasdet " & _
                        " where idempresa = '" & glsEmpresa & "' and iddocumento = '" & strTD & "' and idserie = '" & strSerie & "'  and iddocventas = '" & strValCod & "' And idSucursal = '" & glsSucursal & "' "
                If RstDet.State = 1 Then RstDet.Close
                RstDet.Open csql, Cn, adOpenStatic, adLockReadOnly
                item = 0
                If RstDet.RecordCount > 0 Then
                    RstDet.MoveFirst
                           
                    Do While Not RstDet.EOF
                        cantidadtot = CDbl(Trim("" & RstDet.Fields("cantidad")))
                        csql = " select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferencia" & _
                                " where idempresa = '" & Trim(RstDet.Fields("idempresa")) & "' and tipodocorigen = '" & Trim(RstDet.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim(RstDet.Fields("idserie")) & "' and numdocorigen = '" & Trim(RstDet.Fields("iddocventas")) & "' "
                        If rstdocref.State = 1 Then rstdocref.Close
                        rstdocref.Open csql, Cn, adOpenStatic, adLockReadOnly
                                            
                        Cadmysql = " delete from docreferenciadet " & _
                                   " where idProducto = '" & Trim("" & RstDet.Fields("idproducto")) & "' and idempresa = '" & Trim(RstDet.Fields("idempresa")) & "' and tipodocorigen = '" & Trim(RstDet.Fields("iddocumento")) & "' and seriedocorigen = '" & Trim(RstDet.Fields("idserie")) & "' and numdocorigen = '" & Trim(RstDet.Fields("iddocventas")) & "' "
                        Cn.Execute (Cadmysql)
                                          
                        If rstdocref.RecordCount > 0 Then
                            rstdocref.MoveFirst
                            Do While Not rstdocref.EOF
                                '(cantidad - cantidadimp)
                                csql = " select idempresa,idsucursal,iddocumento,idserie,iddocventas,idproducto,cantidad from docventasdet " & _
                                        " where idempresa = '" & Trim(rstdocref.Fields("idempresa")) & "' and iddocumento = '" & Trim(rstdocref.Fields("tipodocreferencia")) & "' and idserie = '" & Trim(rstdocref.Fields("seriedocreferencia")) & "'  and iddocventas = '" & Trim(rstdocref.Fields("numdocreferencia")) & "' And idProducto = '" & Trim("" & RstDet.Fields("idproducto")) & "'" 'and estdocimportado in ('N','')"
                                If rstdetref.State = 1 Then rstdetref.Close
                                rstdetref.Open csql, Cn, adOpenStatic, adLockReadOnly
                                
                                If rstdetref.RecordCount > 0 Then
                                rstdetref.MoveFirst
                                Do While Not rstdetref.EOF
                                    If Trim("" & RstDet.Fields("idproducto")) = Trim("" & rstdetref.Fields("idproducto")) Then
                                        If cantidadtot <> 0 Then
                                            If cantidadtot = CDbl(Trim("" & rstdetref.Fields("cantidad"))) Then
                                                item = item + 1
                                                cantidad = Format(Trim("" & rstdetref.Fields("cantidad")), "0.00")
                                                cantidadtot = cantidadtot - CDbl(Trim("" & rstdetref.Fields("cantidad")))
                                                                       
                                                csql = " insert into docreferenciadet " & _
                                                       " (idEmpresa, idSucursal, TipoDocOrigen, SerieDocOrigen, NumDocOrigen, TipoDocReferencia, " & _
                                                       " SerieDocReferencia, NumDocReferencia, item, idProducto, Cantidad) Values " & _
                                                       " ('" & glsEmpresa & "','" & Trim(RstDet.Fields("idsucursal")) & "','" & strTD & "','" & strSerie & "','" & strValCod & "','" & Trim(rstdetref.Fields("iddocumento")) & "', " & _
                                                       "  '" & Trim(rstdetref.Fields("idserie")) & "','" & Trim(rstdetref.Fields("iddocventas")) & "','" & item & "','" & Trim("" & RstDet.Fields("idproducto")) & "','" & cantidad & "' )"
                                                Cn.Execute csql
                                                                            
                                            ElseIf cantidadtot > CDbl(Trim("" & rstdetref.Fields("cantidad"))) Then
                                                item = item + 1
                                                cantidad = Format(CDbl(Trim("" & rstdetref.Fields("cantidad"))), "0.00")
                                                cantidadtot = cantidadtot - CDbl(Trim("" & rstdetref.Fields("cantidad")))
                                                
                                                csql = " insert into docreferenciadet " & _
                                                       " (idEmpresa, idSucursal, TipoDocOrigen, SerieDocOrigen, NumDocOrigen, TipoDocReferencia, " & _
                                                       " SerieDocReferencia, NumDocReferencia, item, idProducto, Cantidad) Values " & _
                                                       " ('" & glsEmpresa & "','" & Trim(RstDet.Fields("idsucursal")) & "','" & strTD & "','" & strSerie & "','" & strValCod & "','" & Trim(rstdetref.Fields("iddocumento")) & "', " & _
                                                       "  '" & Trim(rstdetref.Fields("idserie")) & "','" & Trim(rstdetref.Fields("iddocventas")) & "','" & item & "','" & Trim("" & RstDet.Fields("idproducto")) & "','" & cantidad & "' )"
                                                Cn.Execute csql
                                                                              
                                            ElseIf CDbl(Trim("" & rstdetref.Fields("cantidad"))) > cantidadtot Then
                                                item = item + 1
                                                cantidad = cantidadtot
                                                cantidadtot = 0
                                                csql = " insert into docreferenciadet " & _
                                                       " (idEmpresa, idSucursal, TipoDocOrigen, SerieDocOrigen, NumDocOrigen, TipoDocReferencia, " & _
                                                       " SerieDocReferencia, NumDocReferencia, item, idProducto, Cantidad) Values " & _
                                                       " ('" & glsEmpresa & "','" & Trim(RstDet.Fields("idsucursal")) & "','" & strTD & "','" & strSerie & "','" & strValCod & "','" & Trim(rstdetref.Fields("iddocumento")) & "', " & _
                                                       "  '" & Trim(rstdetref.Fields("idserie")) & "','" & Trim(rstdetref.Fields("iddocventas")) & "','" & item & "','" & Trim("" & RstDet.Fields("idproducto")) & "','" & cantidad & "' )"
                                                Cn.Execute csql
                                            End If
                                        End If
                                    End If
                                    rstdetref.MoveNext
                                Loop
                            End If
                            rstdocref.MoveNext
                        Loop
                    End If
                    RstDet.MoveNext
                Loop
            End If
            If rstdocref.State = 1 Then rstdocref.Close: Set rstdocref = Nothing
                       
            '--- Actualizamos los producto importados (detalle) docventasdet
            csql = "select idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia,idproducto,cantidad from docreferenciadet " & _
                    " where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & strTD & "' and seriedocorigen = '" & strSerie & "'  and numdocorigen = '" & strValCod & "' "
            If rstdocref.State = 1 Then rstdocref.Close
            rstdocref.Open csql, Cn, adOpenStatic, adLockReadOnly
            
            If rstdocref.RecordCount > 0 Then
                rstdocref.MoveFirst
                Do While Not rstdocref.EOF
                    csql = " Update docventasdet  set  cantidadimp = cantidadimp + " & CStr(rstdocref.Fields("Cantidad")) & "  " & _
                            " where idempresa = '" & Trim(rstdocref.Fields("idempresa")) & "' and iddocumento = '" & Trim(rstdocref.Fields("tipodocreferencia")) & "' and idserie = '" & Trim(rstdocref.Fields("seriedocreferencia")) & "'  and iddocventas = '" & Trim(rstdocref.Fields("numdocreferencia")) & "' and idproducto = '" & Trim(rstdocref.Fields("idproducto")) & "' "
                    Cn.Execute csql
                        
                    cantidad1 = CDbl(traerCampo("docventasdet", "cantidad", "idproducto", Trim(rstdocref.Fields("idproducto")), True, " iddocventas = '" & Trim(rstdocref.Fields("numdocreferencia") & "' and iddocumento = '" & Trim(rstdocref.Fields("tipodocreferencia")) & "'  and idserie = '" & Trim(rstdocref.Fields("seriedocreferencia")) & "' ")))
                    cantidad2 = CDbl(traerCampo("docventasdet", "cantidadimp", "idproducto", Trim(rstdocref.Fields("idproducto")), True, " iddocventas = '" & Trim(rstdocref.Fields("numdocreferencia") & "' and iddocumento = '" & Trim(rstdocref.Fields("tipodocreferencia")) & "'  and idserie = '" & Trim(rstdocref.Fields("seriedocreferencia")) & "' ")))
                        
                    If cantidad1 = cantidad2 Then
                        csql = " Update docventasdet set  estdocimportado = 'S'   " & _
                                " where idempresa = '" & Trim(rstdocref.Fields("idempresa")) & "' and iddocumento = '" & Trim(rstdocref.Fields("tipodocreferencia")) & "' and idserie = '" & Trim(rstdocref.Fields("seriedocreferencia")) & "'  and iddocventas = '" & Trim(rstdocref.Fields("numdocreferencia")) & "' and idproducto = '" & Trim(rstdocref.Fields("idproducto")) & "'  "
                        Cn.Execute csql
                    ElseIf cantidad1 > cantidad2 Then
                        csql = " Update docventasdet set estdocimportado = 'N'  " & _
                                " where idempresa = '" & Trim(rstdocref.Fields("idempresa")) & "' and iddocumento = '" & Trim(rstdocref.Fields("tipodocreferencia")) & "' and idserie = '" & Trim(rstdocref.Fields("seriedocreferencia")) & "'  and iddocventas = '" & Trim(rstdocref.Fields("numdocreferencia")) & "' and idproducto = '" & Trim(rstdocref.Fields("idproducto")) & "'  "
                        Cn.Execute csql
                    End If
                    rstdocref.MoveNext
                Loop
            End If
            
            '--- Actualizamos la cabecera docventas
              
            csql = "select distinct idempresa,tipodocreferencia,seriedocreferencia,numdocreferencia from docreferenciadet " & _
                    " where idempresa = '" & glsEmpresa & "' and tipodocorigen = '" & strTD & "' and seriedocorigen = '" & strSerie & "'  and numdocorigen = '" & strValCod & "' "
            If rst.State = 1 Then rst.Close
            rst.Open csql, Cn, adOpenStatic, adLockReadOnly
                        
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Do While Not rst.EOF
                    csql = "select distinct estdocimportado from docventasdet " & _
                            " where idempresa = '" & Trim(rst.Fields("idempresa")) & "' and iddocumento = '" & Trim(rst.Fields("tipodocreferencia")) & "' and idserie = '" & Trim(rst.Fields("seriedocreferencia")) & "'  and iddocventas = '" & Trim(rst.Fields("numdocreferencia")) & "'"
                    If rsh.State = 1 Then rsh.Close
                    rsh.Open csql, Cn, adOpenStatic, adLockReadOnly
                    If rsh.RecordCount > 1 Then
                        csql = " Update docventas set  estdocimportado = 'N'   " & _
                                " where idempresa = '" & Trim(rst.Fields("idempresa")) & "' and iddocumento = '" & Trim(rst.Fields("tipodocreferencia")) & "' and idserie = '" & Trim(rst.Fields("seriedocreferencia")) & "'  and iddocventas = '" & Trim(rst.Fields("numdocreferencia")) & "' "
                        Cn.Execute csql
                    
                    ElseIf rsh.RecordCount = 1 And (Trim("" & rsh.Fields("estdocimportado"))) = "N" Then
                        csql = " Update docventas set  estdocimportado = 'N'   " & _
                                " where idempresa = '" & Trim(rst.Fields("idempresa")) & "' and iddocumento = '" & Trim(rst.Fields("tipodocreferencia")) & "' and idserie = '" & Trim(rst.Fields("seriedocreferencia")) & "'  and iddocventas = '" & Trim(rst.Fields("numdocreferencia")) & "'  "
                        Cn.Execute csql
                    ElseIf rsh.RecordCount = 1 And (Trim("" & rsh.Fields("estdocimportado"))) = "S" Then
                        csql = " Update docventas set  estdocimportado = 'S'   " & _
                                " where idempresa = '" & Trim(rst.Fields("idempresa")) & "' and iddocumento = '" & Trim(rst.Fields("tipodocreferencia")) & "' and idserie = '" & Trim(rst.Fields("seriedocreferencia")) & "'  and iddocventas = '" & Trim(rst.Fields("numdocreferencia")) & "' "
                        Cn.Execute csql
                    End If
                    rst.MoveNext
                Loop
            End If
        End If
    End If
    
    '--- Liberando la separacion
    If D.Columns.ColumnByFieldName("idDocumento").Value = "40" Then
        csql = "Select d.idAlmacen, dd.idProducto,d.idEmpresa, d.idSucursal,dd.cantidad From DocventasDet dd " & _
               "INNER JOIN Docventas d On d.idempresa=dd.idempresa  and d.idsucursal = dd.idsucursal and d.idDocventas = dd.idDocventas and d.idserie =  dd.idserie " & _
               "WHERE d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & D.Columns.ColumnByFieldName("idDocumento").Value & "' AND d.idDocVentas = '" & D.Columns.ColumnByFieldName("idDocVentas").Value & "' AND d.idSerie = '" & D.Columns.ColumnByFieldName("idSerie").Value & "' "
        If rstsepar.State = 1 Then rst.Close
        rstsepar.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rstsepar.EOF Then
            rstsepar.MoveFirst
            Do While Not rstsepar.EOF
                If (Val(Format(traerCampo("productosalmacen", "Separacion", "idProducto", Trim("" & rstsepar.Fields("idProducto").Value), True, " idAlmacen = '" & rstsepar.Fields("idAlmacen").Value & "' and idSucursal = '" & glsSucursal & "' "), "0.00"))) >= Val(Format("" & rstsepar.Fields("Cantidad").Value, "0.00")) Then
                    csql = "Update ProductosAlmacen Set Separacion =( Separacion - '" & rstsepar.Fields("cantidad").Value & "') " & _
                          "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' " & _
                          "AND idProducto = '" & rstsepar.Fields("idProducto") & "' AND idAlmacen = '" & rstsepar.Fields("idAlmacen") & "' "
                    Cn.Execute csql
                End If
                rstsepar.MoveNext
            Loop
        End If
    End If
      
    obsdocventas = "" & glsobservacionventas
    csql = "update docventas set obsdocventas = '" & obsdocventas & "',obsdocventas2 = '" & ("" & glsobservacionventas2) & "' " & _
            "WHERE idEmpresa = '" & glsEmpresa & "'" & _
            "AND idSucursal = '" & glsSucursal & "'" & _
            "AND idDocumento = '" & strTD & "'" & _
            "AND idDocVentas = '" & strValCod & "'" & _
            "AND idSerie = '" & strSerie & "' "
    Cn.Execute csql
    
    '--- Acumulamos el peso del detalle del documento
    csql = "SELECT SUM(idTallaPeso) as totalTallaPeso From DocventasDet WHERE idEmpresa = '" & glsEmpresa & "'" & _
           "AND idSucursal = '" & glsSucursal & "'" & _
           "AND idDocumento = '" & strTD & "'" & _
           "AND idDocVentas = '" & strValCod & "'" & _
           "AND idSerie = '" & strSerie & "' "
    If rstUpd.State = 1 Then rstUpd.Close
    rstUpd.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUpd.EOF Then
        cadUpd = "Update docventas Set totalTallaPeso = '" & rstUpd.Fields("totalTallaPeso").Value & "' WHERE idEmpresa = '" & glsEmpresa & "'" & _
                "AND idSucursal = '" & glsSucursal & "'" & _
                "AND idDocumento = '" & strTD & "'" & _
                "AND idDocVentas = '" & strValCod & "'" & _
                "AND idSerie = '" & strSerie & "' "
    End If
    Cn.Execute (cadUpd)
    
    If traerCampo("Parametros", "ValParametro", "GlsParametro", "EXTRAE_PRECIO_ULTIMA_VENTA", True) = "S" Then
        csql = "Call Spu_InsertProductosxCliente('" & glsEmpresa & "','" & glsSucursal & "','" & strTD & "','" & strSerie & "','" & strValCod & "','','','',0,'','','I')"
        Cn.Execute (csql)
    End If
    
    If rstUpd.State = 1 Then rstUpd.Close: Set rstUpd = Nothing
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If rstsepar.State = 1 Then rstsepar.Close: Set rstsepar = Nothing
    If rstcab.State = 1 Then rstcab.Close: Set rstcab = Nothing
        
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Public Sub EjecutaSQLFormDocVentasFactura(F As Form, ByVal tipoOperacion As Integer, ByRef StrMsgError As String, ByVal strNum As String, ByVal strSerie As String, ByVal g As dxDBGrid, NArrTotales() As Double)
On Error GoTo Err
Dim C As Object
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strCodMotivoTranslado As String
Dim strCodCliente As String
Dim strCodVendedorCampo As String
Dim i As Integer

    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                
                If UCase(strCampo) = UCase("idMotivoTraslado") Then
                    strCodMotivoTranslado = C.Value
                End If
                
                If UCase(strCampo) = UCase("idPerVendedorCampo") Then
                    strCodVendedorCampo = C.Value
                End If
                
                If UCase(strCampo) = UCase("idPerCliente") Then
                    strCodCliente = C.Value
                End If
                
                If UCase(strCampo) = UCase("idDocVentas") Or UCase(strCampo) = UCase("idSerie") Then
                Else
                    Select Case tipoOperacion
                        Case 0 'inserta
                            strCampos = strCampos & strCampo & ","
                            
                            Select Case strTipoDato
                                Case "N"
                                    strValores = strValores & Val(C.Value) & ","
                                Case "T"
                                    strValores = strValores & "'" & Trim(C.Value) & "',"
                                Case "F"
                                    strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                            End Select
                        Case 1
                            If UCase(strCampo) <> UCase("idDocVentas") Or UCase(strCampo) <> UCase("idSerie") Then
                                Select Case strTipoDato
                                    Case "N"
                                        strValores = Val(C.Value)
                                    Case "T"
                                        strValores = "'" & C.Value & "'"
                                    Case "F"
                                        strValores = "'" & Format(C.Value, "yyyy-mm-dd") & "'"
                                End Select
                                strCampos = strCampos & strCampo & "=" & strValores & ","
                            End If
                    End Select
                End If
            End If
        End If
    Next
    
    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
    
    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            csql = "INSERT INTO docventas(" & strCampos & ",idDocumento,idSerie,idDocVentas,idEmpresa,idSucursal,FecRegistro,idUsuarioReg,TotalValorVenta," & _
                   "TotalIGVVenta,TotalIvap,TotalDsctoVV,TotalDsctoPV,TotalBaseImponible,TotalExonerado)VALUES(" & _
                   "" & strValores & ",'01','" & strSerie & "','" & strNum & "','" & glsEmpresa & "','" & glsSucursal & "',sysdate(),'" & glsUser & "'," & _
                   "" & NArrTotales(0) & "," & NArrTotales(1) & "," & NArrTotales(2) & "," & NArrTotales(3) & "," & NArrTotales(4) & "," & NArrTotales(5) & "," & _
                   "" & NArrTotales(6) & ")"
                   
        Case 1
            csql = "UPDATE docventas SET " & strCampos & ",TotalValorVenta = " & NArrTotales(0) & ",TotalIGVVenta = " & NArrTotales(1) & "," & _
                   "TotalIvap = " & NArrTotales(2) & ",TotalDsctoVV = " & NArrTotales(3) & ",TotalDsctoPV = " & NArrTotales(4) & "," & _
                   "TotalBaseImponible = " & NArrTotales(5) & ",TotalExonerado = " & NArrTotales(6) & " " & _
                   "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '01' AND idDocVentas = '" & strNum & "' AND idSerie = '" & strSerie & "'"
    End Select
    
    '--- Graba controles
    Cn.Execute csql
    
    '--- Grabando Grilla detalle
    If TypeName(g) <> "Nothing" Then
        Cn.Execute "DELETE FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '01' AND idDocVentas = '" & strNum & "' AND idSerie = '" & strSerie & "'"
        
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
                            strValores = strValores & Val(g.Columns(i).Value) & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO docventasdet(" & strCampos & ",idDocumento,idDocVentas,idSerie,idEmpresa,idSucursal) VALUES(" & strValores & ",'01','" & strNum & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "')"
            Cn.Execute csql
            
            g.Dataset.Next
        Loop
    End If
    
    '--- Actualizamos el codigo del vendedor de campo al cliente
    If strCodVendedorCampo <> "" Then
        csql = "UPDATE clientes SET idVendedorCampo = '" & strCodVendedorCampo & "' WHERE idCliente = '" & strCodCliente & "' AND idEmpresa = '" & glsEmpresa & "'"
        Cn.Execute csql
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub EjecutaSQLFormPagosDocVentas(F As Form, ByRef StrMsgError As String, strTD As String, strNumDoc As String, strSerie As String, strCodMovCaja As String, g As dxDBGrid, v As dxDBGrid, ByVal strCodCliente As String)
On Error GoTo Err
Dim csql As String
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strCodMovDet As String
Dim strCodMoneda As String
Dim strCodFormadePago As String
Dim strFormasPago As String
Dim strFecVactos  As String
Dim dblValMonto As Double
Dim i As Integer
    
    '--- Grabando Grilla
    If TypeName(g) <> "Nothing" Then
        '--- Eliminamos pagosdocventas
        Cn.Execute "DELETE FROM pagosdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
        
        '--- Eliminamos movcajasdet
        Cn.Execute "DELETE FROM movcajasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
        
        g.Dataset.First
        Do While Not g.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To g.Columns.Count - 1
                If UCase(left(g.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(g.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(g.Columns(i).ObjectName, 3)
                    
                    strCampos = strCampos & strCampo & ","
                    
                    If strCampo = "idMoneda" Then
                        strCodMoneda = Trim(g.Columns(i).Value)
                    End If
                    
                    If strCampo = "idFormadePago" Then
                        strCodFormadePago = Trim(g.Columns(i).Value)
                    End If
                    
                    If strCampo = "MontoOri" Then
                        dblValMonto = g.Columns(i).Value
                    End If
                    
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
            
            csql = "INSERT INTO pagosdocventas(" & strCampos & ",idDocumento,idDocVentas,idSerie,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strTD & "','" & strNumDoc & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "')"
            Cn.Execute csql
            
            '--- Insertamos los montos en movimientos de caja
            strCodMovDet = GeneraCorrelativoAnoMesNuevo("movcajasdet", "idMovCajaDet", Year(CVDate(getFechaSistema)) & Format(Month(CVDate(getFechaSistema)), "00"), True)
            
            csql = "INSERT INTO movcajasdet (idMovCajaDet,idMovCaja,idTipoMovCaja,idMoneda,ValMonto,FecRegistro,idEmpresa,idSucursal,ValTipoCambio,idFormadePago,idDocumento,idDocVentas,idSerie) VALUES(" & _
                   "'" & strCodMovDet & "','" & strCodMovCaja & "','99990002','" & strCodMoneda & "'," & dblValMonto & ",sysdate(),'" & glsEmpresa & "','" & glsSucursal & "'," & glsTC & ",'" & strCodFormadePago & "','" & strTD & "','" & strNumDoc & "','" & strSerie & "')"
            Cn.Execute csql
            
            g.Dataset.Next
        Loop
    End If
    
    If TypeName(v) <> "Nothing" Then
        v.Dataset.First
        Do While Not v.Dataset.EOF
            strCodMoneda = v.Columns.ColumnByFieldName("idMoneda").Value
            dblValMonto = v.Columns.ColumnByFieldName("Vuelto").Value
                       
            If dblValMonto > 0 Then
                '--- Insertamos los montos en movimientos de caja
                strCodMovDet = GeneraCorrelativoAnoMes("movcajasdet", "idMovCajaDet")
            
                csql = "INSERT INTO movcajasdet (idMovCajaDet,idMovCaja,idTipoMovCaja,idMoneda,ValMonto,FecRegistro,idEmpresa,idSucursal,ValTipoCambio,idFormadePago,idDocumento,idDocVentas,idSerie) VALUES(" & _
                       "'" & strCodMovDet & "','" & strCodMovCaja & "','99990003','" & strCodMoneda & "'," & dblValMonto & ",sysdate(),'" & glsEmpresa & "','" & glsSucursal & "'," & glsTC & ",'','" & strTD & "','" & strNumDoc & "','" & strSerie & "')"
                Cn.Execute csql
            End If
            v.Dataset.Next
        Loop
    End If
    
    '--- Actualizamos el Documento de venta
    g.Dataset.First
    If Not g.Dataset.EOF Then
        If Trim(g.Columns.ColumnByFieldName("glsFormadePago").Value) <> "" Then
            strFormasPago = strFormasPago + g.Columns.ColumnByFieldName("glsFormadePago").Value & ","
        End If
        If Trim(g.Columns.ColumnByFieldName("fecVctos").Value) <> "" Then
            strFecVactos = strFecVactos + g.Columns.ColumnByFieldName("fecVctos").Value & ","
        End If
    End If
    
    If Len(Trim(strFormasPago)) > 1 Then
        strFormasPago = left(strFormasPago, Len(strFormasPago) - 1)
    End If
    
    If Len(Trim(strFecVactos)) > 1 Then
        strFecVactos = left(strFecVactos, Len(strFecVactos) - 1)
    End If
    
    csql = "UPDATE DOCVENTAS SET GlsFormasPago = '" & strFormasPago & "', GlsFecVectos = '" & strFecVactos & "', FecPago = '" & Format(strFecVactos, "yyyy-mm-dd") & "', estDocVentas = 'CAN', idMovCaja = '" & strCodMovCaja & "' " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND  idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    Cn.Execute csql
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Public Sub generaValeDocVentas(strIdDocumento As String, strIdSerie As String, StrTipVale As String, strIdDocventas As String, ByVal strCodMotivoTraslado As String, ByVal StrFecEmision As String, ByVal strCod_CentroCosto As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rs As New ADODB.Recordset
Dim strCod              As String
Dim strTipoVale         As String
Dim strConcepto         As String
Dim strAbreviaturaDoc   As String
Dim indMod              As Boolean
Dim Cadmysql            As String
Dim CInsertCab          As String
Dim CInsertDet          As String
Dim CIdComprobanteAux   As String
Dim Sw_Entro            As Boolean
Dim rst                 As New ADODB.Recordset
Dim strVVUnit, strIGVUnit, strPVUnit, strTotalVVNeto, strTotalIGVNeto, strTotalPVNeto As Double
Dim strValorTotal       As Double
Dim strIgvTotal         As Double
Dim strPrecioTotal      As Double
Dim strNewPeriodoinv    As String

    strValorTotal = 0: strIgvTotal = 0: strPrecioTotal = 0
    strVVUnit = 0: strIGVUnit = 0: strPVUnit = 0: strTotalVVNeto = 0: strTotalIGVNeto = 0: strTotalPVNeto = 0
    csql = ""
    indMod = False
    
    If Len(Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", glsSucursal, True, " year(FecInicio) = " & Format(StrFecEmision, "yyyy") & " "))) = 0 Then
        strNewPeriodoinv = glsCodPeriodoINV
    Else
        strNewPeriodoinv = Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", glsSucursal, True, " year(FecInicio) = " & Format(StrFecEmision, "yyyy") & " "))
    End If
    
    '--- VERIFICAMOS EL ESTADO DEL DOCUMENTO
    csql = "SELECT indVale,idValesCab FROM docventas " & _
           "WHERE idDocumento = '" & strIdDocumento & "'" & _
            " AND idSerie     = '" & strIdSerie & "'" & _
            " AND idDocVentas = '" & strIdDocventas & "'" & _
            " AND idEmpresa = '" & glsEmpresa & "'" & _
            " AND idSucursal = '" & glsSucursal & "' "
    If rs.State = 1 Then rs.Close
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.Fields("indVale") = "S" Then
        indMod = True
        
        '--- SELECCIONAMOS EL NUMERO DEL VALE DEL DOCUMENTO
        strCod = Trim("" & rs.Fields("idValesCab"))
        
        '--- ACTUALIZAMOS STOCK EN LINEA
        actualizaStock strCod, 1, StrMsgError, StrTipVale, False
        If StrMsgError <> "" Then GoTo Err
        
        '--- ELIMINAMOS EL VALE YA INGRESADO
        csql = "DELETE FROM valesDet WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoVale = '" & StrTipVale & "' "
        Cn.Execute csql
        
        csql = "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = '" & IIf(StrTipVale = "I", "88", "99") & "' AND numDocOrigen = '" & strCod & "' AND serieDocOrigen = '000' "
        Cn.Execute csql
        
        csql = "DELETE FROM valesCab WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoVale = '" & StrTipVale & "'"
        Cn.Execute csql
        
    Else
        indMod = False
    End If
    
'Luis 2406/2017
'    Select Case strIdDocumento
'        Case "01", "03", "12", "90", "91", "08"
'            strTipoVale = "S"
'            'strConcepto = "22"
'            strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
'            If strConcepto = "" Then
'                strConcepto = "22"
'            End If
'
'        Case "86"
'            strTipoVale = "S"
'            If strTipoVale = "" Then strTipoVale = "S"
'            strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
'        Case "07", "89"
'            strTipoVale = "I"
'            'strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
'            strConcepto = Trim("" & traerCampo("MotivosNCD", "IdConcepto", "IdMotivoNCD", strCodMotivoTraslado, False))
'    End Select
    
    Select Case strIdDocumento
        Case "01", "03", "12", "90", "91", "08"

            strTipoVale = ""
            strTipoVale = Trim("" & traerCampo("motivostraslados", "indtipovale", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
            If strTipoVale = "" Then strTipoVale = "S"
            
            strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
            If strConcepto = "" Then
                strConcepto = "22"
            End If
            
        Case "86"
            
            strTipoVale = ""
            strTipoVale = Trim("" & traerCampo("motivostraslados", "indtipovale", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
            If strTipoVale = "" Then strTipoVale = "S"
            
            strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
        Case "07", "89"
            strTipoVale = "I"
            'strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
            strConcepto = Trim("" & traerCampo("MotivosNCD", "IdConcepto", "IdMotivoNCD", strCodMotivoTraslado, False))
    End Select
    
    If indMod = False Then
        strCod = generaCorrelativoAnoMes_Vale("ValesCab", "idValesCab", strTipoVale)
    End If
    strAbreviaturaDoc = "" & traerCampo("documentos", "AbreDocumento", "idDocumento", strIdDocumento, False)
    
    '---------------------------------------------------------------------------------------------------------------------
    Cadmysql = "Select  d.idDocumento,d.idSerie,d.idDocventas,d.idEmpresa,d.idSucursal,Concat(dc.AbreDocumento,' ',d.idSerie,'-',d.idDocventas) as GlsDocreferencia,d.FecEmision, " & _
           "idPercliente , idMotivoTraslado, idAlmacen, obsdocventas, D.idMoneda, D.TipoCambio, dd.idProducto, dd.GlsProducto, dd.idUM, dd.Factor, dd.Afecto, dd.cantidad, " & _
           "dd.item,dd.NumLote, dd.FecVencProd,d.idCentroCosto,dd.vvunit " & _
           "From Docventas d " & _
           "Inner Join Documentos dc   On d.idDocumento =  dc.idDocumento " & _
           "Inner Join DocventasDet dd   On d.idDocumento = dd.idDocumento   And d.idDocventas = dd.idDocventas   And d.idSerie = dd.idSerie   And d.idEmpresa = dd.idEmpresa  And d.idSucursal = dd.idSucursal " & _
           "Inner Join Productos p On dd.idProducto = p.idProducto And dd.idEmpresa = p.idEmpresa " & _
           "Where d.idEmpresa = '" & glsEmpresa & "' And d.idSucursal = '" & glsSucursal & "' And d.idDocumento = '" & strIdDocumento & "' And d.idDocventas = '" & strIdDocventas & "' And d.idSerie = '" & strIdSerie & "' And p.idTipoProducto <> '06002'   Order By d.idDocumento,d.idDocventas,d.idSerie "
    
    rst.Open Cadmysql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        CInsertDet = "Insert Into Valesdet (tipoVale,idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
                        "TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal,NumLote,FecVencProd) Values "
                
        CInsertCab = "Insert Into ValesCab (idValesCab,tipoVale,fechaEmision, valorTotal, igvTotal,precioTotal, idProvCliente,idConcepto, idAlmacen, idMoneda," & _
        "GlsDocReferencia,TipoCambio,idEmpresa,idSucursal,idPeriodoInv,idCentroCosto,FechaRegistro,IdUsuarioRegistro) Values "

        Do While Not rst.EOF
            If CIdComprobanteAux <> rst.Fields("idDocventas") Then
                 CIdComprobanteAux = rst.Fields("idDocventas")
            End If
            
'            strVVUnit = Format(traerCostoUnit(rst.Fields("idProducto").Value, rst.Fields("idAlmacen").Value, rst.Fields("idMoneda").Value, StrMsgError), "0.00000")
'            strIGVUnit = Format((Val(strVVUnit) * glsIGV), "0.00000")
'            strPVUnit = Format(Val(strVVUnit) + Val(strIGVUnit), "0.00000")
            
            If strTipoVale = "I" Then
            
                strVVUnit = Format(rst.Fields("VVUNIT").Value, "0.00000")
                strIGVUnit = Format((Val(strVVUnit) * glsIGV), "0.00000")
                strPVUnit = Format(Val(strVVUnit) + Val(strIGVUnit), "0.00000")
            
            Else

                strVVUnit = 0 'Format(traerCostoUnit(rst.Fields("idProducto").Value, rst.Fields("idAlmacen").Value, rst.Fields("idMoneda").Value, StrMsgError), "0.00000")
                strIGVUnit = 0 'Format((Val(strVVUnit) * glsIGV), "0.00000")
                strPVUnit = 0 'Format(Val(strVVUnit) + Val(strIGVUnit), "0.00000")
            
            End If
            
            
            strTotalVVNeto = Format((Val(strVVUnit) * rst.Fields("Cantidad").Value), "0.00")
            strTotalIGVNeto = Format((Val(strTotalVVNeto) * glsIGV), "0.00")
            strTotalPVNeto = Format(Val(strTotalVVNeto) + Val(strTotalIGVNeto), "0.00")
            strValorTotal = Format(Val(strValorTotal) + Val(strTotalVVNeto), "0.00")
            strIgvTotal = Format(Val(strIgvTotal) + (Val(strValorTotal) * glsIGV), "0.00")
            strPrecioTotal = Format((Val(strPrecioTotal) + (Val(strValorTotal) + Val(strIgvTotal))), "0.00")
            
            CInsertDet = CInsertDet & " ('" & strTipoVale & "','" & strCod & "','" & rst.Fields("item").Value & "','" & rst.Fields("idProducto").Value & "','" & rst.Fields("GlsProducto").Value & "','" & rst.Fields("idUM").Value & "', " & _
                        " '" & rst.Fields("Factor").Value & "','" & rst.Fields("Afecto").Value & "','" & rst.Fields("Cantidad").Value & "','" & strVVUnit & "','" & strIGVUnit & "','" & strPVUnit & "','" & strTotalVVNeto & "', " & _
                        " '" & strTotalIGVNeto & "','" & strTotalPVNeto & "','" & rst.Fields("idMoneda").Value & "','" & rst.Fields("idEmpresa").Value & "','" & rst.Fields("idSucursal").Value & "','" & rst.Fields("NumLote").Value & "','" & rst.Fields("FecVencProd").Value & "' ),"
            rst.MoveNext

            If rst.EOF Then
                rst.MovePrevious
                Sw_Entro = True
            Else
                If CIdComprobanteAux <> rst.Fields("idDocventas") Then
                    rst.MovePrevious
                    Sw_Entro = True
                End If
            End If
              
            If Sw_Entro Then
                CInsertCab = CInsertCab & "('" & strCod & "','" & strTipoVale & "','" & Format(rst.Fields("FecEmision").Value, "YYYY-MM-DD") & "','" & strValorTotal & "','" & strIgvTotal & "','" & strPrecioTotal & "','" & rst.Fields("idPerCliente").Value & "', " & _
                       " '" & strConcepto & "','" & rst.Fields("idAlmacen").Value & "','" & rst.Fields("idMoneda").Value & "','" & rst.Fields("GlsDocreferencia").Value & "', " & _
                       " '" & rst.Fields("TipoCambio").Value & "','" & rst.Fields("idEmpresa").Value & "','" & rst.Fields("idSucursal").Value & "'," & _
                       "'" & strNewPeriodoinv & "','" & rst.Fields("idCentroCosto").Value & "',getdate(),'" & glsUser & "') "
                
                rst.MoveNext
                Sw_Entro = False
                strValorTotal = 0: strIgvTotal = 0: strPrecioTotal = 0
            End If
            strVVUnit = 0: strIGVUnit = 0: strPVUnit = 0: strTotalVVNeto = 0: strTotalIGVNeto = 0: strTotalPVNeto = 0
        Loop
        Cn.Execute (CInsertCab)
        Cn.Execute (left(CInsertDet, Len(CInsertDet) - 1))
    End If
    '---------------------------------------------------------------------------------------------------------------------
    If Len(Trim("" & traerCampo("ValesDet", "IdValesCab", "IdSucursal", glsSucursal, True, "TipoVale = '" & strTipoVale & "' And IdValesCab = '" & strCod & "'"))) = 0 Then
        csql = "Delete From ValesCab " & _
               "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And TipoVale = '" & strTipoVale & "' " & _
               "And IdValesCab = '" & strCod & "'"
        Cn.Execute csql
        
        csql = "Update Docventas Set TipoVale = '',IdValesCab = '',IndVale = 'N' " & _
               "Where IdDocumento = '" & strIdDocumento & "' And IdSerie = '" & strIdSerie & "' And IdDocVentas = '" & strIdDocventas & "' " & _
               "And IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "'"
        Cn.Execute csql
        
    Else
        '--- INSERTAMOS DOCUMENTO DE REFERENCIA
        If strTipoVale = "I" Then
            csql = "INSERT INTO docreferencia (tipoDocOrigen,numDocOrigen,serieDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,item,idEmpresa,idSucursal) " & _
                   "VALUES ('88','" & strCod & "','000','" & strIdDocumento & "','" & strIdDocventas & "','" & strIdSerie & "',1,'" & glsEmpresa & "','" & glsSucursal & "')"
        Else
            csql = "INSERT INTO docreferencia (tipoDocOrigen,numDocOrigen,serieDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,item,idEmpresa,idSucursal) " & _
                   "VALUES ('99','" & strCod & "','000','" & strIdDocumento & "','" & strIdDocventas & "','" & strIdSerie & "',1,'" & glsEmpresa & "','" & glsSucursal & "')"
        End If
        Cn.Execute csql
            
        '--- ACTUALIZAMOS EL ESTADO DEL DOCUMENTO
        csql = "UPDATE docventas SET tipoVale = '" & strTipoVale & "',idValesCab = '" & strCod & "',indVale = 'S' " & _
               "WHERE idDocumento = '" & strIdDocumento & "'" & _
                " AND idSerie     = '" & strIdSerie & "'" & _
                " AND idDocVentas = '" & strIdDocventas & "'" & _
                " AND idEmpresa = '" & glsEmpresa & "'" & _
                " AND idSucursal = '" & glsSucursal & "'"
        Cn.Execute csql
        
        '--- ACTUALIZAMOS STOCK EN LINEA
        actualizaStock strCod, 0, StrMsgError, strTipoVale, False
        If StrMsgError <> "" Then GoTo Err
    End If
    Set rs = Nothing
    
Exit Sub
Err:
    Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub generaValeDocVentasTransfIngreso(strIdDocumento As String, strIdSerie As String, StrTipVale As String, strIdDocventas As String, ByVal strCodMotivoTraslado As String, ByVal StrFecEmision As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rs As New ADODB.Recordset
Dim strCod              As String
Dim strTipoVale         As String
Dim strConcepto         As String
Dim strAbreviaturaDoc   As String
Dim indMod              As Boolean
Dim Cadmysql            As String
Dim CInsertCab          As String
Dim CInsertDet          As String
Dim CIdComprobanteAux   As String
Dim Sw_Entro            As Boolean
Dim rst                 As New ADODB.Recordset
Dim strVVUnit, strIGVUnit, strPVUnit, strTotalVVNeto, strTotalIGVNeto, strTotalPVNeto As Double
Dim strValorTotal       As Double
Dim strIgvTotal         As Double
Dim strPrecioTotal      As Double
Dim strNewPeriodoinv    As String
Dim SucursalDestino     As String

    SucursalDestino = traerCampo("Docventas", "SucursalDestino", "iddocventas", strIdDocventas, True, " idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strIdDocumento & "'  AND idSerie = '" & strIdSerie & "'")
    
    strValorTotal = 0: strIgvTotal = 0: strPrecioTotal = 0
    strVVUnit = 0: strIGVUnit = 0: strPVUnit = 0: strTotalVVNeto = 0: strTotalIGVNeto = 0: strTotalPVNeto = 0
    csql = ""
    indMod = False
    
    If Len(Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", SucursalDestino, True, " year(FecInicio) = " & Format(StrFecEmision, "yyyy") & " "))) = 0 Then
        strNewPeriodoinv = glsCodPeriodoINV
    Else
        strNewPeriodoinv = Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", SucursalDestino, True, " year(FecInicio) = " & Format(StrFecEmision, "yyyy") & " "))
    End If
    
    '--- VERIFICAMOS EL ESTADO DEL DOCUMENTO
    csql = "SELECT ifnull(indValeing,'')as indValeing,ifnull(idValesCabIng,'') as idValesCabIng FROM docventas " & _
           "WHERE idDocumento = '" & strIdDocumento & "'" & _
            " AND idSerie     = '" & strIdSerie & "'" & _
            " AND idDocVentas = '" & strIdDocventas & "'" & _
            " AND idEmpresa = '" & glsEmpresa & "'" & _
            " AND idSucursal = '" & glsSucursal & "' "
                
    If rs.State = 1 Then rs.Close
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.Fields("indValeing") = "S" Then
        indMod = True

        '--- SELECCIONAMOS EL NUMERO DEL VALE DEL DOCUMENTO
        strCod = Trim("" & rs.Fields("idValesCabIng"))

        '--- ACTUALIZAMOS STOCK EN LINEA
        actualizaStockTransferencia strCod, 1, SucursalDestino, StrMsgError, StrTipVale, False
        If StrMsgError <> "" Then GoTo Err

        '--- ELIMINAMOS EL VALE YA INGRESADO
        csql = "DELETE FROM valesDet WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & SucursalDestino & "' AND tipoVale = '" & StrTipVale & "' "
        Cn.Execute csql

        csql = "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & SucursalDestino & "' AND tipoDocOrigen = '" & IIf(StrTipVale = "I", "88", "99") & "' AND numDocOrigen = '" & strCod & "' AND serieDocOrigen = '000' "
        Cn.Execute csql

        csql = "DELETE FROM valesCab WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & SucursalDestino & "' AND tipoVale = '" & StrTipVale & "'"
        Cn.Execute csql

    Else
        indMod = False
    End If
    
    Select Case strIdDocumento
        Case "01", "03", "12", "90", "91"
            strTipoVale = "S"
            strConcepto = "22"
        Case "86"
            strTipoVale = "I"
            If strTipoVale = "" Then strTipoVale = "I"
            strConcepto = Trim("" & traerCampo("motivostraslados", "idconceptoI", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
        Case "07", "89"
            strTipoVale = "I"
            strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
    End Select
    
    If indMod = False Then
        strCod = generaCorrelativoAnoMes_Vale("ValesCab", "idValesCab", strTipoVale)
    End If
    strAbreviaturaDoc = "" & traerCampo("documentos", "AbreDocumento", "idDocumento", strIdDocumento, False)
    
    '---------------------------------------------------------------------------------------------------------------------
    Cadmysql = "Select  d.idDocumento,d.idSerie,d.idDocventas,d.idEmpresa,d.SucursalDestino,Concat(dc.AbreDocumento,' ',d.idSerie,'-',d.idDocventas) as GlsDocreferencia,d.FecEmision, " & _
           "idPercliente , idMotivoTraslado, AlmacenDestino, obsdocventas, D.idMoneda, D.TipoCambio, dd.idProducto, dd.GlsProducto, dd.idUM, dd.Factor, dd.Afecto, dd.cantidad, dd.item,dd.NumLote, dd.FecVencProd, dd.idMoneda " & _
           "From Docventas d " & _
           "Inner Join Documentos dc   On d.idDocumento =  dc.idDocumento " & _
           "Inner Join DocventasDet dd   On d.idDocumento = dd.idDocumento   And d.idDocventas = dd.idDocventas   And d.idSerie = dd.idSerie   And d.idEmpresa = dd.idEmpresa  And d.idSucursal = dd.idSucursal " & _
           "Inner Join Productos p On dd.idProducto = p.idProducto And dd.idEmpresa = p.idEmpresa " & _
           "Where d.idEmpresa = '" & glsEmpresa & "' And d.idSucursal = '" & glsSucursal & "' And d.idDocumento = '" & strIdDocumento & "' And d.idDocventas = '" & strIdDocventas & "' And d.idSerie = '" & strIdSerie & "' And p.idTipoProducto <> '06002'   Order By d.idDocumento,d.idDocventas,d.idSerie "
    
    rst.Open Cadmysql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        CInsertDet = "Insert Into Valesdet (tipoVale,idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
                        "TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal,NumLote,FecVencProd) Values "
                
        CInsertCab = "Insert Into ValesCab (idValesCab,tipoVale,fechaEmision, valorTotal, igvTotal,precioTotal, idProvCliente,idConcepto, idAlmacen, idMoneda," & _
        "GlsDocReferencia,TipoCambio,idEmpresa,idSucursal,idPeriodoInv,FechaRegistro,IdUsuarioRegistro) Values "

        Do While Not rst.EOF
            If CIdComprobanteAux <> rst.Fields("idDocventas") Then
                 CIdComprobanteAux = rst.Fields("idDocventas")
            End If
            
            strVVUnit = Format(traerCostoUnit(rst.Fields("idProducto").Value, rst.Fields("AlmacenDestino").Value, rst.Fields("idMoneda").Value, StrMsgError), "0.00")
            strIGVUnit = Format((Val(strVVUnit) * glsIGV), "0.00")
            strPVUnit = Format(Val(strVVUnit) + Val(strIGVUnit), "0.00")
            strTotalVVNeto = Format((Val(strVVUnit) * rst.Fields("Cantidad").Value), "0.00")
            strTotalIGVNeto = Format((Val(strTotalVVNeto) * glsIGV), "0.00")
            strTotalPVNeto = Format(Val(strTotalVVNeto) + Val(strTotalIGVNeto), "0.00")
            
            strValorTotal = Format(Val(strValorTotal) + Val(strTotalVVNeto), "0.00")
            strIgvTotal = Format(Val(strIgvTotal) + (Val(strValorTotal) * glsIGV), "0.00")
            strPrecioTotal = Format((Val(strPrecioTotal) + (Val(strValorTotal) + Val(strIgvTotal))), "0.00")
            
            CInsertDet = CInsertDet & " ('" & strTipoVale & "','" & strCod & "','" & rst.Fields("item").Value & "','" & rst.Fields("idProducto").Value & "','" & rst.Fields("GlsProducto").Value & "','" & rst.Fields("idUM").Value & "', " & _
                        " '" & rst.Fields("Factor").Value & "','" & rst.Fields("Afecto").Value & "','" & rst.Fields("Cantidad").Value & "','" & strVVUnit & "','" & strIGVUnit & "','" & strPVUnit & "','" & strTotalVVNeto & "', " & _
                        " '" & strTotalIGVNeto & "','" & strTotalPVNeto & "','" & rst.Fields("idMoneda").Value & "','" & rst.Fields("idEmpresa").Value & "','" & rst.Fields("SucursalDestino").Value & "','" & rst.Fields("NumLote").Value & "','" & rst.Fields("FecVencProd").Value & "' ),"
            rst.MoveNext

            If rst.EOF Then
                rst.MovePrevious
                Sw_Entro = True
            Else
                If CIdComprobanteAux <> rst.Fields("idDocventas") Then
                    rst.MovePrevious
                    Sw_Entro = True
                End If
            End If
              
            If Sw_Entro Then
                CInsertCab = CInsertCab & "('" & strCod & "','" & strTipoVale & "','" & Format(rst.Fields("FecEmision").Value, "YYYY-MM-DD") & "','" & strValorTotal & "','" & strIgvTotal & "','" & strPrecioTotal & "','" & rst.Fields("idPerCliente").Value & "', " & _
                       " '" & strConcepto & "','" & rst.Fields("AlmacenDestino").Value & "','" & rst.Fields("idMoneda").Value & "','" & rst.Fields("GlsDocreferencia").Value & "', " & _
                       " '" & rst.Fields("TipoCambio").Value & "','" & rst.Fields("idEmpresa").Value & "','" & rst.Fields("SucursalDestino").Value & "'," & _
                       "'" & strNewPeriodoinv & "',SysDate(),'" & glsUser & "') "
                
                rst.MoveNext
                Sw_Entro = False
                strValorTotal = 0: strIgvTotal = 0: strPrecioTotal = 0
            End If
            strVVUnit = 0: strIGVUnit = 0: strPVUnit = 0: strTotalVVNeto = 0: strTotalIGVNeto = 0: strTotalPVNeto = 0
        Loop
        Cn.Execute (CInsertCab)
        Cn.Execute (left(CInsertDet, Len(CInsertDet) - 1))
    End If
    '---------------------------------------------------------------------------------------------------------------------
    If Len(Trim("" & traerCampo("ValesDet", "IdValesCab", "IdSucursal", SucursalDestino, True, "TipoVale = '" & strTipoVale & "' And IdValesCab = '" & strCod & "'"))) = 0 Then
        csql = "Delete From ValesCab " & _
               "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & SucursalDestino & "' And TipoVale = 'I' " & _
               "And IdValesCab = '" & strCod & "'"
        Cn.Execute csql
        
        csql = "Update Docventas Set TipoVale = '',IdValesCab = '',IndVale = 'N' " & _
               "Where IdDocumento = '" & strIdDocumento & "' And IdSerie = '" & strIdSerie & "' And IdDocVentas = '" & strIdDocventas & "' " & _
               "And IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & SucursalDestino & "'"
        Cn.Execute csql
        
    Else
        '--- INSERTAMOS DOCUMENTO DE REFERENCIA
        If strTipoVale = "I" Then
            csql = "INSERT INTO docreferencia (tipoDocOrigen,numDocOrigen,serieDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,item,idEmpresa,idSucursal) " & _
                   "VALUES ('88','" & strCod & "','000','" & strIdDocumento & "','" & strIdDocventas & "','" & strIdSerie & "',1,'" & glsEmpresa & "','" & SucursalDestino & "')"
        Else
            csql = "INSERT INTO docreferencia (tipoDocOrigen,numDocOrigen,serieDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,item,idEmpresa,idSucursal) " & _
                   "VALUES ('99','" & strCod & "','000','" & strIdDocumento & "','" & strIdDocventas & "','" & strIdSerie & "',1,'" & glsEmpresa & "','" & SucursalDestino & "')"
        End If
        Cn.Execute csql
            
        '--- ACTUALIZAMOS EL ESTADO DEL DOCUMENTO
        csql = "UPDATE docventas SET tipoValeIng = '" & strTipoVale & "',idValesCabIng = '" & strCod & "',indValeING = 'S' " & _
               "WHERE idDocumento = '" & strIdDocumento & "'" & _
                " AND idSerie     = '" & strIdSerie & "'" & _
                " AND idDocVentas = '" & strIdDocventas & "'" & _
                " AND idEmpresa = '" & glsEmpresa & "'" & _
                " AND idSucursal = '" & glsSucursal & "'"
        Cn.Execute csql
        
        '--- ACTUALIZAMOS STOCK EN LINEA
        '--- NUEVO PROCEDIMIENTO AGREGADO POR LA TRANSF. ENTRE ESTABLECIMIENTOS
        actualizaStockTransferencia strCod, 0, SucursalDestino, StrMsgError, strTipoVale, False
        If StrMsgError <> "" Then GoTo Err
    End If
    Set rs = Nothing
    
    Exit Sub

Err:
    Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub generaValeDocVentasFormula(strIdDocumento As String, strIdSerie As String, StrTipVale As String, strIdDocventas As String, ByVal strCodMotivoTraslado As String, ByVal StrFecEmision As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rs                      As New ADODB.Recordset
Dim strCod                  As String
Dim strTipoVale             As String
Dim strConcepto             As String
Dim strAbreviaturaDoc       As String
Dim indMod                  As Boolean
Dim strNewPeriodoinv        As String

    csql = ""
    indMod = False
    
    If Len(Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", glsSucursal, True, " year(FecInicio) = " & Format(StrFecEmision, "yyyy") & " "))) = 0 Then
        strNewPeriodoinv = glsCodPeriodoINV
    Else
        strNewPeriodoinv = Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", glsSucursal, True, " year(FecInicio) = " & Format(StrFecEmision, "yyyy") & " "))
    End If

    '--- VERIFICAMOS EL ESTADO DEL DOCUMENTO
    csql = "SELECT indVale,idValesCab FROM docventas " & _
           "WHERE idDocumento = '" & strIdDocumento & "'" & _
            " AND idSerie     = '" & strIdSerie & "'" & _
            " AND idDocVentas = '" & strIdDocventas & "'" & _
            " AND idEmpresa = '" & glsEmpresa & "'" & _
            " AND idSucursal = '" & glsSucursal & "'"
    
    If rs.State = 1 Then rs.Close
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.Fields("indVale") = "S" Then
        indMod = True
        '--- SELECCIONAMOS EL NUMERO DEL VALE DEL DOCUMENTO
        strCod = Trim("" & rs.Fields("idValesCab"))
        
        '--- ACTUALIZAMOS STOCK EN LINEA
        actualizaStock strCod, 1, StrMsgError, StrTipVale, False
        If StrMsgError <> "" Then GoTo Err
        
        '--- ELIMINAMOS EL VALE YA INGRESADO
        csql = "DELETE FROM valesDet WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoVale = '" & StrTipVale & "' "
        Cn.Execute csql
        
        csql = "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = '" & IIf(StrTipVale = "I", "88", "99") & "' AND numDocOrigen = '" & strCod & "' AND serieDocOrigen = '000' "
        Cn.Execute csql
        
        csql = "DELETE FROM valesCab WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoVale = '" & StrTipVale & "' "
        Cn.Execute csql
    Else
        indMod = False
    End If

    Select Case strIdDocumento
        Case "01", "03", "12", "90", "91"
            strTipoVale = "S"
            strConcepto = "22"
        Case "86"
            strTipoVale = "S"
            If strTipoVale = "" Then strTipoVale = "S"
            strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
        Case "07", "89"
            strTipoVale = "I"
            strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
    End Select
    
    If indMod = False Then
        strCod = generaCorrelativoAnoMes_Vale("ValesCab", "idValesCab", strTipoVale)
    End If
    strAbreviaturaDoc = "" & traerCampo("documentos", "AbreDocumento", "idDocumento", strIdDocumento, False)
    
    '--- GENERAMOS CABECERA
    csql = "INSERT INTO ValesCab (idValesCab,tipoVale,fechaEmision, valorTotal, igvTotal," & _
                       "precioTotal, idProvCliente,idConcepto, idAlmacen, idMoneda,GlsDocReferencia,TipoCambio,idEmpresa,idSucursal,idPeriodoInv,FechaRegistro,IdUsuarioRegistro) " & _
           "SELECT '" & strCod & "','" & strTipoVale & "',d.FecEmision, d.TotalValorVenta, d.TotalIGVVenta," & _
                  "d.TotalPrecioVenta, d.idPerCliente,'" & strConcepto & "', d.idAlmacen, d.idMoneda," & _
                  "'" & strAbreviaturaDoc & " " & strIdSerie & "-" & strIdDocventas & "', d.TipoCambio,idEmpresa,idSucursal,'" & strNewPeriodoinv & "'," & _
                  "SysDate(),'" & glsUser & "' " & _
           "FROM docventas d " & _
           "WHERE d.idDocumento = '" & strIdDocumento & "'" & _
            " AND d.idSerie     = '" & strIdSerie & "'" & _
            " AND d.idDocVentas = '" & strIdDocventas & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'"
    Cn.Execute csql
    
    '--- GENERAMOS DETALLE DEL VALE
    csql = "INSERT INTO valesdet (tipoVale,idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
                             "TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal,FecVencProd)" & _
            "Select '" & strTipoVale & "','" & strCod & "' As IdValesCab,(@i:=@i +1) As item,P.idProducto,P.GlsProducto,p.idUMCompra As IdUM,1 AS Factor,P.AfectoIgv As Afecto," & _
            "((d.Cantidad / d.Factor) * If(A.Cantidad Is Null,1,A.Cantidad)) AS Cantidad," & _
            "If(A.IdProducto Is Null,d.VVUnit,IfNull(V.Unit,0)) As VVUnit,If(A.IdProducto Is Null,d.IGVUnit,0) As IGVUnit," & _
            "If(A.IdProducto Is Null,d.PVUnit,IfNull(V.Unit,0)) As PVUnit," & _
            "If(A.IdProducto Is Null,d.TotalVVNeto,IfNull(V.Unit,0) * A.Cantidad) As TotalVVNeto,If(A.IdProducto Is Null,d.TotalIGVNeto,0) As TotalIGVNeto," & _
            "If(A.IdProducto Is Null,d.TotalPVNeto,IfNull(V.Unit,0) * A.Cantidad) As TotalPVNeto, d.idMoneda, d.idEmpresa, d.idSucursal, d.FecVencProd " & _
            "From (SELECT @i:=0) foo,DocVentasDet D " & _
            "Left Join DocVentasDetFormula A " & _
                "On D.IdEmpresa = A.IdEmpresa And D.IdSucursal = A.IdSucursal And D.IdDocumento = A.IdDocumento And D.IdSerie = A.IdSerie " & _
                "And D.IdDocVentas = A.IdDocVentas And D.Item = A.ItemDocVentas " & _
            "Inner Join Productos P " & _
                "On D.IdEmpresa = P.IdEmpresa And If(A.IdProducto Is Null,D.IdProducto,A.IdProducto) = P.IdProducto " & _
            "Left Join (" & _
                "Select C.IdEmpresa,D.IdProducto," & _
                "Sum(If('PEN' = 'PEN',If(C.IdMoneda = 'PEN',(D.VVUnit * D.Cantidad),((D.VVUnit * C.TipoCambio) * D.Cantidad))," & _
                "If(C.IdMoneda = 'USD',(D.VVUnit * D.Cantidad),((D.VVUnit / C.TipoCambio) * D.Cantidad))) * If(C.TipoVale = 'I',1,-1)) / " & _
                "Sum(If(C.TipoVale = 'I',D.Cantidad,D.Cantidad * -1)) As Unit " & _
                "From ValesCab C " & _
                "Inner Join ValesDet D " & _
                    "On C.IdEmpresa = D.IdEmpresa And C.IdSucursal = D.IdSucursal And C.IdValesCab = D.IdValesCab " & _
                "Where C.IdEmpresa = '" & glsEmpresa & "' " & _
                "Group By D.IdProducto" & _
            ") V "
            csql = csql & "On P.IdEmpresa = V.IdEmpresa And P.IdProducto = V.IdProducto " & _
            "Where D.IdEmpresa = '" & glsEmpresa & "' And D.IdSucursal = '" & glsSucursal & "' And D.IdDocumento = '" & strIdDocumento & "' " & _
            "And D.IdSerie = '" & strIdSerie & "' And D.IdDocVentas = '" & strIdDocventas & "' And P.IdTipoProducto <> '06002'"
    Cn.Execute csql
    
    If Len(Trim("" & traerCampo("ValesDet", "IdValesCab", "IdSucursal", glsSucursal, True, "TipoVale = '" & strTipoVale & "' And IdValesCab = '" & strCod & "'"))) = 0 Then
        csql = "Delete From ValesCab " & _
               "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And TipoVale = '" & strTipoVale & "' " & _
               "And IdValesCab = '" & strCod & "'"
        Cn.Execute csql
        
        csql = "Update Docventas Set TipoVale = '',IdValesCab = '',IndVale = 'N' " & _
               "Where IdDocumento = '" & strIdDocumento & "' And IdSerie = '" & strIdSerie & "' And IdDocVentas = '" & strIdDocventas & "' " & _
               "And IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "'"
        Cn.Execute csql
        
    Else
        '--- INSERTAMOS DOCUMENTO DE REFERENCIA
        If strTipoVale = "I" Then
             csql = "INSERT INTO docreferencia (tipoDocOrigen,numDocOrigen,serieDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,item,idEmpresa,idSucursal) " & _
                    "VALUES ('88','" & strCod & "','000','" & strIdDocumento & "','" & strIdDocventas & "','" & strIdSerie & "',1,'" & glsEmpresa & "','" & glsSucursal & "')"
        Else
            csql = "INSERT INTO docreferencia (tipoDocOrigen,numDocOrigen,serieDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,item,idEmpresa,idSucursal) " & _
                   "VALUES ('99','" & strCod & "','000','" & strIdDocumento & "','" & strIdDocventas & "','" & strIdSerie & "',1,'" & glsEmpresa & "','" & glsSucursal & "')"
        End If
        Cn.Execute csql
            
        '--- ACTUALIZAMOS EL ESTADO DEL DOCUMENTO
         csql = "UPDATE docventas SET tipoVale = '" & strTipoVale & "',idValesCab = '" & strCod & "',indVale = 'S' " & _
                "WHERE idDocumento = '" & strIdDocumento & "'" & _
                " AND idSerie     = '" & strIdSerie & "'" & _
                " AND idDocVentas = '" & strIdDocventas & "'" & _
                " AND idEmpresa = '" & glsEmpresa & "'" & _
                " AND idSucursal = '" & glsSucursal & "'"
        Cn.Execute csql
        
        '--- ACTUALIZAMOS STOCK EN LINEA
        actualizaStock strCod, 0, StrMsgError, strTipoVale, False
        If StrMsgError <> "" Then GoTo Err
    End If
    Set rs = Nothing
    
    Exit Sub

Err:
    Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub generaValeDocVentasLote(strIdDocumento As String, strIdSerie As String, StrTipVale As String, strIdDocventas As String, ByVal strCodMotivoTraslado As String, ByVal StrFecEmision As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rs                      As New ADODB.Recordset
Dim strCod                  As String
Dim strTipoVale             As String
Dim strConcepto             As String
Dim strAbreviaturaDoc       As String
Dim indMod                  As Boolean
Dim strNewPeriodoinv        As String

    csql = ""
    indMod = False
    If Len(Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", glsSucursal, True, " year(FecInicio) = " & Format(StrFecEmision, "yyyy") & " "))) = 0 Then
        strNewPeriodoinv = glsCodPeriodoINV
    Else
        strNewPeriodoinv = Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", glsSucursal, True, " year(FecInicio) = " & Format(StrFecEmision, "yyyy") & " "))
    End If
    
    '--- VERIFICAMOS EL ESTADO DEL DOCUMENTO
    csql = "SELECT indVale,idValesCab FROM docventas " & _
           "WHERE idDocumento = '" & strIdDocumento & "'" & _
            " AND idSerie     = '" & strIdSerie & "'" & _
            " AND idDocVentas = '" & strIdDocventas & "'" & _
            " AND idEmpresa = '" & glsEmpresa & "'" & _
            " AND idSucursal = '" & glsSucursal & "'"
    If rs.State = 1 Then rs.Close
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly

    If rs.Fields("indVale") = "S" Then
        indMod = True
        '--- SELECCIONAMOS EL NUMERO DEL VALE DEL DOCUMENTO
        strCod = Trim("" & rs.Fields("idValesCab"))
        
        '--- ACTUALIZAMOS STOCK EN LINEA
        actualizaStock strCod, 1, StrMsgError, StrTipVale, False
        If StrMsgError <> "" Then GoTo Err
        
        '--- ELIMINAMOS EL VALE YA INGRESADO
        csql = "DELETE FROM valesDet WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoVale = '" & StrTipVale & "' "
        Cn.Execute csql
        
        csql = "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = '" & IIf(StrTipVale = "I", "88", "99") & "' AND numDocOrigen = '" & strCod & "' AND serieDocOrigen = '000' "
        Cn.Execute csql
        
        csql = "DELETE FROM valesCab WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoVale = '" & StrTipVale & "' "
        Cn.Execute csql
        
        csql = "DELETE FROM ValesDetLotes WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoVale = '" & StrTipVale & "' "
        Cn.Execute csql
        
    Else
        indMod = False
    End If
    
    Select Case strIdDocumento
        Case "01", "03", "12", "90", "91", "56"
            strTipoVale = "S"
            strConcepto = "22"
        Case "86"
            strTipoVale = "S"
            If strTipoVale = "" Then strTipoVale = "S"
            strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
        Case "07", "89"
            strTipoVale = "I"
            strConcepto = Trim("" & traerCampo("motivostraslados", "idconcepto", "idMotivoTraslado", strCodMotivoTraslado, False, " idDocumento = '" & strIdDocumento & "' "))
    End Select
    
    If indMod = False Then
        strCod = generaCorrelativoAnoMes_Vale("ValesCab", "idValesCab", strTipoVale)
    End If
    strAbreviaturaDoc = "" & traerCampo("documentos", "AbreDocumento", "idDocumento", strIdDocumento, False)

    '--- GENERAMOS CABECERA
    csql = "INSERT INTO ValesCab (idValesCab,tipoVale,fechaEmision, valorTotal, igvTotal,precioTotal, idProvCliente,idConcepto, idAlmacen, idMoneda," & _
           "GlsDocReferencia,TipoCambio,idEmpresa,idSucursal,idPeriodoInv,FechaRegistro,IdUsuarioRegistro) " & _
           "SELECT '" & strCod & "','" & strTipoVale & "',d.FecEmision, d.TotalValorVenta, d.TotalIGVVenta,d.TotalPrecioVenta, d.idPerCliente," & _
           "'" & strConcepto & "', d.idAlmacen, d.idMoneda,'" & strAbreviaturaDoc & " " & strIdSerie & "-" & strIdDocventas & "', d.TipoCambio,idEmpresa," & _
           "idSucursal,'" & strNewPeriodoinv & "',GETDATE(),'" & glsUser & "' " & _
           "FROM docventas d " & _
           "WHERE d.idDocumento = '" & strIdDocumento & "'" & _
            " AND d.idSerie     = '" & strIdSerie & "'" & _
            " AND d.idDocVentas = '" & strIdDocventas & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'"
            
    Cn.Execute csql

    '--- GENERAMOS DETALLE DEL VALE
    'csql = "INSERT INTO valesdet (tipoVale,idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
                                 "TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal,NumLote,FecVencProd,idlote) " & _
           "SELECT '" & strTipoVale & "','" & StrCod & "',(@i:=@i +1) As item,d.idProducto,d.GlsProducto,p.idUMCompra,1 AS Factor,d.Afecto,(l.Cantidad / d.Factor) AS Cantidad,d.VVUnit,d.IGVUnit,d.PVUnit,((l.Cantidad / d.Factor) * d.VVUnit ) as TotalVVNeto," & _
                  "((l.Cantidad / d.Factor) * d.IGVUnit) as TotalIGVNeto, ((l.Cantidad / d.Factor) * d.PVUnit)  as TotalPVNeto, d.idMoneda, d.idEmpresa, d.idSucursal, d.NumLote, d.FecVencProd,l.idlote " & _
           "From docventasdet d,docventasdetlote L, productos p,(SELECT @i:=0) foo " & _
           "WHERE d.idDocumento = '" & strIdDocumento & "'" & _
            " AND d.idSerie     = '" & strIdSerie & "'" & _
            " AND d.idDocVentas = '" & strIdDocventas & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'" & _
            " and d.iddocumento = l.iddocumento " & _
            " and d.idserie = l.idserie " & _
            " and d.iddocventas = l.iddocventas " & _
            " AND d.idproducto = l.idproducto " & _
            " and d.item = l.ItemDocVentas " & _
            " and d.idempresa = l.idempresa " & _
            " and d.idsucursal = l.idsucursal " & _
            " AND d.idProducto = p.idProducto " & _
            " AND p.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idTipoProducto <> '06002'" 'diferente a servicios
            
    csql = "INSERT INTO valesdet (tipoVale,idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
                                 "TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal,FecVencProd,idlote,NumLote) " & _
           "SELECT '" & strTipoVale & "','" & strCod & "',d.item,d.idProducto,d.GlsProducto,p.idUMCompra,1 AS Factor,d.Afecto,(d.Cantidad / d.Factor) AS Cantidad,d.VVUnit,d.IGVUnit,d.PVUnit,((d.Cantidad / d.Factor) * d.VVUnit ) as TotalVVNeto," & _
                  "((d.Cantidad / d.Factor) * d.IGVUnit) as TotalIGVNeto, ((d.Cantidad / d.Factor) * d.PVUnit)  as TotalPVNeto, d.idMoneda, d.idEmpresa, d.idSucursal, d.FecVencProd,D.IdLote,D.NumLote " & _
           "From docventasdet d,productos p " & _
           "WHERE d.idDocumento = '" & strIdDocumento & "'" & _
            " AND d.idSerie     = '" & strIdSerie & "'" & _
            " AND d.idDocVentas = '" & strIdDocventas & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'" & _
            " AND d.idProducto = p.idProducto " & _
            " AND p.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idTipoProducto <> '06002'" 'diferente a servicios
            
    Cn.Execute csql
    
    csql = "Insert Into ValesDetLotes(IdEmpresa,IdSucursal,TipoVale,IdValesCab,Item,IdLote,IdProducto,IdUM,Cantidad,CantidadAnt) " & _
           "Select D.IdEmpresa,D.IdSucursal,'" & strTipoVale & "','" & strCod & "',D.Item,L.IdLote,D.IdProducto,D.IdUM,L.Cantidad,0 " & _
           "From DocVentasDet D " & _
           "Inner Join DocVentasDetLote L " & _
                "On D.IdDocumento = L.IdDocumento And D.IdSerie = L.IdSerie And D.IdDocVentas = L.IdDocVentas And D.IdProducto = L.IdProducto " & _
                "And D.Item = L.ItemDocVentas And D.IdEmpresa = L.IdEmpresa And D.IdSucursal = L.IdSucursal " & _
           "Inner Join Productos P " & _
                "On D.IdProducto = P.IdProducto And P.IdEmpresa = D.IdEmpresa " & _
           "Where D.IdDocumento = '" & strIdDocumento & "'" & _
            " AND d.idSerie     = '" & strIdSerie & "'" & _
            " AND d.idDocVentas = '" & strIdDocventas & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'" & _
            " AND d.idTipoProducto <> '06002'" 'diferente a servicios
            
    Cn.Execute csql
    
    If Len(Trim("" & traerCampo("ValesDet", "IdValesCab", "IdSucursal", glsSucursal, True, "TipoVale = '" & strTipoVale & "' And IdValesCab = '" & strCod & "'"))) = 0 Then
        csql = "Delete From ValesCab " & _
               "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And TipoVale = '" & strTipoVale & "' " & _
               "And IdValesCab = '" & strCod & "'"
        Cn.Execute csql
        
        csql = "Update Docventas Set TipoVale = '',IdValesCab = '',IndVale = 'N' " & _
               "Where IdDocumento = '" & strIdDocumento & "' And IdSerie = '" & strIdSerie & "' And IdDocVentas = '" & strIdDocventas & "' " & _
               "And IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "'"
        Cn.Execute csql
        
    Else
        '--- INSERTAMOS DOCUMENTO DE REFERENCIA
        If strTipoVale = "I" Then
            csql = "INSERT INTO docreferencia (tipoDocOrigen,numDocOrigen,serieDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,item,idEmpresa,idSucursal) " & _
                   "VALUES ('88','" & strCod & "','000','" & strIdDocumento & "','" & strIdDocventas & "','" & strIdSerie & "',1,'" & glsEmpresa & "','" & glsSucursal & "')"
        Else
            csql = "INSERT INTO docreferencia (tipoDocOrigen,numDocOrigen,serieDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,item,idEmpresa,idSucursal) " & _
                   "VALUES ('99','" & strCod & "','000','" & strIdDocumento & "','" & strIdDocventas & "','" & strIdSerie & "',1,'" & glsEmpresa & "','" & glsSucursal & "')"
        End If
        Cn.Execute csql
            
        '--- ACTUALIZAMOS EL ESTADO DEL DOCUMENTO
         csql = "UPDATE docventas SET tipoVale = '" & strTipoVale & "',idValesCab = '" & strCod & "',indVale = 'S' " & _
                "WHERE idDocumento = '" & strIdDocumento & "'" & _
                " AND idSerie     = '" & strIdSerie & "'" & _
                " AND idDocVentas = '" & strIdDocventas & "'" & _
                " AND idEmpresa = '" & glsEmpresa & "'" & _
                " AND idSucursal = '" & glsSucursal & "'"
        Cn.Execute csql
        
        '--- ACTUALIZAMOS STOCK EN LINEA
        actualizaStock strCod, 0, StrMsgError, strTipoVale, False
        If StrMsgError <> "" Then GoTo Err
                
        'actualizaStock_Lote StrCod, 0, StrMsgError, strTipoVale, False
        'If StrMsgError <> "" Then GoTo Err
    End If
    Set rs = Nothing
    
    Exit Sub

Err:
    Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Public Sub generaValeRecMercaderiaAuto(strIdDocumento As String, strIdSerie As String, strIdDocventas As String, ByVal strCodMotivoTraslado As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rs As New ADODB.Recordset
Dim strCod As String
Dim strTipoVale As String
Dim strConcepto As String
Dim strSucursalRec As String
Dim strCodPeriodoINV As String
Dim strAbreviaturaDoc As String

    csql = ""
    '--- VERIFICAMOS EL ESTADO DEL DOCUMENTO
    csql = "SELECT idPerCliente,idValeCabRec FROM docventas " & _
           "WHERE idDocumento = '" & strIdDocumento & "'" & _
            " AND idSerie     = '" & strIdSerie & "'" & _
            " AND idDocVentas = '" & strIdDocventas & "'" & _
            " AND idEmpresa = '" & glsEmpresa & "'" & _
            " AND idSucursal = '" & glsSucursal & "'"
    
    If rs.State = 1 Then rs.Close
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    strSucursalRec = "" & rs.Fields("idPerCliente")

    If rs.Fields("idValeCabRec") <> "" Then
        '--- SELECCIONAMOS EL NUMERO DEL VALE DEL DOCUMENTO
        strCod = Trim("" & rs.Fields("idValeCabRec"))
        
        '--- ACTUALIZAMOS STOCK EN LINEA
        actualizaStock strCod, 1, StrMsgError, False, strSucursalRec
        If StrMsgError <> "" Then GoTo Err
        
        '--- ELIMINAMOS EL VALE YA INGRESADO
        csql = "DELETE FROM valesDet WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & strSucursalRec & "'"
        Cn.Execute csql
        
        csql = "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & strSucursalRec & "' AND tipoDocOrigen = '99' AND numDocOrigen = '" & strCod & "' AND serieDocOrigen = '000'"
        Cn.Execute csql
        
        csql = "DELETE FROM valesCab WHERE idValesCab = '" & strCod & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & strSucursalRec & "'"
        Cn.Execute csql
        
    Else
        '--- GENERAMOS CODIGO DEL VALE
        strCod = GeneraCorrelativoAnoMes("ValesCab", "idValesCab")
    End If
    
    '--- Seleccionamos el periodo
    rs.Close
    csql = "SELECT pi.idPeriodoInv " & _
           "FROM periodosinv pi " & _
           "WHERE pi.estPeriodoInv = 'ACT' AND pi.idEmpresa = '" & glsEmpresa & "' AND pi.idSucursal = '" & strSucursalRec & "'"
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        strCodPeriodoINV = "" & rs.Fields("idPeriodoInv")
    Else
        StrMsgError = "Periodo no abierto en la sucursal de destino"
        GoTo Err
    End If
    
    '---SELECCIONAMOS EN BASE AL TIPO DE DOCUMENTO EL TIPO DEL VALE !!!! ESTO SE DEBERIA JALAR DE LA TABLA DOCUMENTOS
    strTipoVale = "I"
    If strCodMotivoTraslado = "06090006" Then 'traslado
        strConcepto = "07"
    Else
        strConcepto = "09"
    End If
    strAbreviaturaDoc = "" & traerCampo("documentos", "AbreDocumento", "idDocumento", strIdDocumento, False)
    
    '--- GENERAMOS CABECERA
    csql = "INSERT INTO ValesCab (idValesCab,tipoVale,fechaEmision, valorTotal, igvTotal,precioTotal, idProvCliente,idConcepto, idAlmacen, idMoneda," & _
           "GlsDocReferencia,TipoCambio,idEmpresa,idSucursal,idPeriodoInv,FechaRegistro,IdUsuarioRegistro) " & _
           "SELECT '" & strCod & "','" & strTipoVale & "',d.FecEmision, d.TotalValorVenta, d.TotalIGVVenta,d.TotalPrecioVenta, d.idPerCliente," & _
           "'" & strConcepto & "', d.idAlmacen, d.idMoneda,'" & strAbreviaturaDoc & " " & strIdSerie & "-" & strIdDocventas & "', d.TipoCambio,idEmpresa," & _
           "'" & strSucursalRec & "','" & strCodPeriodoINV & "',SysDate(),'" & glsUser & "' " & _
           "FROM docventas d " & _
           "WHERE d.idDocumento = '" & strIdDocumento & "'" & _
            " AND d.idSerie     = '" & strIdSerie & "'" & _
            " AND d.idDocVentas = '" & strIdDocventas & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'"
            
    Cn.Execute csql
    
    '--- GENERAMOS DETALLE DEL VALE
    csql = "INSERT INTO valesdet (idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
                                 "TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal,NumLote,FecVencProd) " & _
           "SELECT '" & strCod & "',d.item,d.idProducto,d.GlsProducto,p.idUMCompra,1 AS Factor,d.Afecto,(d.Cantidad / d.Factor) AS Cantidad,d.VVUnit,d.IGVUnit,d.PVUnit,d.TotalVVNeto," & _
                  "d.TotalIGVNeto, d.TotalPVNeto, d.idMoneda, d.idEmpresa, '" & strSucursalRec & "', d.NumLote, d.FecVencProd " & _
           "From docventasdet d, productos p " & _
           "WHERE d.idDocumento = '" & strIdDocumento & "'" & _
            " AND d.idSerie     = '" & strIdSerie & "'" & _
            " AND d.idDocVentas = '" & strIdDocventas & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'" & _
            " AND d.idProducto = p.idProducto " & _
            " AND p.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idTipoProducto <> '06002'" 'diferente a servicios
    Cn.Execute csql
    
    '--- INSERTAMOS DOCUMENTO DE REFERENCIA
    csql = "INSERT INTO docreferencia (tipoDocOrigen,numDocOrigen,serieDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,item,idEmpresa,idSucursal) " & _
                              "VALUES ('99','" & strCod & "','000','" & strIdDocumento & "','" & strIdDocventas & "','" & strIdSerie & "',1,'" & glsEmpresa & "','" & glsSucursal & "')"
    Cn.Execute csql
        
    '--- ACTUALIZAMOS EL ESTADO DEL DOCUMENTO
    csql = "UPDATE docventas SET idValeCabRec = '" & strCod & "' " & _
           "WHERE idDocumento = '" & strIdDocumento & "'" & _
            " AND idSerie     = '" & strIdSerie & "'" & _
            " AND idDocVentas = '" & strIdDocventas & "'" & _
            " AND idEmpresa = '" & glsEmpresa & "'" & _
            " AND idSucursal = '" & glsSucursal & "'"
    Cn.Execute csql
    
    '--- ACTUALIZAMOS STOCK EN LINEA
    actualizaStock strCod, 0, StrMsgError, False, strSucursalRec
    If StrMsgError <> "" Then GoTo Err
    Set rs = Nothing
    
    Exit Sub
    
Err:
    Set rs = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Function traerCostoUnit(ByVal codproducto As String, ByVal codalmacen As String, ByVal CodMoneda As String, ByRef StrMsgError As String) As Double
Dim CosUni  As ADODB.Recordset
On Error GoTo Err

    csql = "SELECT ( " & _
           "SUM( " & _
           "Case conceptos.indCosto " & _
           "WHEN 'S' THEN " & _
           "IIF((valescab.tipoVale = 'I'),(valesdet.Cantidad),((valesdet.Cantidad) * -(1))) " & _
           "* " & _
           "CASE '" & CodMoneda & "' WHEN 'PEN' THEN IIF(valescab.idMoneda = 'PEN', valesdet.VVUnit,valesdet.VVUnit * t.tcventa) " & _
           "WHEN 'USD' THEN IIF(valescab.idMoneda = 'USD', valesdet.VVUnit,valesdet.VVUnit / t.tcventa) " & _
           "End " & _
           "Else " & _
           "0 " & _
           "End " & _
           ") " & _
           "/ " & _
           "SUM( " & _
           "Case conceptos.indCosto " & _
           "WHEN 'S' THEN " & _
           "IIF((valescab.tipoVale = 'I'),(valesdet.Cantidad),((valesdet.Cantidad) * -(1))) " & _
           "Else " & _
           "0 " & _
           "End " & _
           ") " & _
           ") AS COSTO_UNITARIO "
    csql = csql & "FROM valescab " & _
            "INNER JOIN valesdet  " & _
            "ON valescab.idValesCab = valesdet.idValesCab  " & _
            "AND valescab.idEmpresa = valesdet.idEmpresa  " & _
            "AND valescab.idSucursal = valesdet.idSucursal  " & _
            "AND valescab.tipoVale = valesdet.tipoVale  " & _
            "INNER JOIN conceptos  " & _
            "ON valescab.idConcepto = conceptos.idConcepto  " & _
            "INNER JOIN  PeriodosINV pi  On valescab.idEmpresa = pi.idEmpresa  And valescab.idPeriodoINV = pi.idPeriodoINV  And valescab.idSucursal = pi.idSucursal " & _
            "LEFT JOIN  tiposdecambio t  " & _
            "On " & _
            "" & _
            "valescab.fechaEmision = t.fecha "
    csql = csql & "WHERE "
    csql = csql & "valescab.idEmpresa = '" & glsEmpresa & "' "
    csql = csql & "AND valescab.idPeriodoInv = '" & glsCodPeriodoINV & "' And valesdet.idProducto = '" & codproducto & "' "
    csql = csql & "AND valescab.idAlmacen = '" & codalmacen & "' "
    csql = csql & "AND valescab.estValeCab <> 'ANU' AND pi.estPeriodoInv = 'ACT' "

    Set CosUni = New ADODB.Recordset
    CosUni.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not CosUni.EOF Then
       traerCostoUnit = IIf(IsNull(CosUni.Fields("COSTO_UNITARIO")), 0, CosUni.Fields("COSTO_UNITARIO"))
    End If
    If CosUni.State = 1 Then CosUni.Close: Set CosUni = Nothing
    
Exit Function
Err:
    If CosUni.State = 1 Then CosUni.Close: Set CosUni = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Function

Public Sub EjecutaSQLFormDocVentasEmpresas(F As Form, ByVal tipoOperacion As Integer, ByRef StrMsgError As String, ByVal strTD As String, ByVal strSerie As String, ByVal g As dxDBGrid, ByVal D As dxDBGrid, ByVal indGeneraVale As Boolean, ByRef rsTempDocVlote As ADODB.Recordset, ByRef RsDetProductos As ADODB.Recordset, ByRef rsTempLiquidacion As ADODB.Recordset, ByRef rsdEmpresas As ADODB.Recordset, NArrTotales() As Double)
On Error GoTo Err
Dim C As Object
Dim csql As String
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strValCod As String
Dim strCodMotivoTranslado As String
Dim strCodCliente As String
Dim strCodVendedorCampo As String
Dim rst As New ADODB.Recordset
Dim strCod As String
Dim strCodFacRef As String
Dim obsdocventas As String
Dim RsTempVale   As New ADODB.Recordset
Dim rstsepar     As New ADODB.Recordset
Dim i As Integer
Dim NItem        As Integer
Dim cadUpd As String
Dim rstUpd As New ADODB.Recordset

    NItem = 0
    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                
                If UCase(strCampo) = UCase("idMotivoTraslado") Then
                    strCodMotivoTranslado = C.Value
                End If
                
                If UCase(strCampo) = UCase("idPerVendedorCampo") Then
                    strCodVendedorCampo = C.Value
                End If
                
                If UCase(strCampo) = UCase("idPerCliente") Then
                    strCodCliente = C.Value
                End If
    
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        If UCase(strCampo) = UCase("idDocVentas") Then
                            If Trim(C.Value) = "" Then
                                strCod = generaCorrelativoDocVentas(strTD, strSerie)
                                C.Text = strCod
                            Else
                                If traerCampo("docventas", "idDocVentas", "idDocumento", strTD, True, "idSucursal = '" & glsSucursal & "' AND idDocVentas = '" & Trim(C.Value) & "' AND idSerie = '" & strSerie & "'") <> "" Then
                                    StrMsgError = "El Documento numero " & C.Value & " con serie " & strSerie & " ya existe"
                                    GoTo Err
                                End If
                            End If
                            strValCod = Trim(C.Value)
                        End If
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & Val(C.Value) & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    Case 1
                        If UCase(strCampo) <> UCase("idDocVentas") Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = Val(C.Value)
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
            csql = "INSERT INTO docventas(" & strCampos & ",idDocumento,idEmpresa,idSucursal,FecRegistro,idUsuarioReg,TotalValorVenta,TotalIGVVenta,TotalIvap," & _
                   "TotalDsctoVV,TotalDsctoPV,TotalBaseImponible,TotalExonerado)VALUES(" & _
                   "" & strValores & ",'" & strTD & "','" & glsEmpresa & "','" & glsSucursal & "',sysdate(),'" & glsUser & "'," & NArrTotales(0) & "," & _
                   "" & NArrTotales(1) & "," & NArrTotales(2) & "," & NArrTotales(3) & "," & NArrTotales(4) & "," & NArrTotales(5) & "," & NArrTotales(6) & ")"
                   
        Case 1
            csql = "UPDATE docventas SET " & strCampos & ",TotalValorVenta = " & NArrTotales(0) & ",TotalIGVVenta = " & NArrTotales(1) & "," & _
                   "TotalIvap = " & NArrTotales(2) & ",TotalDsctoVV = " & NArrTotales(3) & ",TotalDsctoPV = " & NArrTotales(4) & "," & _
                   "TotalBaseImponible = " & NArrTotales(5) & ",TotalExonerado = " & NArrTotales(6) & " " & _
                   "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strValCod & "' AND idSerie = '" & strSerie & "'"
            
    End Select
    
    '--- Graba controles
    Cn.Execute csql
    
    '--- Grabando Grilla detalle
    If TypeName(g) <> "Nothing" Then
        Cn.Execute "DELETE FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strValCod & "' AND idSerie = '" & strSerie & "'"
        
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
                            strValores = strValores & Val(g.Columns(i).Value) & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO docventasdet(" & strCampos & ",idDocumento,idDocVentas,idSerie,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strTD & "','" & strValCod & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "')"
            Cn.Execute csql
            
            g.Dataset.Next
        Loop
    End If
    
    '--- Graba Nueva Tabla docreferena entre empresas
    If TypeName(D) <> "Nothing" Then
        Cn.Execute "DELETE FROM docreferenciaEmpresas WHERE idEmpresaOrigen = '" & glsEmpresa & "' AND idSucursalOrigen = '" & glsSucursal & "' AND tipoDocOrigen = '" & strTD & "' AND numDocOrigen = '" & strValCod & "' AND serieDocOrigen = '" & strSerie & "'"
        
        If sw_graba_factura = True And strTD = "86" Then
            Cn.Execute "DELETE FROM docreferenciaEmpresas WHERE idEmpresaOrigen = '" & glsEmpresa & "' AND idSucursalOrigen = '" & glsSucursal & "' AND tipoDocOrigen = '01' AND tipoDocReferencia = '" & strTD & "' AND numDocReferencia = '" & strValCod & "' AND serieDocReferencia = '" & strSerie & "'"
        End If
        
        D.Dataset.First
        Do While Not D.Dataset.EOF
            If Trim(D.Columns.ColumnByFieldName("idDocumento").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idSerie").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idNumDoc").Value) <> "" Then
                strCampos = ""
                strValores = ""
                strCodFacRef = ""
                For i = 0 To D.Columns.Count - 1
                    If UCase(left(D.Columns(i).ObjectName, 1)) = "W" Then
                        strTipoDato = Mid(D.Columns(i).ObjectName, 2, 1)
                        strCampo = Mid(D.Columns(i).ObjectName, 3)
                        
                        strCampos = strCampos & strCampo & ","
                        If strCampo = "tipoDocReferencia" Then strCodFacRef = D.Columns(i).Value
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & D.Columns(i).Value & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(D.Columns(i).Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(D.Columns(i).Value, "yyyy-mm-dd") & "',"
                        End Select
                    End If
                Next
                
                If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
                If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
                
                If sw_graba_factura = True And strTD = "86" And strCodFacRef = "01" Then
                    csql = "INSERT INTO docreferenciaEmpresas(Item,tipoDocOrigen,serieDocOrigen,numDocOrigen,tipoDocReferencia,numDocReferencia,serieDocReferencia,idEmpresa,idSucursal,idEmpresaReferencia,idSucursalReferencia) VALUES(" & strValores & ",'" & strTD & "','" & strValCod & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "','" & D.Columns.ColumnByFieldName("empresaReferencia").Value & "','" & D.Columns.ColumnByFieldName("sucursalReferencia").Value & "')"
                    Cn.Execute csql
                End If
                
                csql = "INSERT INTO docreferenciaEmpresas(" & strCampos & ",idEmpresaReferencia,idSucursalReferencia,tipoDocOrigen,numDocOrigen,serieDocOrigen,idEmpresaOrigen,idSucursalOrigen) VALUES(" & strValores & ",'" & D.Columns.ColumnByFieldName("empresaReferencia").Value & "','" & D.Columns.ColumnByFieldName("sucursalReferencia").Value & "', '" & strTD & "','" & strValCod & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "')"
                Cn.Execute csql
            End If
            D.Dataset.Next
        Loop
    End If
        
    If rsdEmpresas.State = 1 Then
        With rsdEmpresas
            If Not .EOF Then
                .MoveFirst
               Do While Not .EOF
                   NItem = NItem + 1
                   
                   csql = "Insert Into DocrefeRenciaEmpresasDet(item, idEmpresaOrigen, idSucursalOrigen, tipoDocOrigen, numDocOrigen, serieDocOrigen, idEmpresaReferencia, idSucursalReferencia, tipoDocReferencia, numDocReferencia, serieDocReferencia, idProducto, Cantidad) " & _
                          "Values('" & NItem & "','" & glsEmpresa & "','" & glsSucursal & "','" & strTD & "','" & strValCod & "','" & strSerie & "','" & .Fields("idEmpresa").Value & "','" & .Fields("idSucursal").Value & "','" & .Fields("idDocumento").Value & "','" & .Fields("idDocventas").Value & "','" & .Fields("idSerie").Value & "','" & .Fields("idProducto").Value & "','" & .Fields("Cantidad").Value & "')"
                   Cn.Execute (csql)
                   .MoveNext
                Loop
            End If
        End With
    End If
    '------------------------------------------------------------------------------------------
    
    obsdocventas = "" & glsobservacionventas
    csql = "update docventas set obsdocventas = '" & obsdocventas & "',obsdocventas2 = '" & ("" & glsobservacionventas2) & "' " & _
            "WHERE idEmpresa = '" & glsEmpresa & "'" & _
            "AND idSucursal = '" & glsSucursal & "'" & _
            "AND idDocumento = '" & strTD & "'" & _
            "AND idDocVentas = '" & strValCod & "'" & _
            "AND idSerie = '" & strSerie & "' "
    Cn.Execute csql
     
    '--- Acumulamos el peso del detalle del documento
    csql = "SELECT SUM(idTallaPeso) as totalTallaPeso From DocventasDet WHERE idEmpresa = '" & glsEmpresa & "' " & _
            "AND idSucursal = '" & glsSucursal & "' " & _
            "AND idDocumento = '" & strTD & "' " & _
            "AND idDocVentas = '" & strValCod & "' " & _
            "AND idSerie = '" & strSerie & "' "
               
    If rstUpd.State = 1 Then rstUpd.Close
    rstUpd.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstUpd.EOF Then
        cadUpd = "Update docventas Set totalTallaPeso = '" & rstUpd.Fields("totalTallaPeso").Value & "' WHERE idEmpresa = '" & glsEmpresa & "' " & _
                "AND idSucursal = '" & glsSucursal & "' " & _
                "AND idDocumento = '" & strTD & "' " & _
                "AND idDocVentas = '" & strValCod & "' " & _
                "AND idSerie = '" & strSerie & "' "
    End If
    Cn.Execute (cadUpd)
    
    If rstUpd.State = 1 Then rstUpd.Close: Set rstUpd = Nothing
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If rstsepar.State = 1 Then rstsepar.Close: Set rstsepar = Nothing
    
    Exit Sub

Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
