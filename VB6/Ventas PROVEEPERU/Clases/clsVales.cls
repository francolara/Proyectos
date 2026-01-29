VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub EjecutaSQLFormValesConver(F As Form, tipoOperacion As Integer, ByRef StrMsgError As String, gOri As dxDBGrid, gDes As dxDBGrid, ByRef strVarCodValeIng As String, ByRef strVarCodValeSal As String)
On Error GoTo Err
Dim C As Object
Dim rst As New ADODB.Recordset
Dim csql As String
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strValCod As String
Dim strCod As String
Dim strCodValeSalida As String
Dim strCodValeIngreso As String
Dim indTrans As Boolean

    indTrans = False
    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then 'And C.Visible = True Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        If UCase(strCampo) = UCase("idValesConver") Then
                            strValCod = generaCorrelativoAnoMes("valesconver", "idValesConver")
                            C.Text = strValCod
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
                        If UCase(strCampo) <> UCase("idValesConver") Then
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
            csql = "INSERT INTO valesconver(" & strCampos & ",idEmpresa,idSucursal) VALUES(" & strValores & ",'" & glsEmpresa & "','" & glsSucursal & "')"
        Case 1
            csql = "UPDATE valesconver SET " & strCampos & " WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesConver = '" & strValCod & "'"
    End Select
    
    indTrans = True
    Cn.BeginTrans
    
    'Graba controles
    Cn.Execute csql
    
    'Grabando Grilla detalle - Origen
    If TypeName(gOri) <> "Nothing" Then
        Cn.Execute "DELETE FROM valesconverdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesConver = '" & strValCod & "' AND idTipoDetalle = 'O'"
        
        gOri.Dataset.First
        Do While Not gOri.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To gOri.Columns.Count - 1
                If UCase(left(gOri.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(gOri.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(gOri.Columns(i).ObjectName, 3)
                    
                    strCampos = strCampos & strCampo & ","
                    
                    Select Case strTipoDato
                        Case "N"
                            strValores = strValores & Val(gOri.Columns(i).Value) & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(gOri.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(gOri.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO valesconverdet(" & strCampos & ",idValesConver,idTipoDetalle,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strValCod & "','O','" & glsEmpresa & "','" & glsSucursal & "')"
            Cn.Execute csql
            
            gOri.Dataset.Next
        Loop
    End If
    
    
    'Grabando Grilla detalle - Destino
    If TypeName(gDes) <> "Nothing" Then
        Cn.Execute "DELETE FROM valesconverdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesConver = '" & strValCod & "' AND idTipoDetalle = 'D'"
        
        gDes.Dataset.First
        Do While Not gDes.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To gDes.Columns.Count - 1
                If UCase(left(gDes.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(gDes.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(gDes.Columns(i).ObjectName, 3)
                    
                    strCampos = strCampos & strCampo & ","
                    
                    Select Case strTipoDato
                        Case "N"
                            strValores = strValores & Val(gDes.Columns(i).Value) & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(gDes.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(gDes.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO valesconverdet(" & strCampos & ",idValesConver,idTipoDetalle,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strValCod & "','D','" & glsEmpresa & "','" & glsSucursal & "')"
            Cn.Execute csql
            
            gDes.Dataset.Next
        Loop
    End If
    
    
    If tipoOperacion = 1 Then 'si es modificacion
        csql = "SELECT idValeIngreso,idValeSalida FROM valesconver WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesConver = '" & strValCod & "'"
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        
        If Not rst.EOF Then
            strCodValeIngreso = "" & rst.Fields("idValeIngreso")
            strCodValeSalida = "" & rst.Fields("idValeSalida")
            
            'Eliminamos el vale de ingreso
            Cn.Execute "DELETE FROM valesdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & strCodValeIngreso & "'"
            Cn.Execute "DELETE FROM valescab WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & strCodValeIngreso & "'"
            
            'Eliminamos el vale de salida
            Cn.Execute "DELETE FROM valesdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & strCodValeSalida & "'"
            Cn.Execute "DELETE FROM valescab WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & strCodValeSalida & "'"
        End If
    Else
        strCodValeIngreso = ""
        strCodValeSalida = ""
    End If
    
    'GENERAMOS VALE SALIDA
    'GENERAMOS CODIGO DEL VALE
    If strCodValeSalida = "" Then
        strCodValeSalida = generaCorrelativoAnoMes("ValesCab", "idValesCab")
    End If
    
    'GENERAMOS CABECERA
    csql = "INSERT INTO ValesCab (idValesCab,tipoVale,fechaEmision, valorTotal, igvTotal," & _
                       "precioTotal, idProvCliente,idConcepto, idAlmacen, idMoneda,GlsDocReferencia,TipoCambio,idEmpresa,idSucursal,idPeriodoInv) " & _
           "SELECT '" & strCodValeSalida & "','S',d.FecRegistro,0, 0," & _
                  "0, '" & glsSystem & "','16', d.idAlmacenOrigen, d.idMoneda,'', 0 ,idEmpresa,idSucursal,'" & glsCodPeriodoINV & "' " & _
           "FROM valesconver d " & _
           "WHERE d.idValesConver = '" & strValCod & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'"
    Cn.Execute csql
    
    'GENERAMOS DETALLE DEL VALE
    csql = "INSERT INTO valesdet (idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
                                 "TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal) " & _
           "SELECT '" & strCodValeSalida & "',d.item,d.idProducto,p.GlsProducto,d.idUM,d.Factor,0,(d.Cantidad / d.Factor) AS Cantidad,0,0,0,0," & _
                  "0 , 0, '',d.idEmpresa,d.idSucursal " & _
           "From valesconverdet d,productos p " & _
           "WHERE d.idProducto = p.idProducto" & _
            " AND p.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idValesConver = '" & strValCod & "'" & _
            " AND d.idTipoDetalle = 'O'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'"
    Cn.Execute csql
    
    
    'GENERAMOS VALE DE INGRESO
    'GENERAMOS CODIGO DEL VALE
    If strCodValeIngreso = "" Then
        strCodValeIngreso = generaCorrelativoAnoMes("ValesCab", "idValesCab")
    End If
    
    'GENERAMOS CABECERA
    csql = "INSERT INTO ValesCab (idValesCab,tipoVale,fechaEmision, valorTotal, igvTotal," & _
                       "precioTotal, idProvCliente,idConcepto, idAlmacen, idMoneda,GlsDocReferencia,TipoCambio,idEmpresa,idSucursal,idPeriodoInv) " & _
           "SELECT '" & strCodValeIngreso & "','I',d.FecRegistro,0, 0," & _
                  "0, '" & glsSystem & "','13', d.idAlmacenDestino, d.idMoneda,'', 0 ,idEmpresa,idSucursal,'" & glsCodPeriodoINV & "' " & _
           "FROM valesconver d " & _
           "WHERE d.idValesConver = '" & strValCod & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'"
    Cn.Execute csql
    
    'GENERAMOS DETALLE DEL VALE
    csql = "INSERT INTO valesdet (idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVNeto," & _
                                 "TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal) " & _
           "SELECT '" & strCodValeIngreso & "',d.item,d.idProducto,p.GlsProducto,d.idUM,d.Factor,0,(d.Cantidad / d.Factor) AS Cantidad,0,0,0,0," & _
                  "0 , 0, '',d.idEmpresa,d.idSucursal " & _
           "From valesconverdet d,productos p " & _
           "WHERE d.idProducto = p.idProducto" & _
            " AND p.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idValesConver = '" & strValCod & "'" & _
            " AND d.idTipoDetalle = 'D'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'"
    Cn.Execute csql
    
    'Actualizamos los numeros de vales al registro
    csql = "UPDATE valesconver SET idValeIngreso = '" & strCodValeIngreso & "',idValeSalida = '" & strCodValeSalida & "' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesConver = '" & strValCod & "'"
    Cn.Execute csql
    
    strVarCodValeIng = strCodValeIngreso
    strVarCodValeSal = strCodValeSalida
    
    Cn.CommitTrans
    
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
End Sub
Public Sub EjecutaSQLFormPeriodoINV(F As Form, tipoOperacion As Integer, ByRef StrMsgError As String, Optional indGeneraSaldoInicial As Integer = 0, Optional strMoneda As String = "", Optional strFechaCorte As String = "")
On Error GoTo Err
Dim C As Object
Dim csql As String
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strValCod As String
Dim strCod As String
Dim indTrans As Boolean
Dim rsa As New ADODB.Recordset
Dim strCodVale As String

    indTrans = False
    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then 'And C.Visible = True Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        If UCase(strCampo) = UCase("idPeriodoInv") Then
                            strValCod = C.Value
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
                        If UCase(strCampo) <> UCase("idPeriodoInv") Then
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
            csql = "INSERT INTO periodosinv(" & strCampos & ",idEmpresa,idSucursal,estPeriodoInv,idUsuarioReg,FecRegistro,FecInicio) VALUES(" & strValores & ",'" & glsEmpresa & "','" & glsSucursal & "','ACT','" & glsUser & "',sysdate(),'" & Format(DateAdd("D", 1, strFechaCorte), "yyyy-mm-dd") & "')"
        Case 1
            csql = "UPDATE periodosinv SET " & strCampos & " WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idPeriodoInv = '" & strValCod & "'"
    End Select
    
    indTrans = True
    Cn.BeginTrans
    
    If indGeneraSaldoInicial = 1 Then
        rsa.Open "SELECT idAlmacen FROM almacenes WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'", Cn, adOpenForwardOnly, adLockReadOnly
        Do While Not rsa.EOF
            strCodVale = generaCorrelativoAnoMes_Vale("ValesCab", "idValesCab", "I", True)
            Cn.Execute "CALL spu_Ins_ValeNuevoPeriodo('" & glsEmpresa & "','" & glsSucursal & "','" & rsa.Fields("idAlmacen") & "','" & strMoneda & "','" & strFechaCorte & "','" & strValCod & "','" & strCodVale & "')"
            
            rsa.MoveNext
        Loop
    End If
    '--- Graba controles
    Cn.Execute csql
    
    'ACTUALIZANDO EL ESTADO DE LOS PERIODOS PASADOS
    csql = "UPDATE periodosinv SET estPeriodoInv = 'INA',FecFin = '" & strFechaCorte & "' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idPeriodoInv <> '" & strValCod & "' And estPeriodoInv = 'ACT'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
End Sub
    
