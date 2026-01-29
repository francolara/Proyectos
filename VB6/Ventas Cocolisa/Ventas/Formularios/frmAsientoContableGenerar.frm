VERSION 5.00
Begin VB.Form frmAsientoContableGenerar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar Asientos Contables"
   ClientHeight    =   2160
   ClientLeft      =   4545
   ClientTop       =   1605
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbOperar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1620
      Width           =   1140
   End
   Begin VB.CommandButton cmbCancelar 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1620
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   5100
      Begin VB.ComboBox CmbOpciones 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.ComboBox cbxMes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmAsientoContableGenerar.frx":0000
         Left            =   1710
         List            =   "frmAsientoContableGenerar.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   675
         Width           =   2340
      End
      Begin VB.ComboBox cbxAno 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmAsientoContableGenerar.frx":0090
         Left            =   1710
         List            =   "frmAsientoContableGenerar.frx":00AC
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   2340
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1170
         TabIndex        =   8
         Top             =   1125
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1170
         TabIndex        =   4
         Top             =   765
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1170
         TabIndex        =   1
         Top             =   360
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmAsientoContableGenerar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As New ADODB.Recordset
Dim cnn_empresa     As New ADODB.Connection
Dim rsasiento As New ADODB.Recordset
Dim rsDocumentos As New ADODB.Recordset
Dim strAno As String
Dim strMes As String
Dim stridSucursal As String
Dim strglspersona As String
Dim strIdFormaPago As String
Dim strCtaContTotalContado As String
Dim strCtaContTotalCredito As String
Dim strCtaContAfecto As String
Dim strCtaContInaAfecto As String
Dim cdirectorio As String
Dim strSimboloMoneda As String
Dim cconex_empresa  As String
Dim cta12soles As String
Dim cta12dolares As String
Dim cta40IGV As String
Dim cta40IGVPorPagar As String
Dim cta70Detalle As String
Dim i As Integer
Dim ncorrel As Integer
Dim NItem As Integer
Dim dbltotalbaseimponible As Double
Dim dbltotaligvventa As Double
Dim dbltotalexonerado As Double
Dim dbltotalprecioventa As Double
Dim dbltotalprecioventa_Cont As Double
Dim dbltotalprecioventa_CRED As Double
Dim sHaberSoles As Double
Dim sDebeSoles As Double
Dim diferenciaSoles As Double
Dim sHaberDolares As Double
Dim sDebeDolares As Double
Dim diferenciaDolares As Double
Dim cadenadoc          As String
Dim docofi             As String
Dim CIVAPCuenta             As String
Dim cta12soles_relacionada  As String
Dim cta12dolares_relacionada As String
Dim rsDetalle               As New ADODB.Recordset
Dim IndGeneraInt                            As Boolean

Public Sub Genera_Internamente(strMsgError As String, PPeriodo As String, PIndGeneraInt As Boolean)
On Error GoTo ERR
    
    Load Me
    
    For i = 0 To cbxAno.ListCount - 1
        cbxAno.ListIndex = i
        If cbxAno.Text = left(PPeriodo, 4) Then Exit For
    Next
    
    cbxMes.ListIndex = Val(right(PPeriodo, 2)) - 1
    
    IndGeneraInt = True
    
    cmbOperar_Click
    
    PIndGeneraInt = IndGeneraInt
    
    Unload Me
    
    Exit Sub
ERR:
    If strMsgError = "" Then strMsgError = ERR.Description
End Sub

Private Sub cmbCancelar_Click()

    Unload Me
    
End Sub

Private Sub cmbOperar_Click()
On Error GoTo ERR
Dim rsFiltro                As New ADODB.Recordset
Dim strMsgError             As String

    If Trim("" & traerCampo("cierresmes", "estcierre", "Idmes", Format(cbxMes.ListIndex + 1, "00"), True, " idano = '" & cbxAno.Text & "' and IdSistema = '21008' ")) = "C" Then
        strMsgError = "Contabilidad Se Encuentra Cerrado."
        GoTo ERR
    End If
    
    docofi = traerCampo("Parametros", "ValParametro", "GlsParametro", "VISUALIZA_FILTRO_DOCUMENTO", True)
    
    If traerCampo("parametros", "ValParametro", "glsParametro", "TIPO_TRANS_CONTA", True) = "G" Then
        
        AsientosContablesGeneral
    
    Else
        strAno = cbxAno.Text
        strMes = Format(cbxMes.ListIndex + 1, "00")
        CFecha = CAPTURA_FECHA_FIN(strMes, strAno)
        
        Cadena_Oficial = ""
        If right(CmbOpciones.Text, 2) = "01" Then
            Cadena_Oficial = " and d.indoficial = '1' "
        Else
            Cadena_Oficial = " and d.indoficial in('1','0') "
        End If
        
        Cadena_Transferido = ""
        If right(CmbOpciones.Text, 2) = "01" Then
            Cadena_Transferido = " and dv.indTrasladoConta <> 'S' "
        Else
            Cadena_Transferido = " and dv.indTrasladoContaFin <> 'S' "
        End If
        
        If Not IndGeneraInt Then
        
            strSQL = "SELECT p.idproducto,Left(p.glsproducto,150) GlsProducto " & _
                    "FROM docventas dv inner join docventasdet dvd on dv.iddocventas = dvd.iddocventas and dv.idserie = dvd.idserie " & _
                    "and dv.iddocumento = dvd.iddocumento and dv.idempresa = dvd.idempresa " & _
                    "inner join productos p " & _
                    "on dvd.idproducto = p.idproducto and dvd.idempresa = p.idempresa " & _
                    "inner join Documentos d on dv.iddocumento = d.iddocumento " & _
                    "where dv.idEmpresa = '" & glsEmpresa & "' " & _
                    "and length(trim(ifnull(p.ctacontable,''))) = 0 " & _
                    Cadena_Oficial & _
                    Cadena_Transferido & _
                    "and month(dv.fecEmision) = " & Val(strMes) & " and year(dv.fecEmision) = " & Val(strAno) & " " & _
                    "group by p.idproducto order by dv.iddocumento, dv.idSerie, dv.iddocventas "
            
            If rsFiltro.State = 1 Then rsFiltro.Close
            rsFiltro.Open strSQL, Cn, adOpenKeyset, adLockOptimistic
            
            If Not rsFiltro.EOF Then
                rsFiltro.MoveFirst
                Do While Not rsFiltro.EOF
                    strMsgError = strMsgError & "El Producto Con Codigo " & " " & Trim("" & rsFiltro.Fields("idproducto")) & " - " & " " & Trim("" & rsFiltro.Fields("glsproducto")) & " " & " No Tiene Cuenta Contable " & Chr(13) & Chr(10)
                    rsFiltro.MoveNext
                Loop
            End If
            strSQL = ""
        
        End If
        
        If Len(Trim("" & strMsgError)) = 0 Then
            strMsgError = ""
            If traerCampo("parametros", "ValParametro", "glsParametro", "TIPO_TRANS_CONTA_DETALLE", True) = "1" Then
                AsientosContablesDetallado_2
            ElseIf traerCampo("parametros", "ValParametro", "glsParametro", "TIPO_TRANS_CONTA_DETALLE", True) = "2" Then
                AsientosContablesDetalladoRally
            Else
                AsientosContablesDetallado
            End If
        Else
        
            If MsgBox("¿Se encontraron productos que no tienen cuenta contable desea ver la lista y cancelar la generacion del asiento?", vbInformation + vbYesNo, App.Title) = vbYes Then
                MsgBox strMsgError, vbInformation, VB.App.Title
            Else
                strMsgError = ""
                If traerCampo("parametros", "ValParametro", "glsParametro", "TIPO_TRANS_CONTA_DETALLE", True) = "1" Then
                    AsientosContablesDetallado_2
                ElseIf traerCampo("parametros", "ValParametro", "glsParametro", "TIPO_TRANS_CONTA_DETALLE", True) = "2" Then
                    AsientosContablesDetalladoRally
                Else
                    AsientosContablesDetallado
                End If
            End If
        End If
        
    End If
    
    Exit Sub

ERR:
If strMsgError = "" Then strMsgError = ERR.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim fecha As Date
Dim i As Integer
    
    Me.top = 0
    Me.left = 0
    
    IndGeneraInt = False
    
    fecha = Format(getFechaSistema, "dd/mm/yyyy")
    strAno = Format(Year(fecha), "0000")
    strMes = Format(Month(fecha), "00")
    
    cbxAno.Clear
    For i = 2008 To Val(strAno)
        cbxAno.AddItem i
    Next
    
    For i = 0 To cbxAno.ListCount - 1
        cbxAno.ListIndex = i
        If cbxAno.Text = strAno Then Exit For
    Next
    cbxMes.ListIndex = Val(strMes) - 1
    
    CmbOpciones.AddItem "Tributaria" & Space(150) & "01"
    CmbOpciones.AddItem "Financiera" & Space(150) & "02"
    CmbOpciones.ListIndex = 0
    
    If Trim("" & traerCampo("Parametros", "Valparametro", "Glsparametro", "VISUALIZA_FILTRO_DOCUMENTO", True)) = "S" Then
        Label3.Visible = True
        CmbOpciones.Visible = True
    Else
        Label3.Visible = False
        CmbOpciones.Visible = False
    End If

End Sub

Private Sub AsientosContablesGeneral()
On Error GoTo ERR
Dim rsd         As New ADODB.Recordset
Dim strSQL      As String
Dim dblsaldo    As Double
Dim dblValSaldo As Double
Dim CodProd     As String
Dim CodProdAnt  As String
Dim CFecha      As String
Dim cconex_dbbancos As String, cselect As String
Dim cnn_dbbancos    As New ADODB.Connection
Dim rs              As New ADODB.Recordset
Dim ntc             As Double
 
    strAno = cbxAno.Text
    strMes = Format(cbxMes.ListIndex + 1, "00")
    CFecha = CAPTURA_FECHA_FIN(strMes, strAno)
    
    cdirectorio = traerCampo("empresas", "Carpeta", "idEmpresa", glsEmpresa, False)
    cruta = glsRuta_Access & cdirectorio
    
    If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    cconex_dbbancos = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cruta & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
    cnn_dbbancos.Open cconex_dbbancos
    
    ntc = 0#
    cselect = "SELECT CAMBIO FROM CAMBIOS WHERE CVDATE(FECHA) = CVDATE('" & CFecha & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open cselect, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        ntc = Val(rs.Fields("CAMBIO") & "")
    End If
    rs.Close: Set rs = Nothing
    cnn_dbbancos.Close: Set cnn_dbbancos = Nothing
    
    If ntc = 0# Then ntc = 2.85
    cadenadoc = IIf(docofi = "S", "AND (d.iddocumento = '01' or d.iddocumento = '03' or d.iddocumento = '07' or d.iddocumento = '08' or d.iddocumento = '12' or d.iddocumento = '90')", "AND (d.iddocumento = '01' or d.iddocumento = '03' or d.iddocumento = '07' or d.iddocumento = '08' or d.iddocumento = '12'")
    
    strSQL = "SELECT DISTINCT  d.idsucursal,d.IDDOCUMENTO, d.IDSERIE, " & _
            "d.IDDOCVENTAS,d.idtipoticket, " & _
            "IF(d.idmoneda ='PEN' ,d.totalbaseimponible,d.totalbaseimponible * tc.tcVenta)AS totalbaseimponible, " & _
            "IF(d.idmoneda ='PEN' ,d.totaligvventa,d.totaligvventa * tc.tcVenta) AS totaligvventa, " & _
            "IF (d.idmoneda ='PEN' ,d.totalexonerado,d.totalexonerado * tc.tcVenta) AS totalexonerado, " & _
            "IF (d.idmoneda ='PEN', D.totalprecioventa,d.totalprecioventa * tc.tcVenta) AS totalprecioventa , " & _
            "d.estdocventas, " & _
            "IF (pd.idtipoformaPago IS NULL ,(SELECT IdTipoFormaPago from tipoformaspago " & _
            "where tipoformapago ='C'),pd.idtipoformaPago ) as idtipoformaPago, " & _
            "IF (t.tipoformapago is null ,'C',t.tipoformapago) as tipoformapago," & _
            "P.glspersona " & _
            "FROM docventas d " & _
            "left join pagosdocventas pd on " & _
            "d.idsucursal=pd.idsucursal and d.iddocumento=pd.iddocumento and " & _
            "D.IDSERIE = pd.IDSERIE And D.IDDOCVENTAS = pd.IDDOCVENTAS " & _
            "left join  tipoformaspago t " & _
            "on pd.idtipoformapago=t.idtipoformapago " & _
            "left join sucursales s " & _
            "on d.idsucursal=s.idsucursal " & _
            "and s.idempresa= d.idempresa " & _
            "left join  personas p " & _
            "on p.idpersona=s.idsucursal "
            
    strSQL = strSQL & " left join tiposdecambio tc " & _
            "ON day(tc.fecha)=day(d.FecEmision) and  " & _
            "month(tc.fecha)=month(d.FecEmision) and " & _
            "year(tc.fecha)=year(d.FecEmision) " & _
            "Where Month(D.fecEmision) = " & Val(strMes) & " And Year(D.fecEmision) = '" & strAno & "' " & _
             cadenadoc & _
            "AND d.idempresa = '" & glsEmpresa & "'and d.estdocventas <> 'ANU' order by d.idsucursal,tipoformapago "
        
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.Open strSQL, Cn, adOpenKeyset, adLockOptimistic
    
    If cnn_empresa.State = adStateOpen Then cnn_empresa.Close
    cconex_empresa = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glsRuta_Access & "DB_CONTA.MDB" & ";Persist Security Info=False"
    cnn_empresa.Open cconex_empresa
    
    If rsAsientosContables.State = adStateOpen Then rsAsientosContables.Close
    rsAsientosContables.Open "CONTABLE", cnn_empresa, adOpenDynamic, adLockOptimistic
    cnn_empresa.Execute ("DELETE FROM CONTABLE")

    i = 0
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
    
        stridSucursal = rsTemp.Fields("IdSucursal") & ""
        strglspersona = rsTemp.Fields("glspersona") & ""
        strCtaContTotalContado = traerCampo("sucursales", "CtaContTotalContado", "idsucursal", stridSucursal, True)
        strCtaContTotalCredito = traerCampo("sucursales", "CtaContTotalCredito", "idsucursal", stridSucursal, True)
        strCtaContAfecto = traerCampo("sucursales", "CtaContAfecto", "idsucursal", stridSucursal, True)
        strCtaContInaAfecto = traerCampo("sucursales", "CtaContInaAfecto", "idsucursal", stridSucursal, True)
        strIdFormaPago = rsTemp.Fields("tipoformapago") & ""
        
        dbltotalprecioventa_Cont = 0#: dbltotalprecioventa_CRED = 0#: dbltotalbaseimponible = 0#: dbltotalexonerado = 0#
        dbltotaligvventa = 0#
        
        Do While rsTemp.Fields("IdSucursal") & "" = stridSucursal And Not rsTemp.EOF
            
            If Trim(rsTemp.Fields("tipoformapago") & "") = "C" Then
                dbltotalprecioventa_Cont = dbltotalprecioventa_Cont + Val(Format(rsTemp.Fields("totalprecioventa"), "0.00"))
            Else
                dbltotalprecioventa_CRED = dbltotalprecioventa_CRED + Val(Format(rsTemp.Fields("totalprecioventa"), "0.00"))
            End If
         
            dbltotalexonerado = dbltotalexonerado + rsTemp.Fields("totalexonerado") & ""
            dbltotalbaseimponible = dbltotalbaseimponible + rsTemp.Fields("totalbaseimponible") & ""
            dbltotaligvventa = dbltotaligvventa + rsTemp.Fields("totaligvventa") & ""
                        
            rsTemp.MoveNext
            
            If rsTemp.EOF Then Exit Do
            If rsTemp.Fields("IdSucursal") <> stridSucursal Then Exit Do
        Loop
       
        i = i + 1
        If dbltotalprecioventa_Cont > 0 Then
            rsAsientosContables.AddNew
            rsAsientosContables.Fields("idPeriodo") = strAno & strMes
            rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(1, "00000")
            rsAsientosContables.Fields("ValItem") = Format(i, "0000")
            rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
            rsAsientosContables.Fields("GlsDetalle") = "FACTURA AL CONTADO " & strglspersona
            rsAsientosContables.Fields("idCtaContable") = strCtaContTotalContado
            rsAsientosContables.Fields("NumCheque") = 0
            rsAsientosContables.Fields("F3NROREF") = ""
            rsAsientosContables.Fields("TotalImporteS") = Val(Format(dbltotalprecioventa_Cont, "0.00"))
            rsAsientosContables.Fields("idTipoDH") = "D"
            rsAsientosContables.Fields("FecCompro") = CFecha
            
            rsAsientosContables.Fields("ValorTipoCambio") = ntc
            rsAsientosContables.Fields("TotalImporteSD") = Val(Format(Val(Format(dbltotalprecioventa_Cont, "0.00")) / ntc, "0.00"))
            
            rsAsientosContables.Update
        End If
       
        If dbltotalprecioventa_CRED > 0 Then
            i = i + 1
            rsAsientosContables.AddNew
            rsAsientosContables.Fields("idPeriodo") = strAno & strMes
            rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(1, "00000")
            rsAsientosContables.Fields("ValItem") = Format(i, "0000")
            rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
            rsAsientosContables.Fields("GlsDetalle") = "FACTURA AL CREDITO" & strglspersona
            rsAsientosContables.Fields("idCtaContable") = strCtaContTotalCredito
            rsAsientosContables.Fields("NumCheque") = 0
            rsAsientosContables.Fields("F3NROREF") = ""
            rsAsientosContables.Fields("TotalImporteS") = Val(Format(dbltotalprecioventa_CRED, "0.00"))
            rsAsientosContables.Fields("idTipoDH") = "D"
            rsAsientosContables.Fields("FecCompro") = CFecha
            
            rsAsientosContables.Fields("ValorTipoCambio") = ntc
            rsAsientosContables.Fields("TotalImporteSD") = Val(Format(Val(Format(dbltotalprecioventa_CRED, "0.00")) / ntc, "0.00"))
            
            rsAsientosContables.Update
        End If
      
        If dbltotalbaseimponible > 0 Then
            i = i + 1
            rsAsientosContables.AddNew
            rsAsientosContables.Fields("idPeriodo") = strAno & strMes
            rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(1, "00000")
            rsAsientosContables.Fields("ValItem") = Format(i, "0000")
            rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
            rsAsientosContables.Fields("GlsDetalle") = "VENTAS AFECTAS " & strglspersona
            rsAsientosContables.Fields("idCtaContable") = strCtaContAfecto
            rsAsientosContables.Fields("NumCheque") = 0
            rsAsientosContables.Fields("F3NROREF") = ""
            rsAsientosContables.Fields("TotalImporteS") = Val(Format(dbltotalbaseimponible, "0.00"))
            
            rsAsientosContables.Fields("idTipoDH") = "H"
            rsAsientosContables.Fields("FecCompro") = CFecha
            
            rsAsientosContables.Fields("ValorTipoCambio") = ntc
            rsAsientosContables.Fields("TotalImporteSD") = Val(Format(Val(Format(dbltotalbaseimponible, "0.00")) / ntc, "0.00"))
            
            rsAsientosContables.Update
        End If
        
        If dbltotalexonerado > 0 Then
            i = i + 1
            rsAsientosContables.AddNew
            rsAsientosContables.Fields("idPeriodo") = strAno & strMes
            rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(1, "00000")
            rsAsientosContables.Fields("ValItem") = Format(i, "0000")
            rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
            rsAsientosContables.Fields("GlsDetalle") = "VENTAS INAFECTAS " & strglspersona
            rsAsientosContables.Fields("idCtaContable") = strCtaContInaAfecto
            rsAsientosContables.Fields("NumCheque") = 0
            rsAsientosContables.Fields("F3NROREF") = ""
            rsAsientosContables.Fields("TotalImporteS") = Val(Format(dbltotalexonerado, "0.00"))
            
            rsAsientosContables.Fields("idTipoDH") = "H"
            rsAsientosContables.Fields("FecCompro") = CFecha
            
            rsAsientosContables.Fields("ValorTipoCambio") = ntc
            rsAsientosContables.Fields("TotalImporteSD") = Val(Format(Val(Format(dbltotalexonerado, "0.00")) / ntc, "0.00"))
            
            rsAsientosContables.Update
        End If
        
        '---- IGV
        If dbltotaligvventa > 0 Then
            i = i + 1
            rsAsientosContables.AddNew
            rsAsientosContables.Fields("idPeriodo") = strAno & strMes
            rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(1, "00000")
            rsAsientosContables.Fields("ValItem") = Format(i, "0000")
            rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
            rsAsientosContables.Fields("GlsDetalle") = "I.G.V."
            rsAsientosContables.Fields("idCtaContable") = glsCuenta_Igv_Ventas
            rsAsientosContables.Fields("NumCheque") = 0
            rsAsientosContables.Fields("F3NROREF") = ""
            rsAsientosContables.Fields("TotalImporteS") = Val(Format(dbltotaligvventa, "0.00"))
            
            rsAsientosContables.Fields("idTipoDH") = "H"
            rsAsientosContables.Fields("FecCompro") = CFecha
            
            rsAsientosContables.Fields("ValorTipoCambio") = ntc
            rsAsientosContables.Fields("TotalImporteSD") = Val(Format(Val(Format(dbltotaligvventa, "0.00")) / ntc, "0.00"))
            
            rsAsientosContables.Update
        End If
        
    Loop

    MsgBox "Generación Satisfactoria.", vbInformation, "Atención"

    If rsTemp.State = 1 Then rsTemp.Close:  Set rsTemp = Nothing
    If rsAsientosContables.State = 1 Then rsAsientosContables.Close:  Set rsasiento = Nothing
    
    Exit Sub
    
ERR:
    If rsTemp.State = 1 Then rsTemp.Close
    Set rsTemp = Nothing
    MsgBox ERR.Description, vbInformation, App.Title
    
End Sub

Private Sub AsientosContablesDetallado()
On Error GoTo ERR
Dim rscontrol               As New ADODB.Recordset
Dim rsd                     As New ADODB.Recordset
Dim rs                      As New ADODB.Recordset
Dim csql                    As String
Dim strSQL                  As String
Dim CodProd                 As String
Dim CodProdAnt              As String
Dim CFecha                  As String
Dim cconex_dbbancos         As String
Dim cselect                 As String
Dim CCua                    As Double
Dim dblsaldo                As Double
Dim dblValSaldo             As Double
Dim Cadena_Oficial          As String
Dim Cadena_Transferido      As String
Dim Cadena_Oficial2         As String
Dim CSqlC                                       As String

    Me.MousePointer = 11
    strAno = cbxAno.Text
    strMes = Format(cbxMes.ListIndex + 1, "00")
    CFecha = CAPTURA_FECHA_FIN(strMes, strAno)
    
    Cadena_Oficial = ""
    If right(CmbOpciones.Text, 2) = "01" Then
        Cadena_Oficial = " and d.indoficial = '1' "
    Else
        Cadena_Oficial = " and d.indoficial in('1','0') "
    End If
    
    Cadena_Oficial2 = ""
    If right(CmbOpciones.Text, 2) = "01" Then
        Cadena_Oficial2 = " and dc.indoficial = '1' "
    Else
        Cadena_Oficial2 = " and dc.indoficial in('1','0') "
    End If
    
    Cadena_Transferido = ""
    If right(CmbOpciones.Text, 2) = "01" Then
        Cadena_Transferido = " and dv.indTrasladoConta <> 'S' "
    Else
        Cadena_Transferido = " and dv.indTrasladoContaFin <> 'S' "
    End If
            
    strSQL = "SELECT dv.iddocumento,dv.idSerie,dv.iddocventas,dv.idPerCliente,dv.glsCLiente,dv.RUCCliente,dv.idAlmacen,dv.FecEmision,dv.estDocVentas," & _
            "dv.idMoneda,dv.idEmpresa,dv.idSucursal,dv.TotalValorVenta,dv.TotalIGVVenta,DV.TotalIVAP,dv.TotalPrecioVenta,dv.totalbaseimponible," & _
            "if(dv.iddocumento <> '07', t.tcVenta,ifnull(tc.tcNC,T.tcventa)) TipoCambio,dv.idcentrocosto,"
    strSQL = strSQL & _
            "If(Dv.IdMoneda = 'PEN'," & _
                "If(IfNull(E.IdPersona,'') <> ''," & _
                    "If(E.IndAsociada = '1'," & _
                        "DC.IdCtaVtaTerSoles" & _
                    "," & _
                        "If(E.IndSubSidiaria = '1'," & _
                            "DC.IdCtaVtaTerSubSiS" & _
                        "," & _
                            "DC.IdCtaVtaTerMatrizS" & _
                        ")" & _
                    ")" & _
                "," & _
                    "If(IfNull(Z.IdPersona,'') <> ''," & _
                        "If(Z.IndPersonal = '1'," & _
                            "DC.IdCtaVtaPerS" & _
                        "," & _
                            "DC.IdCtaVtaSocS" & _
                        ")" & _
                    "," & _
                        "DC.IdCtaVtaSoles" & _
                    ")" & _
                ")"
    strSQL = strSQL & _
            "," & _
                "If(IfNull(E.IdPersona,'') <> ''," & _
                    "If(E.IndAsociada = '1'," & _
                        "DC.IdCtaVtaTerDolares" & _
                    "," & _
                        "If(E.IndSubSidiaria = '1'," & _
                            "DC.IdCtaVtaTerSubSiD" & _
                        "," & _
                            "DC.IdCtaVtaTerMatrizD" & _
                        ")" & _
                    ")" & _
                "," & _
                    "If(IfNull(Z.IdPersona,'') <> ''," & _
                        "If(Z.IndPersonal = '1'," & _
                            "DC.IdCtaVtaPerD" & _
                        "," & _
                            "DC.IdCtaVtaSocD" & _
                        ")" & _
                    "," & _
                        "DC.IdCtaVtaDolares" & _
                    ")" & _
                ")" & _
            ") IdCtaCliente,Dv.IndTransGratuita, dv.IndTransGratuitaMP "
    strSQL = strSQL & _
            "FROM docventas dv " & _
            "Left Join EmpresasRelacionadas E " & _
                "On Dv.IdEmpresa = E.IdEmpresa And Dv.IdPerCliente = E.IdPersona " & _
            "Left Join Personal Z " & _
                "On Dv.IdEmpresa = Z.IdEmpresa And Dv.IdPerCliente = Z.IdPersona " & _
            "inner join tiposdecambio t " & _
                "on (Day(dv.FecEmision) = Day(t.fecha) and Year(dv.FecEmision) = Year(t.fecha) and Month(dv.FecEmision) = Month(t.fecha)) " & _
            "left join (select x.tcVenta as tcNC, r.tipoDocOrigen, r.serieDocOrigen, r.numDocOrigen " & _
                    "from docventas dt inner join docreferencia r " & _
                    "on dt.IdEmpresa = r.IdEmpresa And dt.iddocumento = r.tipoDocReferencia and dt.idSerie = r.serieDocReferencia and dt.idDocVentas = r.numDocReferencia " & _
                    "inner join tiposdecambio x on (Day(dt.FecEmision) = Day(x.fecha) and Year(dt.FecEmision) = Year(x.fecha) and Month(dt.FecEmision) = Month(x.fecha)) " & _
                    "where dt.IdEmpresa = '" & glsEmpresa & "' And r.tipoDocOrigen = '07' " & _
                    "group by r.tipodocorigen, r.numdocorigen) tc " & _
                "on tc.tipoDocOrigen = dv.idDocumento and tc.serieDocOrigen = dv.idSerie and tc.numDocOrigen = dv.idDocVentas " & _
            "inner join Documentos d on dv.iddocumento = d.iddocumento " & _
            "Inner Join DocumentosCuentas DC " & _
                "On Dv.IdEmpresa = DC.IdEmpresa And Dv.IdDocumento = DC.IdDocumento " & _
            "where dv.idEmpresa = '" & glsEmpresa & "' " & _
            Cadena_Oficial & _
            Cadena_Transferido & _
            "and month(dv.fecEmision) = " & Val(strMes) & " and year(dv.fecEmision) = " & Val(strAno) & " " & _
            "order by dv.iddocumento, dv.idSerie, dv.iddocventas "
    
    'strSQL = strSQL & _
            "FROM docventas dv " & _
            "Left Join EmpresasRelacionadas E " & _
                "On Dv.IdEmpresa = E.IdEmpresa And Dv.IdPerCliente = E.IdPersona " & _
            "Left Join Personal Z " & _
                "On Dv.IdEmpresa = Z.IdEmpresa And Dv.IdPerCliente = Z.IdPersona " & _
            "inner join tiposdecambio t " & _
                "on (Day(dv.FecEmision) = Day(t.fecha) and Year(dv.FecEmision) = Year(t.fecha) and Month(dv.FecEmision) = Month(t.fecha)) " & _
            "left join (select x.tcVenta as tcNC, r.tipoDocOrigen, r.serieDocOrigen, r.numDocOrigen " & _
                    "from docventas dt inner join docreferencia r " & _
                    "on dt.IdEmpresa = r.IdEmpresa And dt.iddocumento = r.tipoDocReferencia and dt.idSerie = r.serieDocReferencia and dt.idDocVentas = r.numDocReferencia " & _
                    "inner join tiposdecambio x on (Day(dt.FecEmision) = Day(x.fecha) and Year(dt.FecEmision) = Year(x.fecha) and Month(dt.FecEmision) = Month(x.fecha)) " & _
                    "where dt.IdEmpresa = '" & glsEmpresa & "' And r.tipoDocOrigen = '07' " & _
                    "group by r.tipodocorigen, r.numdocorigen) tc " & _
                "on tc.tipoDocOrigen = dv.idDocumento and tc.serieDocOrigen = dv.idSerie and tc.numDocOrigen = dv.idDocVentas " & _
            "inner join Documentos d on dv.iddocumento = d.iddocumento " & _
            "Inner Join DocumentosCuentas DC " & _
                "On Dv.IdEmpresa = DC.IdEmpresa And Dv.IdDocumento = DC.IdDocumento " & _
            "where dv.idEmpresa = '" & glsEmpresa & "' " & _
            Cadena_Oficial & _
            Cadena_Transferido & _
            "and month(dv.fecEmision) = " & Val(strMes) & " and year(dv.fecEmision) = " & Val(strAno) & " And Dv.IndTransGratuitaMP <> '1' " & _
            "order by dv.iddocumento, dv.idSerie, dv.iddocventas "
    
    If rsDocumentos.State = 1 Then rsDocumentos.Close
    rsDocumentos.Open strSQL, Cn, adOpenKeyset, adLockOptimistic
    
    cta12soles = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_SOLES", True)
    cta12dolares = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_DOLARES", True)
    
    cta12soles_relacionada = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_SOLES_RELACIONADA", True)
    cta12dolares_relacionada = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_DOLARES_RELACIONADA", True)
        
    cta40IGV = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_40", True)
    cta40IGVPorPagar = traerCampo("parametros", "valParametro", "glsParametro", "IGV_POR_PAGAR_TG", True)
    cta70Detalle = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_70", True)
    CIVAPCuenta = traerCampo("parametros", "valParametro", "glsParametro", "IVAP_CUENTA", True)
    
    ncorrel = 0
    If rsDocumentos.RecordCount <> 0 Then
        
        CSqlC = "Select p.CtaContable_Relacionada,p.CtaContable, d.idProducto, Left(d.glsProducto,150) GlsProducto, d.Cantidad,d.TotalVVNeto," & _
                "d.TotalIGVNeto,d.TotalPVNeto,d.iddocumento,d.iddocventas,d.idempresa," & _
                "d.idserie,d.IdCentroCosto,Dv.IndTransGratuita, Dv.IndTransGratuitaMP,DV.TotalIVAP " & _
                "From DocVentas Dv " & _
                "inner join Documentos dc " & _
                    "on dv.iddocumento = dc.iddocumento " & _
                "Inner Join DocVentasDet D " & _
                    "On Dv.IdEmpresa = D.IdEmpresa And Dv.IdDocVentas = D.IdDocVentas And dv.idserie = d.idserie And dv.iddocumento = d.iddocumento " & _
                "Inner Join Productos p " & _
                    "On D.IdProducto = P.IdProducto And d.idempresa = p.idempresa " & _
                "where dv.idEmpresa = '" & glsEmpresa & "' " & _
                Cadena_Oficial2 & _
                Cadena_Transferido & _
                "and month(dv.fecEmision) = " & Val(strMes) & " and year(dv.fecEmision) = " & Val(strAno) & " " & _
                "order by dv.iddocumento,dv.idserie,dv.iddocventas,d.idProducto "
        
        If rsDetalle.State = adStateOpen Then rsDetalle.Close
        rsDetalle.Open CSqlC, Cn, adOpenKeyset, adLockOptimistic
        If rsAsientosContables.State = 1 Then rsAsientosContables.Close
    
        rsAsientosContables.Fields.Append "Item", adInteger, 2, adFldRowID
        rsAsientosContables.Fields.Append "idComprobante", adVarChar, 9, adFldRowID
        rsAsientosContables.Fields.Append "idPeriodo", adDouble, 11, adFldIsNullable
        rsAsientosContables.Fields.Append "ValItem", adVarChar, 4, adFldRowID
        rsAsientosContables.Fields.Append "IDOrigen", adVarChar, 2, adFldIsNullable
        rsAsientosContables.Fields.Append "FecCompro", adVarChar, 30, adFldRowID
        rsAsientosContables.Fields.Append "GlsDetalle", adVarChar, 250, adFldIsNullable
        rsAsientosContables.Fields.Append "idCtaContable", adVarChar, 150, adFldRowID
        rsAsientosContables.Fields.Append "idGasto", adVarChar, 4, adFldIsNullable
        rsAsientosContables.Fields.Append "NumCheque", adDouble, adFldRowID
        rsAsientosContables.Fields.Append "NumReferencia", adVarChar, 45, adFldIsNullable
        rsAsientosContables.Fields.Append "TotalImporteS", adDouble, adFldRowID
        rsAsientosContables.Fields.Append "TotalImporteD", adDouble, adFldIsNullable
        rsAsientosContables.Fields.Append "idMoneda", adChar, 3, adFldRowID
        rsAsientosContables.Fields.Append "ValorTipoCambio", adDouble, adFldIsNullable
        rsAsientosContables.Fields.Append "idTipoDoc", adVarChar, 3, adFldRowID
        rsAsientosContables.Fields.Append "idTipoDH", adChar, 1, adFldIsNullable
        rsAsientosContables.Fields.Append "idCosto", adVarChar, 8, adFldRowID
        rsAsientosContables.Fields.Append "Destino", adVarChar, 1, adFldIsNullable
        rsAsientosContables.Fields.Append "CtaAuxiliar", adVarChar, 11, adFldRowID
        rsAsientosContables.Fields.Append "F3AUTOMATICO", adDouble, 11, adFldRowID
        rsAsientosContables.Fields.Append "F3ORIGAUTO", adDouble, 11, adFldRowID
        rsAsientosContables.Fields.Append "F4FECVENC", adDouble, 11, adFldRowID
        rsAsientosContables.Fields.Append "F3RUC", adDouble, 11, adFldRowID
        rsAsientosContables.Fields.Append "SerieDoc", adChar, 4, adFldRowID
        rsAsientosContables.Fields.Append "f3nummov", adVarChar, 11, adFldRowID
        rsAsientosContables.Fields.Append "f2tipdoc", adVarChar, 2, adFldRowID
        rsAsientosContables.Fields.Append "obra", adVarChar, 85, adFldRowID
        rsAsientosContables.Fields.Append "F3FECHADOCUM", adVarChar, 30, adFldRowID
        rsAsientosContables.Fields.Append "F3FECHACOMP", adVarChar, 30, adFldRowID
        rsAsientosContables.Fields.Append "F3NUMEROCOMP", adVarChar, 250, adFldRowID
        rsAsientosContables.Fields.Append "indAfecto", adVarChar, 1, adFldRowID
        rsAsientosContables.Fields.Append "ValImporte", adDouble, adFldRowID
        rsAsientosContables.Fields.Append "IdCtaCorriente", adInteger, 2, adFldRowID
        rsAsientosContables.Fields.Append "IdMovBancosCab", adInteger, 2, adFldRowID
        rsAsientosContables.Fields.Append "TipoContable", adVarChar, 2, adFldRowID
        rsAsientosContables.Fields.Append "glosa", adVarChar, 255, adFldRowID
        rsAsientosContables.Open
        
        rsDocumentos.MoveFirst
        Do While Not rsDocumentos.EOF
            sHaberSoles = 0#: sDebeSoles = 0#
            sHaberDolares = 0#: sDebeDolares = 0#
            ncorrel = ncorrel + 1
            
            'If Trim("" & rsDocumentos.Fields("iddocventas")) = "00003518" Then MsgBox ""
            'If Trim("" & rsDocumentos.Fields("IndTransGratuita")) <> "1" And Trim("" & rsDocumentos.Fields("IndTransGratuitaMP")) <> "1" Then
            '    TOTAL
            'ElseIf rsDocumentos.Fields("estDocVentas") & "" = "ANU" Then
                TOTAL
            'End If
            
            If rsDocumentos.Fields("estDocVentas") & "" <> "ANU" And Trim("" & rsDocumentos.Fields("IndTransGratuita")) <> "1" And Trim("" & rsDocumentos.Fields("IndTransGratuitaMP")) <> "1" Then
                IGV
                IVAP
                'If Trim("" & rsDocumentos.Fields("IndTransGratuita")) = "1" Or Trim("" & rsDocumentos.Fields("IndTransGratuitaMP")) = "1" Then
                '    IGVPorPagar
                'End If
                
                'DESCUENTOS
                If Trim("" & rsDocumentos.Fields("IndTransGratuita")) <> "1" And Trim("" & rsDocumentos.Fields("IndTransGratuitaMP")) <> "1" Then
                    Elemen = 3
                    DETALLE
                End If
            End If
            
            rsDocumentos.MoveNext
            
        Loop
        
        If Not IndGeneraInt Then MsgBox "Fin del proceso.", vbInformation, App.Title
    Else
        If Not IndGeneraInt Then MsgBox "No hay registros para transferir", vbInformation, App.Title
    End If
    
    Me.MousePointer = 1
        
    Exit Sub
ERR:
    IndGeneraInt = False
    If rsTemp.State = 1 Then rsTemp.Close: Set rsTemp = Nothing
    Me.MousePointer = 1
    MsgBox ERR.Description, vbInformation, App.Title
    Exit Sub
    Resume
End Sub

Private Function CAPTURA_FECHA_FIN(pmes As String, panno As String)
Dim CFecha  As String

    Select Case pmes
        Case "01": CFecha = Format("31/" & "01/" & panno, "DD/MM/YYYY")
        Case "02":
            If (panno Mod 4) <> 0 Then
                CFecha = Format("28/" & "02/" & panno, "DD/MM/YYYY")
            Else
                CFecha = Format("29/" & "02/" & panno, "DD/MM/YYYY")
            End If
        Case "03": CFecha = Format("31/" & "03/" & panno, "DD/MM/YYYY")
        Case "04": CFecha = Format("30/" & "04/" & panno, "DD/MM/YYYY")
        Case "05": CFecha = Format("31/" & "05/" & panno, "DD/MM/YYYY")
        Case "06": CFecha = Format("30/" & "06/" & panno, "DD/MM/YYYY")
        Case "07": CFecha = Format("31/" & "07/" & panno, "DD/MM/YYYY")
        Case "08": CFecha = Format("31/" & "08/" & panno, "DD/MM/YYYY")
        Case "09": CFecha = Format("30/" & "09/" & panno, "DD/MM/YYYY")
        Case "10": CFecha = Format("31/" & "10/" & panno, "DD/MM/YYYY")
        Case "11": CFecha = Format("30/" & "11/" & panno, "DD/MM/YYYY")
        Case "12": CFecha = Format("31/" & "12/" & panno, "DD/MM/YYYY")
    End Select
    CAPTURA_FECHA_FIN = CFecha

End Function

Private Sub DELETEREC_N(PTabla As String, pconexion As ADODB.Connection)
On Error Resume Next
    
    pconexion.Execute ("DELETE * FROM " & PTabla)
    
End Sub

Private Sub TOTAL()
Dim wdetalle    As String
Dim ctercon     As String
Dim cnomtip     As String
Dim ntotfac     As Double

    NItem = NItem + 1
    rsAsientosContables.AddNew
    rsAsientosContables.Fields("ITEM") = Format(NItem, "0000")
    rsAsientosContables.Fields("idPeriodo") = strAno & strMes
    rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(ncorrel, "0000000")
    rsAsientosContables.Fields("ValItem") = "0001"
    
    If Trim("" & rsDocumentos.Fields("IndTransGratuita")) <> "1" And Trim("" & rsDocumentos.Fields("IndTransGratuitaMP")) <> "1" Then
        rsAsientosContables.Fields("Glosa") = Trim("" & rsDocumentos.Fields("GlsCliente")) & " " & Trim("" & rsDocumentos.Fields("idDocumento")) & "/" & Trim("" & rsDocumentos.Fields("idSerie")) & "/" & Trim("" & rsDocumentos.Fields("idDocventas"))
    Else
        rsAsientosContables.Fields("Glosa") = Trim("" & rsDocumentos.Fields("GlsCliente")) & " " & Trim("" & rsDocumentos.Fields("idDocumento")) & "/" & Trim("" & rsDocumentos.Fields("idSerie")) & "/" & Trim("" & rsDocumentos.Fields("idDocventas")) & " - Venta Gratuita"
    End If
'    If rsDocumentos.Fields("idMoneda") & "" = "PEN" Then
'        If Len(Trim("" & traerCampo("empresasrelacionadas", "idpersona", "idpersona", Trim("" & rsDocumentos.Fields("idPerCliente")), True))) = 0 Then
'            rsAsientosContables.Fields("idCtaContable") = Trim(cta12soles)
'        Else
'            rsAsientosContables.Fields("idCtaContable") = Trim(cta12soles_relacionada)
'        End If
'    Else
'        If Len(Trim("" & traerCampo("empresasrelacionadas", "idpersona", "idpersona", Trim("" & rsDocumentos.Fields("idPerCliente")), True))) = 0 Then
'            rsAsientosContables.Fields("idCtaContable") = Trim(cta12dolares)
'        Else
'            rsAsientosContables.Fields("idCtaContable") = Trim(cta12dolares_relacionada)
'        End If
'    End If
    
    rsAsientosContables.Fields("idCtaContable") = Trim("" & rsDocumentos.Fields("IdCtaCliente"))
    
    If rsDocumentos.Fields("estDocventas") & "" <> "ANU" And Trim("" & rsDocumentos.Fields("IndTransGratuita")) <> "1" And Trim("" & rsDocumentos.Fields("IndTransGratuitaMP")) <> "1" Then
        If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
            If Val(rsDocumentos.Fields("TotalPrecioVenta") & "") < 0 Then
                ntotfac = Format(Val("" & rsDocumentos.Fields("TotalPrecioVenta")) * -1, "0.00")
            Else
                ntotfac = Format(Val("" & rsDocumentos.Fields("TotalPrecioVenta")), "0.00")
            End If
        Else
            ntotfac = Format(Val("" & rsDocumentos.Fields("TotalPrecioVenta")), "0.00")
        End If
    
        If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
            rsAsientosContables.Fields("TotalImporteS") = ntotfac
            If Val("" & rsDocumentos.Fields("TipoCambio")) <> 0# Then
                rsAsientosContables.Fields("TotalImporteD") = Format(Val(ntotfac) / Val("" & rsDocumentos.Fields("TipoCambio")), "#0.00")
            End If
        Else
            rsAsientosContables.Fields("TotalImporteS") = Format(Val(ntotfac) * Val("" & rsDocumentos.Fields("TipoCambio")), "#0.00")
            rsAsientosContables.Fields("TotalImporteD") = ntotfac
        End If
        
        If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
            rsAsientosContables.Fields("idTipoDH") = "H"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntothab = Format(Val("" & rsAsientosContables.Fields("TotalImporteS")) / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntothab = Format(Val("" & rsAsientosContables.Fields("TotalImporteS")) * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        Else
            rsAsientosContables.Fields("idTipoDH") = "D"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntotdeb = Format(Val("" & rsAsientosContables.Fields("TotalImporteS")) / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntotdeb = Format(Val("" & rsAsientosContables.Fields("TotalImporteS")) * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        End If
    Else
        rsAsientosContables.Fields("TotalImporteS") = 0#
        rsAsientosContables.Fields("TotalImporteD") = 0#
        If rsDocumentos.Fields("iddocumento") = "07" Then
            rsAsientosContables.Fields("idTipoDH") = "H"
        Else
            rsAsientosContables.Fields("idTipoDH") = "D"
        End If
    End If

    cnomtip = ""
    If docofi = "S" Then
        Select Case rsDocumentos.Fields("iddocumento") & ""
            Case "01": cnomtip = "Fac"
            Case "03": cnomtip = "Bol"
            Case "07": cnomtip = "Cre"
            Case "08": cnomtip = "Deb"
            Case "12": cnomtip = "T/C"
            Case "90": cnomtip = "Npd"
        End Select
    Else
        Select Case rsDocumentos.Fields("iddocumento") & ""
            Case "01": cnomtip = "Fac"
            Case "03": cnomtip = "Bol"
            Case "07": cnomtip = "Cre"
            Case "08": cnomtip = "Deb"
            Case "12": cnomtip = "T/C"
        End Select
    End If
    
    If Trim("" & rsDocumentos.Fields("IndTransGratuita")) <> "1" And Trim("" & rsDocumentos.Fields("IndTransGratuitaMP")) <> "1" Then
        wdetalle = cnomtip & rsDocumentos.Fields("idSerie") & "/" & rsDocumentos.Fields("iddocVentas") & Space(2) & left(rsDocumentos.Fields("glsCliente"), 50)
    Else
        wdetalle = cnomtip & rsDocumentos.Fields("idSerie") & "/" & rsDocumentos.Fields("iddocVentas") & Space(2) & left(rsDocumentos.Fields("glsCliente"), 50) & " - Venta Gratuita"
    End If
    

    rsAsientosContables.Fields("ValorTipoCambio") = Val(rsDocumentos.Fields("TipoCambio") & "")
    rsAsientosContables.Fields("idMoneda") = rsDocumentos.Fields("idMoneda") & ""
    rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
    rsAsientosContables.Fields("FecCompro") = rsDocumentos.Fields("FecEmision") & ""
    If rsDocumentos.Fields("estDocVentas") & "" <> "ANU" Then
        rsAsientosContables.Fields("GlsDetalle") = wdetalle
    Else
        rsAsientosContables.Fields("GlsDetalle") = "A N U L A D A"
    End If
    rsAsientosContables.Fields("idTipoDoc") = cnomtip
    rsAsientosContables.Fields("NumCheque") = rsDocumentos.Fields("iddocventas") & ""
    rsAsientosContables.Fields("NumReferencia") = rsDocumentos.Fields("iddocventas") & ""
    rsAsientosContables.Fields("SerieDoc") = rsDocumentos.Fields("idSerie") & ""
    rsAsientosContables.Fields("CtaAuxiliar") = rsDocumentos.Fields("RUCCliente") & ""
    
    If rsAsientosContables!idTipoDH = "H" Then
        sHaberSoles = sHaberSoles + rsAsientosContables.Fields("TotalImporteS")
        sHaberDolares = sHaberDolares + rsAsientosContables.Fields("TotalImporteD")
    Else
        sDebeSoles = sDebeSoles + rsAsientosContables.Fields("TotalImporteS")
        sDebeDolares = sDebeDolares + rsAsientosContables.Fields("TotalImporteD")
    End If
    
    '--- Agregado el tipo de contabilidad a trabajar tributaria o contable
    rsAsientosContables.Fields("TipoContable") = right(CmbOpciones.Text, 2)
    rsAsientosContables.Update
    
End Sub

Private Sub TOTALRally()
Dim wdetalle    As String
Dim ctercon     As String
Dim cnomtip     As String
Dim ntotfac     As Double

    NItem = NItem + 1
    rsAsientosContables.AddNew
    rsAsientosContables.Fields("ITEM") = Format(NItem, "0000")
    rsAsientosContables.Fields("idPeriodo") = strAno & strMes
    rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(ncorrel, "0000000")
    rsAsientosContables.Fields("ValItem") = "0001"
    rsAsientosContables.Fields("Glosa") = Trim("" & rsDocumentos.Fields("GlsCliente")) & " " & Trim("" & rsDocumentos.Fields("idDocumento")) & "/" & Trim("" & rsDocumentos.Fields("idSerie")) & "/" & Trim("" & rsDocumentos.Fields("idDocventas"))
    
    If rsDocumentos.Fields("idMoneda") & "" = "PEN" Then
        If Len(Trim("" & traerCampo("empresasrelacionadas", "idpersona", "idpersona", Trim("" & rsDocumentos.Fields("idPerCliente")), True))) = 0 Then
            rsAsientosContables.Fields("idCtaContable") = Trim(cta12soles)
        Else
            rsAsientosContables.Fields("idCtaContable") = Trim(cta12soles_relacionada)
        End If
    Else
        If Len(Trim("" & traerCampo("empresasrelacionadas", "idpersona", "idpersona", Trim("" & rsDocumentos.Fields("idPerCliente")), True))) = 0 Then
            rsAsientosContables.Fields("idCtaContable") = Trim(cta12dolares)
        Else
            rsAsientosContables.Fields("idCtaContable") = Trim(cta12dolares_relacionada)
        End If
    End If

    If rsDocumentos.Fields("estDocventas") & "" <> "ANU" Then
        If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
            If Val(rsDocumentos.Fields("TotalPrecioVenta") & "") < 0 Then
                ntotfac = Format(Val("" & rsDocumentos.Fields("TotalPrecioVenta")) * -1, "0.00")
            Else
                ntotfac = Format(Val("" & rsDocumentos.Fields("TotalPrecioVenta")), "0.00")
            End If
        Else
            ntotfac = Format(Val("" & rsDocumentos.Fields("TotalPrecioVenta")), "0.00")
        End If
    
        If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
            rsAsientosContables.Fields("TotalImporteS") = ntotfac
            If Val("" & rsDocumentos.Fields("TipoCambio")) <> 0# Then
                rsAsientosContables.Fields("TotalImporteD") = Format(Val(ntotfac) / Val("" & rsDocumentos.Fields("TipoCambio")), "#0.00")
            End If
        Else
            rsAsientosContables.Fields("TotalImporteS") = Format(Val(ntotfac) * Val("" & rsDocumentos.Fields("TipoCambio")), "#0.00")
            rsAsientosContables.Fields("TotalImporteD") = ntotfac
        End If
        
        If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
            rsAsientosContables.Fields("idTipoDH") = "H"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntothab = Format(Val("" & rsAsientosContables.Fields("TotalImporteS")) / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntothab = Format(Val("" & rsAsientosContables.Fields("TotalImporteS")) * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        Else
            rsAsientosContables.Fields("idTipoDH") = "D"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntotdeb = Format(Val("" & rsAsientosContables.Fields("TotalImporteS")) / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntotdeb = Format(Val("" & rsAsientosContables.Fields("TotalImporteS")) * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        End If
    Else
        rsAsientosContables.Fields("TotalImporteS") = 0#
        rsAsientosContables.Fields("TotalImporteD") = 0#
        If rsDocumentos.Fields("iddocumento") = "07" Then
            rsAsientosContables.Fields("idTipoDH") = "H"
        Else
            rsAsientosContables.Fields("idTipoDH") = "D"
        End If
    End If

    cnomtip = ""
    If docofi = "S" Then
        Select Case rsDocumentos.Fields("iddocumento") & ""
            Case "01": cnomtip = "Fac"
            Case "03": cnomtip = "Bol"
            Case "07": cnomtip = "Cre"
            Case "08": cnomtip = "Deb"
            Case "12": cnomtip = "T/C"
            Case "90": cnomtip = "Npd"
        End Select
    Else
        Select Case rsDocumentos.Fields("iddocumento") & ""
            Case "01": cnomtip = "Fac"
            Case "03": cnomtip = "Bol"
            Case "07": cnomtip = "Cre"
            Case "08": cnomtip = "Deb"
            Case "12": cnomtip = "T/C"
        End Select
    End If
        
    wdetalle = cnomtip & rsDocumentos.Fields("idSerie") & "/" & rsDocumentos.Fields("iddocVentas") & Space(2) & left(rsDocumentos.Fields("glsCliente"), 50)

    rsAsientosContables.Fields("ValorTipoCambio") = Val(rsDocumentos.Fields("TipoCambio") & "")
    rsAsientosContables.Fields("idMoneda") = rsDocumentos.Fields("idMoneda") & ""
    rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
    rsAsientosContables.Fields("FecCompro") = rsDocumentos.Fields("FecEmision") & ""
    If rsDocumentos.Fields("estDocVentas") & "" <> "ANU" Then
        rsAsientosContables.Fields("GlsDetalle") = wdetalle
    Else
        rsAsientosContables.Fields("GlsDetalle") = "A N U L A D A"
    End If
    rsAsientosContables.Fields("idTipoDoc") = cnomtip
    rsAsientosContables.Fields("NumCheque") = rsDocumentos.Fields("iddocventas") & ""
    rsAsientosContables.Fields("NumReferencia") = rsDocumentos.Fields("iddocventas") & ""
    rsAsientosContables.Fields("SerieDoc") = rsDocumentos.Fields("idSerie") & ""
    rsAsientosContables.Fields("CtaAuxiliar") = rsDocumentos.Fields("RUCCliente") & ""
    
    If rsAsientosContables!idTipoDH = "H" Then
        sHaberSoles = sHaberSoles + rsAsientosContables.Fields("TotalImporteS")
        sHaberDolares = sHaberDolares + rsAsientosContables.Fields("TotalImporteD")
    Else
        sDebeSoles = sDebeSoles + rsAsientosContables.Fields("TotalImporteS")
        sDebeDolares = sDebeDolares + rsAsientosContables.Fields("TotalImporteD")
    End If
    
    '--- Agregado el tipo de contabilidad a trabajar tributaria o contable
    rsAsientosContables.Fields("TipoContable") = right(CmbOpciones.Text, 2)
    rsAsientosContables.Update
    
End Sub

Private Sub IGV()
Dim wfecha      As Variant
Dim wdetalle    As String
Dim cnomtip     As String
Dim NIGV        As Double

    If Val(rsDocumentos.Fields("TotalIGVVenta") & "") <> 0 Then
        rsAsientosContables.AddNew
        
        NItem = NItem + 1
        rsAsientosContables.Fields("ITEM") = Format(NItem, "0000")
        rsAsientosContables.Fields("Glosa") = Trim("" & rsDocumentos.Fields("GlsCliente")) & " " & Trim("" & rsDocumentos.Fields("idDocumento")) & "/" & Trim("" & rsDocumentos.Fields("idSerie")) & "/" & Trim("" & rsDocumentos.Fields("idDocventas"))
        rsAsientosContables.Fields("idPeriodo") = strAno & strMes
        rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(ncorrel, "0000000")
        If Trim("" & rsDocumentos.Fields("IndTransGratuita")) = "1" Then
            rsAsientosContables.Fields("ValItem") = "0001"
        Else
            rsAsientosContables.Fields("ValItem") = "0002"
        End If
        rsAsientosContables.Fields("idCtaContable") = cta40IGV
        TC = Val(rsDocumentos.Fields("TipoCambio") & "")
        
        NIGV = 0#: NTotalS = 0#: ntotbase = 0#
        If rsDocumentos.Fields("iddocumento") = "07" Then
            If Val(rsDocumentos.Fields("TotalIGVVenta") & "") < 0 Then
                NIGV = Format(Val("" & rsDocumentos.Fields("TotalIGVVenta")) * -1, "0.00")
            Else
                NIGV = Format(Val("" & rsDocumentos.Fields("TotalIGVVenta")), "0.00")
            End If
        Else
            NIGV = Format(Val(rsDocumentos.Fields("TotalIGVVenta") & ""), "0.00")
        End If
        
        If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
            rsAsientosContables.Fields("TotalImporteS") = Format(NIGV, "0.00")
            If Val(rsDocumentos.Fields("TipoCambio") & "") <> 0# Then
                rsAsientosContables.Fields("TotalImporteD") = Format(NIGV / Val(rsDocumentos.Fields("TipoCambio") & ""), "0.00")
            End If
        Else
            rsAsientosContables.Fields("TotalImporteS") = Format(NIGV * Val(rsDocumentos.Fields("TipoCambio") & ""), "0.00")
            rsAsientosContables.Fields("TotalImporteD") = Format(NIGV, "0.00")
        End If
        
        cnomtip = ""
        If docofi = "S" Then
            Select Case rsDocumentos.Fields("iddocumento")
                Case "01": cnomtip = "Fac"
                Case "03": cnomtip = "Bol"
                Case "07": cnomtip = "Cre"
                Case "08": cnomtip = "Deb"
                Case "12": cnomtip = "T/C"
                Case "90": cnomtip = "Npd"
            End Select
        Else
            Select Case rsDocumentos.Fields("iddocumento")
                Case "01": cnomtip = "Fac"
                Case "03": cnomtip = "Bol"
                Case "07": cnomtip = "Cre"
                Case "08": cnomtip = "Deb"
                Case "12": cnomtip = "T/C"
            End Select
        End If
        
        wdetalle = "" & cnomtip & rsDocumentos.Fields("idserie") & "/" & rsDocumentos.Fields("iddocventas") & Space(2) & left(rsDocumentos.Fields("glsCliente"), 50)
        
        If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
            rsAsientosContables.Fields("idTipoDH") = "D"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntotdeb = ntotdeb + Format(NIGV / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntotdeb = ntotdeb + Format(NIGV * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        Else
            rsAsientosContables.Fields("idTipoDH") = "H"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntothab = ntothab + Format(NIGV / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntothab = ntothab + Format(NIGV * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        End If
    
        rsAsientosContables.Fields("ValorTipoCambio") = Val("" & rsDocumentos.Fields("TipoCambio"))
        rsAsientosContables.Fields("idMoneda") = rsDocumentos.Fields("idMoneda") & ""
        rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
        rsAsientosContables.Fields("FecCompro") = rsDocumentos.Fields("FecEmision") & ""
        rsAsientosContables.Fields("GlsDetalle") = wdetalle
        rsAsientosContables.Fields("idTipoDoc") = cnomtip
        rsAsientosContables.Fields("NumCheque") = rsDocumentos.Fields("iddocventas") & ""
        rsAsientosContables.Fields("NumReferencia") = rsDocumentos.Fields("iddocventas") & ""
        rsAsientosContables.Fields("SerieDoc") = rsDocumentos.Fields("idSerie") & ""
        
        
        
        If rsAsientosContables!idTipoDH = "H" Then
            sHaberSoles = sHaberSoles + rsAsientosContables.Fields("TotalImporteS")
            sHaberDolares = sHaberDolares + rsAsientosContables.Fields("TotalImporteD")
        Else
            sDebeSoles = sDebeSoles + rsAsientosContables.Fields("TotalImporteS")
            sDebeDolares = sDebeDolares + rsAsientosContables.Fields("TotalImporteD")
        End If
        
        '--- Agregado el tipo de contabilidad a trabajar tributaria o contable
        rsAsientosContables.Fields("TipoContable") = right(CmbOpciones.Text, 2)
        rsAsientosContables.Update
    End If

End Sub

Private Sub IVAP()
Dim wfecha      As Variant
Dim wdetalle    As String
Dim cnomtip     As String
Dim NIGV        As Double

    If Val(rsDocumentos.Fields("TotalIVAP") & "") <> 0 Then
        rsAsientosContables.AddNew
        
        NItem = NItem + 1
        rsAsientosContables.Fields("ITEM") = Format(NItem, "0000")
        rsAsientosContables.Fields("Glosa") = Trim("" & rsDocumentos.Fields("GlsCliente")) & " " & Trim("" & rsDocumentos.Fields("idDocumento")) & "/" & Trim("" & rsDocumentos.Fields("idSerie")) & "/" & Trim("" & rsDocumentos.Fields("idDocventas"))
        rsAsientosContables.Fields("idPeriodo") = strAno & strMes
        rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(ncorrel, "0000000")
        If Trim("" & rsDocumentos.Fields("IndTransGratuita")) = "1" Then
            rsAsientosContables.Fields("ValItem") = "0002"
        Else
            rsAsientosContables.Fields("ValItem") = "0003"
        End If
        rsAsientosContables.Fields("idCtaContable") = CIVAPCuenta
        TC = Val(rsDocumentos.Fields("TipoCambio") & "")
        
        NIGV = 0#: NTotalS = 0#: ntotbase = 0#
        If rsDocumentos.Fields("iddocumento") = "07" Then
            If Val(rsDocumentos.Fields("TotalIVAP") & "") < 0 Then
                NIGV = Format(Val("" & rsDocumentos.Fields("TotalIVAP")) * -1, "0.00")
            Else
                NIGV = Format(Val("" & rsDocumentos.Fields("TotalIVAP")), "0.00")
            End If
        Else
            NIGV = Format(Val(rsDocumentos.Fields("TotalIVAP") & ""), "0.00")
        End If
        
        If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
            rsAsientosContables.Fields("TotalImporteS") = Format(NIGV, "0.00")
            If Val(rsDocumentos.Fields("TipoCambio") & "") <> 0# Then
                rsAsientosContables.Fields("TotalImporteD") = Format(NIGV / Val(rsDocumentos.Fields("TipoCambio") & ""), "0.00")
            End If
        Else
            rsAsientosContables.Fields("TotalImporteS") = Format(NIGV * Val(rsDocumentos.Fields("TipoCambio") & ""), "0.00")
            rsAsientosContables.Fields("TotalImporteD") = Format(NIGV, "0.00")
        End If
        
        cnomtip = ""
        If docofi = "S" Then
            Select Case rsDocumentos.Fields("iddocumento")
                Case "01": cnomtip = "Fac"
                Case "03": cnomtip = "Bol"
                Case "07": cnomtip = "Cre"
                Case "08": cnomtip = "Deb"
                Case "12": cnomtip = "T/C"
                Case "90": cnomtip = "Npd"
            End Select
        Else
            Select Case rsDocumentos.Fields("iddocumento")
                Case "01": cnomtip = "Fac"
                Case "03": cnomtip = "Bol"
                Case "07": cnomtip = "Cre"
                Case "08": cnomtip = "Deb"
                Case "12": cnomtip = "T/C"
            End Select
        End If
        
        wdetalle = "" & cnomtip & rsDocumentos.Fields("idserie") & "/" & rsDocumentos.Fields("iddocventas") & Space(2) & left(rsDocumentos.Fields("glsCliente"), 50)
        
        If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
            rsAsientosContables.Fields("idTipoDH") = "D"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntotdeb = ntotdeb + Format(NIGV / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntotdeb = ntotdeb + Format(NIGV * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        Else
            rsAsientosContables.Fields("idTipoDH") = "H"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntothab = ntothab + Format(NIGV / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntothab = ntothab + Format(NIGV * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        End If
    
        rsAsientosContables.Fields("ValorTipoCambio") = Val("" & rsDocumentos.Fields("TipoCambio"))
        rsAsientosContables.Fields("idMoneda") = rsDocumentos.Fields("idMoneda") & ""
        rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
        rsAsientosContables.Fields("FecCompro") = rsDocumentos.Fields("FecEmision") & ""
        rsAsientosContables.Fields("GlsDetalle") = wdetalle
        rsAsientosContables.Fields("idTipoDoc") = cnomtip
        rsAsientosContables.Fields("NumCheque") = rsDocumentos.Fields("iddocventas") & ""
        rsAsientosContables.Fields("SerieDoc") = rsDocumentos.Fields("idSerie") & ""
        
        
        
        If rsAsientosContables!idTipoDH = "H" Then
            sHaberSoles = sHaberSoles + rsAsientosContables.Fields("TotalImporteS")
            sHaberDolares = sHaberDolares + rsAsientosContables.Fields("TotalImporteD")
        Else
            sDebeSoles = sDebeSoles + rsAsientosContables.Fields("TotalImporteS")
            sDebeDolares = sDebeDolares + rsAsientosContables.Fields("TotalImporteD")
        End If
        
        '--- Agregado el tipo de contabilidad a trabajar tributaria o contable
        rsAsientosContables.Fields("TipoContable") = right(CmbOpciones.Text, 2)
        rsAsientosContables.Update
    End If

End Sub

Private Sub IGVPorPagar()
Dim wfecha      As Variant
Dim wdetalle    As String
Dim cnomtip     As String
Dim NIGV        As Double

    If Val(rsDocumentos.Fields("TotalIGVVenta") & "") <> 0 Then
        rsAsientosContables.AddNew
        
        NItem = NItem + 1
        rsAsientosContables.Fields("ITEM") = Format(NItem, "0000")
        rsAsientosContables.Fields("Glosa") = Trim("" & rsDocumentos.Fields("GlsCliente")) & " " & Trim("" & rsDocumentos.Fields("idDocumento")) & "/" & Trim("" & rsDocumentos.Fields("idSerie")) & "/" & Trim("" & rsDocumentos.Fields("idDocventas"))
        rsAsientosContables.Fields("idPeriodo") = strAno & strMes
        rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(ncorrel, "0000000")
        rsAsientosContables.Fields("ValItem") = "0002"
        rsAsientosContables.Fields("idCtaContable") = cta40IGVPorPagar
        TC = Val(rsDocumentos.Fields("TipoCambio") & "")
        
        NIGV = 0#: NTotalS = 0#: ntotbase = 0#
        If rsDocumentos.Fields("iddocumento") = "07" Then
            If Val(rsDocumentos.Fields("TotalIGVVenta") & "") < 0 Then
                NIGV = Format(Val("" & rsDocumentos.Fields("TotalIGVVenta")) * -1, "0.00")
            Else
                NIGV = Format(Val("" & rsDocumentos.Fields("TotalIGVVenta")), "0.00")
            End If
        Else
            NIGV = Format(Val(rsDocumentos.Fields("TotalIGVVenta") & ""), "0.00")
        End If
        
        If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
            rsAsientosContables.Fields("TotalImporteS") = Format(NIGV, "0.00")
            If Val(rsDocumentos.Fields("TipoCambio") & "") <> 0# Then
                rsAsientosContables.Fields("TotalImporteD") = Format(NIGV / Val(rsDocumentos.Fields("TipoCambio") & ""), "0.00")
            End If
        Else
            rsAsientosContables.Fields("TotalImporteS") = Format(NIGV * Val(rsDocumentos.Fields("TipoCambio") & ""), "0.00")
            rsAsientosContables.Fields("TotalImporteD") = Format(NIGV, "0.00")
        End If
        
        cnomtip = ""
        If docofi = "S" Then
            Select Case rsDocumentos.Fields("iddocumento")
                Case "01": cnomtip = "Fac"
                Case "03": cnomtip = "Bol"
                Case "07": cnomtip = "Cre"
                Case "08": cnomtip = "Deb"
                Case "12": cnomtip = "T/C"
                Case "90": cnomtip = "Npd"
            End Select
        Else
            Select Case rsDocumentos.Fields("iddocumento")
                Case "01": cnomtip = "Fac"
                Case "03": cnomtip = "Bol"
                Case "07": cnomtip = "Cre"
                Case "08": cnomtip = "Deb"
                Case "12": cnomtip = "T/C"
            End Select
        End If
        
        wdetalle = "" & cnomtip & rsDocumentos.Fields("idserie") & "/" & rsDocumentos.Fields("iddocventas") & Space(2) & left(rsDocumentos.Fields("glsCliente"), 50)
        
        If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
            rsAsientosContables.Fields("idTipoDH") = "H"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntothab = ntothab + Format(NIGV / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntothab = ntothab + Format(NIGV * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        Else
            rsAsientosContables.Fields("idTipoDH") = "D"
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
                    ntotdeb = ntotdeb + Format(NIGV / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
                End If
            Else
                ntotdeb = ntotdeb + Format(NIGV * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            End If
        End If
    
        rsAsientosContables.Fields("ValorTipoCambio") = Val("" & rsDocumentos.Fields("TipoCambio"))
        rsAsientosContables.Fields("idMoneda") = rsDocumentos.Fields("idMoneda") & ""
        rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
        rsAsientosContables.Fields("FecCompro") = rsDocumentos.Fields("FecEmision") & ""
        rsAsientosContables.Fields("GlsDetalle") = wdetalle
        rsAsientosContables.Fields("idTipoDoc") = cnomtip
        rsAsientosContables.Fields("NumCheque") = rsDocumentos.Fields("iddocventas") & ""
        rsAsientosContables.Fields("SerieDoc") = rsDocumentos.Fields("idSerie") & ""
        
        If rsAsientosContables!idTipoDH = "H" Then
            sHaberSoles = sHaberSoles + rsAsientosContables.Fields("TotalImporteS")
            sHaberDolares = sHaberDolares + rsAsientosContables.Fields("TotalImporteD")
        Else
            sDebeSoles = sDebeSoles + rsAsientosContables.Fields("TotalImporteS")
            sDebeDolares = sDebeDolares + rsAsientosContables.Fields("TotalImporteD")
        End If
        
        '--- Agregado el tipo de contabilidad a trabajar tributaria o contable
        rsAsientosContables.Fields("TipoContable") = right(CmbOpciones.Text, 2)
        rsAsientosContables.Update
    End If

End Sub

Private Sub IgvRally()
Dim wfecha              As Variant
Dim wdetalle            As String
Dim cnomtip             As String
Dim NIGV                As Double
Dim NTotal              As Double
Dim NBase               As Double
Dim NIgvOri             As Double

    If Val(rsDocumentos.Fields("TotalIGVVenta") & "") <> 0 Then
        rsAsientosContables.AddNew
        
        NItem = NItem + 1
        rsAsientosContables.Fields("ITEM") = Format(NItem, "0000")
        
        rsAsientosContables.Fields("idPeriodo") = strAno & strMes
        rsAsientosContables.Fields("idComprobante") = glsOrigen_Contable & Format(ncorrel, "0000000")
        rsAsientosContables.Fields("ValItem") = "0002"
        rsAsientosContables.Fields("idCtaContable") = cta40IGV
        TC = Val(rsDocumentos.Fields("TipoCambio") & "")
        
        NIGV = 0#: NTotalS = 0#: ntotbase = 0#
        NTotal = 0: NBase = 0: NIgvOri = 0
        
        If "" & rsDocumentos.Fields("idMoneda") = "USD" Then
            
            NTotal = Format(Val("" & rsDocumentos.Fields("TotalPrecioVenta")) * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
            NBase = Format(NTotal / 1.18, "0.00")
            NIGV = Format(NTotal - NBase, "0.00")
            
        Else
        
            NTotal = Format(Val("" & rsDocumentos.Fields("TotalPrecioVenta")), "0.00")
            NBase = Format(NTotal / 1.18, "0.00")
            NIGV = Format(NTotal - NBase, "0.00")
        
        End If
        
        If rsDocumentos.Fields("iddocumento") = "07" Then
            If Val(rsDocumentos.Fields("TotalIGVVenta") & "") < 0 Then
                NIGV = Format(NIGV * -1, "0.00")
            Else
                NIGV = Format(NIGV, "0.00")
            End If
        Else
            NIGV = Format(NIGV, "0.00")
        End If
        
        If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
            rsAsientosContables.Fields("TotalImporteS") = Format(NIGV, "0.00")
            If Val(rsDocumentos.Fields("TipoCambio") & "") <> 0# Then
                rsAsientosContables.Fields("TotalImporteD") = Format(NIGV * Val(rsDocumentos.Fields("TipoCambio") & ""), "0.00")
            End If
        Else
            rsAsientosContables.Fields("TotalImporteS") = Format(NIGV, "0.00")
            rsAsientosContables.Fields("TotalImporteD") = Format(NIGV / Val(rsDocumentos.Fields("TipoCambio") & ""), "0.00")
        End If
        
        cnomtip = ""
        If docofi = "S" Then
            Select Case rsDocumentos.Fields("iddocumento")
                Case "01": cnomtip = "Fac"
                Case "03": cnomtip = "Bol"
                Case "07": cnomtip = "Cre"
                Case "08": cnomtip = "Deb"
                Case "12": cnomtip = "T/C"
                Case "90": cnomtip = "Npd"
            End Select
        Else
            Select Case rsDocumentos.Fields("iddocumento")
                Case "01": cnomtip = "Fac"
                Case "03": cnomtip = "Bol"
                Case "07": cnomtip = "Cre"
                Case "08": cnomtip = "Deb"
                Case "12": cnomtip = "T/C"
            End Select
        End If
        
        wdetalle = "" & cnomtip & rsDocumentos.Fields("idserie") & "/" & rsDocumentos.Fields("iddocventas") & Space(2) & left(rsDocumentos.Fields("glsCliente"), 50)
        
        If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
            rsAsientosContables.Fields("idTipoDH") = "D"
            
            ntotdeb = ntotdeb + Format(NIGV, "0.00")
            
'            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
'                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
'                    ntotdeb = ntotdeb + Format(NIGV / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
'                End If
'            Else
'                ntotdeb = ntotdeb + Format(NIGV * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
'            End If
        Else
            rsAsientosContables.Fields("idTipoDH") = "H"
            
            ntothab = ntothab + Format(NIGV, "0.00")
            
'            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
'                If Val("" & rsDocumentos.Fields("TipoCambio")) > 0 Then
'                    ntothab = ntothab + Format(NIGV / Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
'                End If
'            Else
'                ntothab = ntothab + Format(NIGV * Val("" & rsDocumentos.Fields("TipoCambio")), "0.00")
'            End If
        End If
    
        rsAsientosContables.Fields("ValorTipoCambio") = Val("" & rsDocumentos.Fields("TipoCambio"))
        rsAsientosContables.Fields("idMoneda") = rsDocumentos.Fields("idMoneda") & ""
        rsAsientosContables.Fields("IDOrigen") = glsOrigen_Contable
        rsAsientosContables.Fields("FecCompro") = rsDocumentos.Fields("FecEmision") & ""
        rsAsientosContables.Fields("GlsDetalle") = wdetalle
        rsAsientosContables.Fields("idTipoDoc") = cnomtip
        rsAsientosContables.Fields("NumCheque") = rsDocumentos.Fields("iddocventas") & ""
        rsAsientosContables.Fields("SerieDoc") = rsDocumentos.Fields("idSerie") & ""
        
        If rsAsientosContables!idTipoDH = "H" Then
            sHaberSoles = sHaberSoles + rsAsientosContables.Fields("TotalImporteS")
            sHaberDolares = sHaberDolares + rsAsientosContables.Fields("TotalImporteD")
        Else
            sDebeSoles = sDebeSoles + rsAsientosContables.Fields("TotalImporteS")
            sDebeDolares = sDebeDolares + rsAsientosContables.Fields("TotalImporteD")
        End If
        
        '--- Agregado el tipo de contabilidad a trabajar tributaria o contable
        rsAsientosContables.Fields("TipoContable") = right(CmbOpciones.Text, 2)
        rsAsientosContables.Update
    End If

End Sub

Private Sub DETALLE()
Dim wfecha      As Variant
Dim wdetalle    As String
Dim cquery      As String
Dim rsif5pla    As New ADODB.Recordset
Dim cnomtip     As String
Dim NVALVTA     As Double
Dim Elemen As Integer
    
    rsDetalle.Filter = ""
    rsDetalle.Filter = adFilterNone
    rsDetalle.MoveFirst
    rsDetalle.Sort = "iddocumento,idserie,iddocventas,idproducto"
    
    rsDetalle.Filter = ""
    rsDetalle.Filter = adFilterNone
    rsDetalle.MoveFirst
    rsDetalle.Filter = " iddocumento = '" & rsDocumentos.Fields("iddocumento") & "' and idserie = '" & rsDocumentos.Fields("idSerie") & "' and iddocventas = '" & rsDocumentos.Fields("iddocventas") & "' "
    
    Elemen = 2
    Do While Not rsDetalle.EOF
        If rsDetalle!TotalVVNeto <> 0 Then
            Elemen = Elemen + 1
            
            rsAsientosContables.AddNew
            NItem = NItem + 1
            rsAsientosContables.Fields("ITEM") = Format(NItem, "0000")
            If leeParametro("AYUDA_CENTRO_COSTO_DETALLE") = "S" Then
                rsAsientosContables!idCosto = Trim("" & rsDetalle!IdCentroCosto)
            Else
                rsAsientosContables!idCosto = Trim("" & rsDocumentos!IdCentroCosto)
            End If
            rsAsientosContables.Fields("Glosa") = Trim("" & rsDocumentos.Fields("GlsCliente")) & " " & Trim("" & rsDocumentos.Fields("idDocumento")) & "/" & Trim("" & rsDocumentos.Fields("idSerie")) & "/" & Trim("" & rsDocumentos.Fields("idDocventas"))
            rsAsientosContables!idPeriodo = strAno & strMes
            rsAsientosContables!idComprobante = glsOrigen_Contable & Format(ncorrel, "0000000")
            rsAsientosContables!ValItem = Format(Elemen, "0000")
            
            If "" & rsDetalle.Fields("ctaContable") <> "" Then
                If Len(Trim("" & traerCampo("empresasrelacionadas", "idpersona", "idpersona", Trim("" & rsDocumentos.Fields("idPerCliente")), True))) = 0 Then
                    rsAsientosContables!idCtaContable = Trim("" & rsDetalle.Fields("ctaContable"))
                Else
                    rsAsientosContables!idCtaContable = Trim("" & rsDetalle.Fields("CtaContable_Relacionada"))
                End If
            Else
                rsAsientosContables!idCtaContable = cta70Detalle
            End If
            
            NVALVTA = 0#: ntotbase = 0#
            NVALVTA = IIf(Val(Format(rsDetalle.Fields("TotalVVNeto") & "", "0.00")) < 0#, Val(Format(rsDetalle.Fields("TotalVVNeto") & "", "0.00")) * -1, Val(Format(rsDetalle.Fields("TotalVVNeto") & "", "0.00")))
            
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                rsAsientosContables!TotalImporteS = NVALVTA
                If Val("" & rsDocumentos.Fields("TipoCambio")) <> 0# Then
                    rsAsientosContables!TotalImporteD = Format(NVALVTA / Val("" & rsDocumentos.Fields("TipoCambio")), "#0.00")
                End If
            Else
                rsAsientosContables!TotalImporteS = Format(NVALVTA * Val("" & rsDocumentos.Fields("TipoCambio")), "#0.00")
                rsAsientosContables!TotalImporteD = NVALVTA
            End If
            
            cnomtip = ""
            If docofi = "S" Then
                Select Case rsDocumentos.Fields("idDocumento")
                    Case "01": cnomtip = "Fac"
                    Case "03": cnomtip = "Bol"
                    Case "07": cnomtip = "Cre"
                    Case "08": cnomtip = "Deb"
                    Case "12": cnomtip = "T/C"
                    Case "90": cnomtip = "Npd"
                End Select
            Else
                Select Case rsDocumentos.Fields("idDocumento")
                    Case "01": cnomtip = "Fac"
                    Case "03": cnomtip = "Bol"
                    Case "07": cnomtip = "Cre"
                    Case "08": cnomtip = "Deb"
                    Case "12": cnomtip = "T/C"
                End Select
            End If
            
            wdetalle = cnomtip & rsDocumentos.Fields("idSerie") & "/" & rsDocumentos.Fields("iddocventas") & Space(2) & left(rsDocumentos.Fields("glsCliente"), 50)
            
            If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
                rsAsientosContables!idTipoDH = "D"
                If rsDetalle!TotalVVNeto > 0 Then
                    If "" & rsDocumentos!idMoneda = "PEN" Then
                        If Val("" & rsDocumentos!TipoCambio) > 0 Then
                            ntotdeb = ntotdeb + Format(NVALVTA / Val("" & rsDocumentos!TipoCambio), "#0.00")
                        End If
                    Else
                        ntotdeb = ntotdeb + Format(NVALVTA * Val("" & rsDocumentos!TipoCambio), "#0.00")
                    End If
                Else
                    If "" & rsDocumentos!idMoneda = "PEN" Then
                        If Val("" & rsDocumentos!TipoCambio) > 0 Then
                            ntothab = ntothab + Format(NVALVTA / Val("" & rsDocumentos!TipoCambio), "#0.00")
                        End If
                    Else
                        ntothab = ntothab + Format(NVALVTA * Val("" & rsDocumentos!TipoCambio), "#0.00")
                    End If
                End If
            Else
                rsAsientosContables!idTipoDH = "H"
                If rsDetalle!TotalVVNeto > 0 Then
                    If "" & rsDocumentos!idMoneda = "PEN" Then
                        If Val("" & rsDocumentos!TipoCambio) > 0 Then
                            ntothab = ntothab + Format(NVALVTA / Val("" & rsDocumentos!TipoCambio), "#0.00")
                        End If
                    Else
                        ntothab = ntothab + Format(NVALVTA * Val("" & rsDocumentos!TipoCambio), "#0.00")
                    End If
                Else
                    rsAsientosContables!idTipoDH = "D"
                    If "" & rsDocumentos!idMoneda = "PEN" Then
                        If Val("" & rsDocumentos!TipoCambio) > 0 Then
                            ntotdeb = ntotdeb + Format(NVALVTA / Val("" & rsDocumentos!TipoCambio), "#0.00")
                        End If
                    Else
                        ntotdeb = ntotdeb + Format(NVALVTA * Val("" & rsDocumentos!TipoCambio), "#0.00")
                    End If
                End If
            End If
            rsAsientosContables!ValorTipoCambio = Val("" & rsDocumentos!TipoCambio)
            rsAsientosContables!idMoneda = rsDocumentos.Fields("idMoneda") & ""
            rsAsientosContables!IDOrigen = glsOrigen_Contable
            rsAsientosContables!FecCompro = rsDocumentos.Fields("FecEmision") & ""
            rsAsientosContables!GlsDetalle = wdetalle
            rsAsientosContables.Fields("idTipoDoc") = cnomtip
            rsAsientosContables!NumCheque = rsDocumentos.Fields("iddocventas") & ""
            rsAsientosContables.Fields("SerieDoc") = rsDocumentos.Fields("idSerie") & ""
            
            rsAsientosContables.Fields("NumReferencia") = rsDocumentos.Fields("iddocventas") & ""
            
            If rsAsientosContables!idTipoDH = "H" Then
                sHaberSoles = sHaberSoles + rsAsientosContables.Fields("TotalImporteS")
                sHaberDolares = sHaberDolares + rsAsientosContables.Fields("TotalImporteD")
            Else
                sDebeSoles = sDebeSoles + rsAsientosContables.Fields("TotalImporteS")
                sDebeDolares = sDebeDolares + rsAsientosContables.Fields("TotalImporteD")
            End If
            
            '--- Agregado el tipo de contabilidad a trabajar tributaria o contable
            rsAsientosContables.Fields("TipoContable") = right(CmbOpciones.Text, 2)
            rsAsientosContables.Fields("CtaAuxiliar") = IIf(left("" & rsAsientosContables.Fields("IdCtaContable"), 3) = "122", rsDocumentos.Fields("RUCCliente") & "", "")
            rsAsientosContables.Update
            
        End If
        rsDetalle.MoveNext
    Loop
    
    '------------------------------------------------------------------------------------------------------------
    If Format((sDebeSoles - sHaberSoles), "#0.000") <> 0 Then
        diferenciaSoles = Format(sDebeSoles - sHaberSoles, "#0.00")
        If diferenciaSoles < 0 Then
            If rsDocumentos.Fields("idDocumento") & "" = "07" Or rsDocumentos.Fields("idDocumento") & "" = "89" Then
                diferenciaSoles = diferenciaSoles * -1
            End If
            rsAsientosContables.Fields("TotalImporteS") = Format(rsAsientosContables.Fields("TotalImporteS") + diferenciaSoles, "#0.00")
            rsAsientosContables.Update
        ElseIf diferenciaSoles > 0 Then
            If rsDocumentos.Fields("idDocumento") & "" = "07" Or rsDocumentos.Fields("idDocumento") & "" = "89" Then
                diferenciaSoles = diferenciaSoles * -1
            End If
            rsAsientosContables.Fields("TotalImporteS") = Format(rsAsientosContables.Fields("TotalImporteS") + diferenciaSoles, "#0.00")
            rsAsientosContables.Update
        End If
    End If

    If Format((sDebeDolares - sHaberDolares), "#0.000") <> 0 Then
        diferenciaDolares = Format((sDebeDolares - sHaberDolares), "#0.00")
        If diferenciaDolares < 0 Then
            rsAsientosContables.Fields("TotalImporteD") = Format(rsAsientosContables.Fields("TotalImporteD") + diferenciaDolares, "#0.00")
            rsAsientosContables.Update
        ElseIf diferenciaDolares > 0 Then
            rsAsientosContables.Fields("TotalImporteD") = Format(rsAsientosContables.Fields("TotalImporteD") + diferenciaDolares, "#0.00")
            rsAsientosContables.Update
        End If
    End If

End Sub

Private Sub DetalleRally()
Dim wfecha      As Variant
Dim wdetalle    As String
Dim cquery      As String
Dim rsif5pla    As New ADODB.Recordset
Dim cnomtip     As String
Dim NVALVTA     As Double
Dim NVALVTAD    As Double
Dim Elemen As Integer
   
    rsDetalle.Filter = ""
    rsDetalle.Filter = adFilterNone
    rsDetalle.MoveFirst
    rsDetalle.Sort = "iddocumento,idserie,iddocventas,idproducto"
    
    rsDetalle.Filter = ""
    rsDetalle.Filter = adFilterNone
    rsDetalle.MoveFirst
    rsDetalle.Filter = " iddocumento = '" & rsDocumentos.Fields("iddocumento") & "' and idserie = '" & rsDocumentos.Fields("idSerie") & "' and iddocventas = '" & rsDocumentos.Fields("iddocventas") & "' "
    
    Elemen = 2
    Do While Not rsDetalle.EOF
        If rsDetalle!TotalVVNeto <> 0 Then
            Elemen = Elemen + 1
            
            rsAsientosContables.AddNew
            NItem = NItem + 1
            rsAsientosContables.Fields("ITEM") = Format(NItem, "0000")
            rsAsientosContables!idCosto = Trim("" & rsDocumentos!IdCentroCosto)
            rsAsientosContables.Fields("Glosa") = Trim("" & rsDocumentos.Fields("GlsCliente")) & " " & Trim("" & rsDocumentos.Fields("idDocumento")) & "/" & Trim("" & rsDocumentos.Fields("idSerie")) & "/" & Trim("" & rsDocumentos.Fields("idDocventas"))
            rsAsientosContables!idPeriodo = strAno & strMes
            rsAsientosContables!idComprobante = glsOrigen_Contable & Format(ncorrel, "0000000")
            rsAsientosContables!ValItem = Format(Elemen, "0000")
            
            If "" & rsDetalle.Fields("ctaContable") <> "" Then
                If Len(Trim("" & traerCampo("empresasrelacionadas", "idpersona", "idpersona", Trim("" & rsDocumentos.Fields("idPerCliente")), True))) = 0 Then
                    rsAsientosContables!idCtaContable = Trim("" & rsDetalle.Fields("ctaContable"))
                Else
                    rsAsientosContables!idCtaContable = Trim("" & rsDetalle.Fields("CtaContable_Relacionada"))
                End If
            Else
                rsAsientosContables!idCtaContable = cta70Detalle
            End If
            
            NVALVTA = 0#: NVALVTAD = 0#: ntotbase = 0#
            NVALVTAD = IIf(Val(Format(rsDetalle.Fields("TotalVVNeto") & "", "0.00")) < 0#, Val(Format(rsDetalle.Fields("TotalVVNeto") & "", "0.00")) * -1, Val(Format(rsDetalle.Fields("TotalVVNeto") & "", "0.00")))
            
            If "" & rsDocumentos.Fields("idMoneda") = "USD" Then
            
                NVALVTA = Format(Format((Val("" & rsDetalle.Fields("TotalPvNeto")) * Val("" & rsDocumentos.Fields("TipoCambio"))), "0.00") / 1.18, "0.00")
                
            Else
                
                NVALVTA = Format(Format((Val("" & rsDetalle.Fields("TotalPvNeto"))), "0.00") / 1.18, "0.00")
                
            End If
                
            If "" & rsDocumentos.Fields("idMoneda") = "PEN" Then
                
                rsAsientosContables!TotalImporteS = NVALVTA
                If Val("" & rsDocumentos.Fields("TipoCambio")) <> 0# Then
                    rsAsientosContables!TotalImporteD = Format(NVALVTAD / Val("" & rsDocumentos.Fields("TipoCambio")), "#0.00")
                End If
                
            Else
            
                rsAsientosContables!TotalImporteS = NVALVTA 'Format(NVALVTA * Val("" & rsDocumentos.Fields("TipoCambio")), "#0.00")
                rsAsientosContables!TotalImporteD = NVALVTAD
                
            End If
            
            cnomtip = ""
            If docofi = "S" Then
                Select Case rsDocumentos.Fields("idDocumento")
                    Case "01": cnomtip = "Fac"
                    Case "03": cnomtip = "Bol"
                    Case "07": cnomtip = "Cre"
                    Case "08": cnomtip = "Deb"
                    Case "12": cnomtip = "T/C"
                    Case "90": cnomtip = "Npd"
                End Select
            Else
                Select Case rsDocumentos.Fields("idDocumento")
                    Case "01": cnomtip = "Fac"
                    Case "03": cnomtip = "Bol"
                    Case "07": cnomtip = "Cre"
                    Case "08": cnomtip = "Deb"
                    Case "12": cnomtip = "T/C"
                End Select
            End If
            
            wdetalle = cnomtip & rsDocumentos.Fields("idSerie") & "/" & rsDocumentos.Fields("iddocventas") & Space(2) & left(rsDocumentos.Fields("glsCliente"), 50)
            
            If rsDocumentos.Fields("iddocumento") & "" = "07" Or rsDocumentos.Fields("iddocumento") & "" = "89" Then
                rsAsientosContables!idTipoDH = "D"
                If rsDetalle!TotalVVNeto > 0 Then
                
                    ntotdeb = ntotdeb + Format(NVALVTA, "#0.00")
                    
                    If "" & rsDocumentos!idMoneda = "PEN" Then
                        If Val("" & rsDocumentos!TipoCambio) > 0 Then
                            ntotdeb = ntotdeb + Format(NVALVTA / Val("" & rsDocumentos!TipoCambio), "#0.00")
                        End If
                    Else
                        ntotdeb = ntotdeb + Format(NVALVTA * Val("" & rsDocumentos!TipoCambio), "#0.00")
                    End If
                Else
                    If "" & rsDocumentos!idMoneda = "PEN" Then
                        If Val("" & rsDocumentos!TipoCambio) > 0 Then
                            ntothab = ntothab + Format(NVALVTA / Val("" & rsDocumentos!TipoCambio), "#0.00")
                        End If
                    Else
                        ntothab = ntothab + Format(NVALVTA * Val("" & rsDocumentos!TipoCambio), "#0.00")
                    End If
                End If
            Else
                rsAsientosContables!idTipoDH = "H"
                If rsDetalle!TotalVVNeto > 0 Then
                    If "" & rsDocumentos!idMoneda = "PEN" Then
                        If Val("" & rsDocumentos!TipoCambio) > 0 Then
                            ntothab = ntothab + Format(NVALVTA / Val("" & rsDocumentos!TipoCambio), "#0.00")
                        End If
                    Else
                        ntothab = ntothab + Format(NVALVTA * Val("" & rsDocumentos!TipoCambio), "#0.00")
                    End If
                Else
                    rsAsientosContables!idTipoDH = "D"
                    If "" & rsDocumentos!idMoneda = "PEN" Then
                        If Val("" & rsDocumentos!TipoCambio) > 0 Then
                            ntotdeb = ntotdeb + Format(NVALVTA / Val("" & rsDocumentos!TipoCambio), "#0.00")
                        End If
                    Else
                        ntotdeb = ntotdeb + Format(NVALVTA * Val("" & rsDocumentos!TipoCambio), "#0.00")
                    End If
                End If
            End If
            rsAsientosContables!ValorTipoCambio = Val("" & rsDocumentos!TipoCambio)
            rsAsientosContables!idMoneda = rsDocumentos.Fields("idMoneda") & ""
            rsAsientosContables!IDOrigen = glsOrigen_Contable
            rsAsientosContables!FecCompro = rsDocumentos.Fields("FecEmision") & ""
            rsAsientosContables!GlsDetalle = wdetalle
            rsAsientosContables.Fields("idTipoDoc") = cnomtip
            rsAsientosContables!NumCheque = rsDocumentos.Fields("iddocventas") & ""
            rsAsientosContables.Fields("SerieDoc") = rsDocumentos.Fields("idSerie") & ""
            
            If rsAsientosContables!idTipoDH = "H" Then
                sHaberSoles = sHaberSoles + rsAsientosContables.Fields("TotalImporteS")
                sHaberDolares = sHaberDolares + rsAsientosContables.Fields("TotalImporteD")
            Else
                sDebeSoles = sDebeSoles + rsAsientosContables.Fields("TotalImporteS")
                sDebeDolares = sDebeDolares + rsAsientosContables.Fields("TotalImporteD")
            End If
            
            '--- Agregado el tipo de contabilidad a trabajar tributaria o contable
            rsAsientosContables.Fields("TipoContable") = right(CmbOpciones.Text, 2)
            rsAsientosContables.Update
            
        End If
        rsDetalle.MoveNext
    Loop
    
    '------------------------------------------------------------------------------------------------------------
    If Format((sDebeSoles - sHaberSoles), "#0.000") <> 0 Then
        diferenciaSoles = Format(sDebeSoles - sHaberSoles, "#0.00")
        If diferenciaSoles < 0 Then
            If rsDocumentos.Fields("idDocumento") & "" = "07" Or rsDocumentos.Fields("idDocumento") & "" = "89" Then
                diferenciaSoles = diferenciaSoles * -1
            End If
            rsAsientosContables.Fields("TotalImporteS") = Format(rsAsientosContables.Fields("TotalImporteS") + diferenciaSoles, "#0.00")
            rsAsientosContables.Update
        ElseIf diferenciaSoles > 0 Then
            If rsDocumentos.Fields("idDocumento") & "" = "07" Or rsDocumentos.Fields("idDocumento") & "" = "89" Then
                diferenciaSoles = diferenciaSoles * -1
            End If
            rsAsientosContables.Fields("TotalImporteS") = Format(rsAsientosContables.Fields("TotalImporteS") + diferenciaSoles, "#0.00")
            rsAsientosContables.Update
        End If
    End If

    If Format((sDebeDolares - sHaberDolares), "#0.000") <> 0 Then
        diferenciaDolares = Format((sDebeDolares - sHaberDolares), "#0.00")
        If diferenciaDolares < 0 Then
            rsAsientosContables.Fields("TotalImporteD") = Format(rsAsientosContables.Fields("TotalImporteD") + diferenciaDolares, "#0.00")
            rsAsientosContables.Update
        ElseIf diferenciaDolares > 0 Then
            rsAsientosContables.Fields("TotalImporteD") = Format(rsAsientosContables.Fields("TotalImporteD") + diferenciaDolares, "#0.00")
            rsAsientosContables.Update
        End If
    End If

End Sub

Private Sub AsientosContablesDetallado_2()
On Error GoTo ERR
Dim rscontrol               As New ADODB.Recordset
Dim rsd                     As New ADODB.Recordset
Dim rs                      As New ADODB.Recordset
Dim csql                    As String
Dim strSQL                  As String
Dim CodProd                 As String
Dim CodProdAnt              As String
Dim CFecha                  As String
Dim cconex_dbbancos         As String
Dim cselect                 As String
Dim CCua                    As Double
Dim dblsaldo                As Double
Dim dblValSaldo             As Double
Dim Cadena_Oficial          As String
Dim Cadena_Transferido      As String
Dim Cadmysql                As String
Dim OrigenCon               As String
Dim CIVAP_Cuenta            As String

    Me.MousePointer = 11
    
    strAno = cbxAno.Text
    strMes = Format(cbxMes.ListIndex + 1, "00")
    CFecha = CAPTURA_FECHA_FIN(strMes, strAno)
    
    cadenadoc = IIf(docofi = "S", " and dv.iddocumento in('01','03','07','08','12','90')", " and dv.iddocumento in('01','03','07','08','12')")
    Cadena_Oficial = ""
    If right(CmbOpciones.Text, 2) = "01" Then
        Cadena_Oficial = " and d.indoficial = '1' "
    Else
        Cadena_Oficial = " and d.indoficial in('1','0') "
    End If

    Cadena_Transferido = ""
    If right(CmbOpciones.Text, 2) = "01" Then
        Cadena_Transferido = " and dv.indTrasladoConta <> 'S' "
    Else
        Cadena_Transferido = " and dv.indTrasladoContaFin <> 'S' "
    End If

    cta40IGV = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_40", True)
    cta70Detalle = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_70", True)
    OrigenCon = traerCampo("parametros", "valParametro", "glsParametro", "ORIGEN_CONTABLE", True)
    CIVAP_Cuenta = traerCampo("parametros", "valParametro", "glsParametro", "IVAP_CUENTA", True)
    
    ncorrel = 0
    strSQL = "SELECT dv.idempresa,dv.idsucursal,dv.iddocumento, dv.idSerie, dv.iddocventas,dv.Fecemision,dv.idcomprobante,dv.idperiodo " & _
            "FROM docventas dv " & _
            "inner join Documentos d on dv.iddocumento = d.iddocumento " & _
            "where dv.idEmpresa = '" & glsEmpresa & "' " & _
            Cadena_Oficial & _
            Cadena_Transferido & _
            "and month(dv.fecEmision) = " & Val(strMes) & " and year(dv.fecEmision) = " & Val(strAno) & " " & _
            "order by dv.iddocumento, dv.idSerie, dv.iddocventas "
    
    If rsDocumentos.State = 1 Then rsDocumentos.Close
    rsDocumentos.Open strSQL, Cn, adOpenKeyset, adLockOptimistic
            
    If Not rsDocumentos.EOF Then
        rsDocumentos.MoveFirst
        Do While Not rsDocumentos.EOF
                
            Cadmysql = "Call Spu_AsientoContableVentas('" & rsDocumentos.Fields("idempresa") & "','" & rsDocumentos.Fields("idsucursal") & "', " & _
                       "'" & rsDocumentos.Fields("iddocumento") & "','" & rsDocumentos.Fields("idSerie") & "','" & rsDocumentos.Fields("iddocventas") & "', " & _
                       "'" & Format(rsDocumentos.Fields("Fecemision"), "yyyy-mm-dd") & "','" & OrigenCon & "','" & cta40IGV & "','" & cta70Detalle & "'," & _
                       "'" & rsDocumentos.Fields("idcomprobante") & "','" & rsDocumentos.Fields("idperiodo") & "','','" & CIVAP_Cuenta & "','" & glsUser & "','" & fpComputerName & "','" & fpUsuarioActual & "')"
                
            CnConta.Execute (Cadmysql)
            
            rsDocumentos.MoveNext
        Loop
        MsgBox "Fin del proceso.", vbInformation, App.Title
    Else
        MsgBox "No hay registros para transferir", vbInformation, App.Title
    End If
    Me.MousePointer = 1
        
    Exit Sub
    
ERR:
    If rsTemp.State = 1 Then rsTemp.Close: Set rsTemp = Nothing
    Me.MousePointer = 1
    MsgBox ERR.Description, vbInformation, App.Title
    Exit Sub
End Sub

Private Sub AsientosContablesDetalladoRally()
On Error GoTo ERR
Dim rscontrol               As New ADODB.Recordset
Dim rsd                     As New ADODB.Recordset
Dim rs                      As New ADODB.Recordset
Dim csql                    As String
Dim strSQL                  As String
Dim CodProd                 As String
Dim CodProdAnt              As String
Dim CFecha                  As String
Dim cconex_dbbancos         As String
Dim cselect                 As String
Dim CCua                    As Double
Dim dblsaldo                As Double
Dim dblValSaldo             As Double
Dim Cadena_Oficial          As String
Dim Cadena_Transferido      As String
Dim Cadena_Oficial2         As String
    
    Me.MousePointer = 11
    strAno = cbxAno.Text
    strMes = Format(cbxMes.ListIndex + 1, "00")
    CFecha = CAPTURA_FECHA_FIN(strMes, strAno)
    
    Cadena_Oficial = ""
    If right(CmbOpciones.Text, 2) = "01" Then
        Cadena_Oficial = " and d.indoficial = '1' "
    Else
        Cadena_Oficial = " and d.indoficial in('1','0') "
    End If
    
    Cadena_Oficial2 = ""
    If right(CmbOpciones.Text, 2) = "01" Then
        Cadena_Oficial2 = " and dc.indoficial = '1' "
    Else
        Cadena_Oficial2 = " and dc.indoficial in('1','0') "
    End If
    
    Cadena_Transferido = ""
    If right(CmbOpciones.Text, 2) = "01" Then
        Cadena_Transferido = " and dv.indTrasladoConta <> 'S' "
    Else
        Cadena_Transferido = " and dv.indTrasladoContaFin <> 'S' "
    End If
            
    strSQL = "SELECT dv.iddocumento, dv.idSerie, dv.iddocventas, dv.idPerCliente, dv.glsCLiente, dv.RUCCliente, dv.idAlmacen, " & _
            "dv.FecEmision, dv.estDocVentas, dv.idMoneda, dv.idEmpresa, dv.idSucursal, " & _
            "dv.TotalValorVenta, dv.TotalIGVVenta, dv.TotalPrecioVenta, dv.totalbaseimponible, " & _
            "if(dv.iddocumento <> '07', t.tcVenta,ifnull(tc.tcNC,T.tcventa)) as TipoCambio, dv.idcentrocosto " & _
            "FROM docventas dv inner join tiposdecambio t " & _
            "on (Day(dv.FecEmision) = Day(t.fecha) and Year(dv.FecEmision) = Year(t.fecha) and Month(dv.FecEmision) = Month(t.fecha)) " & _
            "left join (select x.tcVenta as tcNC, r.tipoDocOrigen, r.serieDocOrigen, r.numDocOrigen " & _
                    "from docventas dt inner join docreferencia r " & _
                    "on dt.IdEmpresa = r.IdEmpresa And dt.iddocumento = r.tipoDocReferencia and dt.idSerie = r.serieDocReferencia and dt.idDocVentas = r.numDocReferencia " & _
                    "inner join tiposdecambio x on (Day(dt.FecEmision) = Day(x.fecha) and Year(dt.FecEmision) = Year(x.fecha) and Month(dt.FecEmision) = Month(x.fecha)) " & _
                    "where dt.IdEmpresa = '" & glsEmpresa & "' And r.tipoDocOrigen = '07' " & _
                    "group by r.tipodocorigen, r.numdocorigen) tc " & _
            "on tc.tipoDocOrigen = dv.idDocumento and tc.serieDocOrigen = dv.idSerie and tc.numDocOrigen = dv.idDocVentas " & _
            "inner join Documentos d on dv.iddocumento = d.iddocumento " & _
            "where dv.idEmpresa = '" & glsEmpresa & "' " & _
            Cadena_Oficial & _
            Cadena_Transferido & _
            "and month(dv.fecEmision) = " & Val(strMes) & " and year(dv.fecEmision) = " & Val(strAno) & " " & _
            "order by dv.iddocumento, dv.idSerie, dv.iddocventas "
    
    If rsDocumentos.State = 1 Then rsDocumentos.Close
    rsDocumentos.Open strSQL, Cn, adOpenKeyset, adLockOptimistic
    
    cta12soles = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_SOLES", True)
    cta12dolares = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_DOLARES", True)
    
    cta12soles_relacionada = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_SOLES_RELACIONADA", True)
    cta12dolares_relacionada = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_DOLARES_RELACIONADA", True)
        
    cta40IGV = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_40", True)
    cta70Detalle = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_70", True)
    
    ncorrel = 0
    If rsDocumentos.RecordCount <> 0 Then
        cquery = "select p.CtaContable_Relacionada,p.CtaContable, d.idProducto, Left(d.glsProducto,150) GlsProducto, d.Cantidad, d.TotalVVNeto, d.TotalIGVNeto, d.TotalPVNeto, " & _
                 "d.iddocumento,d.iddocventas,d.idempresa,d.idserie " & _
                 "from docventas dv inner join docventasdet d " & _
                 "on dv.idempresa = d.idempresa " & _
                 "and dv.iddocventas = d.iddocventas " & _
                 "and dv.idserie = d.idserie " & _
                 "and dv.iddocumento = d.iddocumento " & _
                 "inner join Productos p " & _
                 "on d.idProducto = p.idProducto " & _
                 "and d.idempresa = p.idempresa " & _
                 "inner join Documentos dc on dv.iddocumento = dc.iddocumento " & _
                "where dv.idEmpresa = '" & glsEmpresa & "' " & _
                Cadena_Oficial2 & _
                Cadena_Transferido & _
                "and month(dv.fecEmision) = " & Val(strMes) & " and year(dv.fecEmision) = " & Val(strAno) & " " & _
                "order by dv.iddocumento,dv.idserie,dv.iddocventas,d.idProducto "
        
        If rsDetalle.State = adStateOpen Then rsDetalle.Close
        rsDetalle.Open cquery, Cn, adOpenKeyset, adLockOptimistic
        If rsAsientosContables.State = 1 Then rsAsientosContables.Close
    
        rsAsientosContables.Fields.Append "Item", adInteger, 2, adFldRowID
        rsAsientosContables.Fields.Append "idComprobante", adVarChar, 9, adFldRowID
        rsAsientosContables.Fields.Append "idPeriodo", adDouble, 11, adFldIsNullable
        rsAsientosContables.Fields.Append "ValItem", adVarChar, 4, adFldRowID
        rsAsientosContables.Fields.Append "IDOrigen", adVarChar, 2, adFldIsNullable
        rsAsientosContables.Fields.Append "FecCompro", adVarChar, 30, adFldRowID
        rsAsientosContables.Fields.Append "GlsDetalle", adVarChar, 250, adFldIsNullable
        rsAsientosContables.Fields.Append "idCtaContable", adVarChar, 150, adFldRowID
        rsAsientosContables.Fields.Append "idGasto", adVarChar, 4, adFldIsNullable
        rsAsientosContables.Fields.Append "NumCheque", adDouble, adFldRowID
        rsAsientosContables.Fields.Append "NumReferencia", adVarChar, 45, adFldIsNullable
        rsAsientosContables.Fields.Append "TotalImporteS", adDouble, adFldRowID
        rsAsientosContables.Fields.Append "TotalImporteD", adDouble, adFldIsNullable
        rsAsientosContables.Fields.Append "idMoneda", adChar, 3, adFldRowID
        rsAsientosContables.Fields.Append "ValorTipoCambio", adDouble, adFldIsNullable
        rsAsientosContables.Fields.Append "idTipoDoc", adVarChar, 3, adFldRowID
        rsAsientosContables.Fields.Append "idTipoDH", adChar, 1, adFldIsNullable
        rsAsientosContables.Fields.Append "idCosto", adVarChar, 8, adFldRowID
        rsAsientosContables.Fields.Append "Destino", adVarChar, 1, adFldIsNullable
        rsAsientosContables.Fields.Append "CtaAuxiliar", adVarChar, 11, adFldRowID
        rsAsientosContables.Fields.Append "F3AUTOMATICO", adDouble, 11, adFldRowID
        rsAsientosContables.Fields.Append "F3ORIGAUTO", adDouble, 11, adFldRowID
        rsAsientosContables.Fields.Append "F4FECVENC", adDouble, 11, adFldRowID
        rsAsientosContables.Fields.Append "F3RUC", adDouble, 11, adFldRowID
        rsAsientosContables.Fields.Append "SerieDoc", adChar, 4, adFldRowID
        rsAsientosContables.Fields.Append "f3nummov", adVarChar, 11, adFldRowID
        rsAsientosContables.Fields.Append "f2tipdoc", adVarChar, 2, adFldRowID
        rsAsientosContables.Fields.Append "obra", adVarChar, 85, adFldRowID
        rsAsientosContables.Fields.Append "F3FECHADOCUM", adVarChar, 30, adFldRowID
        rsAsientosContables.Fields.Append "F3FECHACOMP", adVarChar, 30, adFldRowID
        rsAsientosContables.Fields.Append "F3NUMEROCOMP", adVarChar, 250, adFldRowID
        rsAsientosContables.Fields.Append "indAfecto", adVarChar, 1, adFldRowID
        rsAsientosContables.Fields.Append "ValImporte", adDouble, adFldRowID
        rsAsientosContables.Fields.Append "IdCtaCorriente", adInteger, 2, adFldRowID
        rsAsientosContables.Fields.Append "IdMovBancosCab", adInteger, 2, adFldRowID
        rsAsientosContables.Fields.Append "TipoContable", adVarChar, 2, adFldRowID
        rsAsientosContables.Fields.Append "glosa", adVarChar, 255, adFldRowID
        rsAsientosContables.Open
        
        rsDocumentos.MoveFirst
        Do While Not rsDocumentos.EOF
            sHaberSoles = 0#: sDebeSoles = 0#
            sHaberDolares = 0#: sDebeDolares = 0#
            ncorrel = ncorrel + 1
            TOTALRally
            If rsDocumentos.Fields("estDocVentas") & "" <> "ANU" Then
                IgvRally
                Elemen = 3
                'DESCUENTOS
                DetalleRally
            End If
            rsDocumentos.MoveNext
        Loop
        MsgBox "Fin del proceso.", vbInformation, App.Title
    Else
        MsgBox "No hay registros para transferir", vbInformation, App.Title
    End If
    Me.MousePointer = 1
        
    Exit Sub
    
ERR:
    If rsTemp.State = 1 Then rsTemp.Close: Set rsTemp = Nothing
    Me.MousePointer = 1
    MsgBox ERR.Description, vbInformation, App.Title
    Exit Sub
End Sub

