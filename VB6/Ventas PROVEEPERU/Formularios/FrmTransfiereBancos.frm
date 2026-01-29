VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmTransfiereBancos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencia a Bancos"
   ClientHeight    =   4335
   ClientLeft      =   3585
   ClientTop       =   3750
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnmodificar 
      Caption         =   "Modificar"
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
      Left            =   2835
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3735
      Width           =   1320
   End
   Begin VB.CommandButton btnsalir 
      Caption         =   "Salir"
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
      Left            =   4275
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3735
      Width           =   1320
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   " Moneda a Tranferir "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   850
      Left            =   90
      TabIndex        =   19
      Top             =   2700
      Width           =   6945
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dólares"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         TabIndex        =   5
         Top             =   405
         Width           =   1005
      End
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   4
         Top             =   405
         Width           =   1050
      End
      Begin VB.OptionButton OptAmbos 
         Caption         =   "Ambos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   3
         Top             =   405
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   850
      Index           =   9
      Left            =   90
      TabIndex        =   15
      Top             =   990
      Width           =   6945
      Begin VB.CommandButton cmbAyudaUsuario 
         Height          =   315
         Left            =   6330
         Picture         =   "FrmTransfiereBancos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   320
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Usuario 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Tag             =   "TidMoneda"
         Top             =   315
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "FrmTransfiereBancos.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Usuario 
         Height          =   315
         Left            =   2160
         TabIndex        =   17
         Top             =   315
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "FrmTransfiereBancos.frx":03A6
         Vacio           =   -1  'True
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   285
         TabIndex        =   18
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   850
      Left            =   90
      TabIndex        =   13
      Top             =   1800
      Width           =   6945
      Begin MSComCtl2.DTPicker dtpDiaTrans 
         Height          =   300
         Left            =   3090
         TabIndex        =   2
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   107479041
         CurrentDate     =   39888
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   2475
         TabIndex        =   14
         Top             =   405
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   850
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   6945
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6330
         Picture         =   "FrmTransfiereBancos.frx":03C2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "FrmTransfiereBancos.frx":074C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "FrmTransfiereBancos.frx":0768
         Vacio           =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   285
         TabIndex        =   12
         Top             =   405
         Width           =   645
      End
   End
   Begin VB.CommandButton btnaceptar 
      Caption         =   "&Aceptar"
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
      Top             =   3735
      Width           =   1320
   End
End
Attribute VB_Name = "FrmTransfiereBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cnntabla    As New ADODB.Connection
Dim cnnbase     As New ADODB.Connection

Private Sub Btnaceptar_Click()
On Error GoTo Err
Dim StrMsgError As String

    If Trim(txtCod_Sucursal.Text = "") Then
        StrMsgError = "Ingrese  Sucursal": txtCod_Sucursal.SetFocus
        GoTo Err
    End If
    
    If Trim(txtCod_Usuario.Text = "") Then
        StrMsgError = "Ingrese  Usuario": txtCod_Usuario.SetFocus
        GoTo Err
    End If
    
    If MsgBox("Està seguro(a) de Realizar la TransFerencia ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        transferir StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub btnmodificar_Click()

    limpia_caja
    MsgBox ("Se Modificaron los Datos"), vbInformation, App.Title

End Sub

Private Sub limpia_caja()
Dim csql    As String
Dim csq2    As String
Dim usuario As String
Dim fecha   As String
 
    If OptAmbos.Value = True Or OptSoles.Value = True Then
        csql = "Update movcajas Set indTransferido='N', CodCta_Sol=0, NumMov_Sol=0 " & _
               "Where idSucursal = '" & Trim(txtCod_Sucursal.Text) & "' and idusuario = '" & Trim(txtCod_Usuario.Text) & "' " & _
               "And feccaja = '" & Format(dtpDiaTrans.Value, "yyyy-mm-dd") & "' And idEmpresa ='" & glsEmpresa & "' "
        Cn.Execute (csql)
    End If
    
    If OptAmbos.Value = True Or OptDolares.Value = True Then
        csq2 = "Update movcajas Set CodCta_Dol=0, NumMov_Dol=0 " & _
               "Where idSucursal = '" & Trim(txtCod_Sucursal.Text) & "' and idusuario = '" & Trim(txtCod_Usuario.Text) & "' " & _
               "And feccaja = '" & Format(dtpDiaTrans.Value, "yyyy-mm-dd") & "' And idEmpresa ='" & glsEmpresa & "'  "
        Cn.Execute (csq2)
    End If

End Sub

Private Sub btnSalir_Click()
    
    Unload Me

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub transferir(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsmovcajadetSol     As New ADODB.Recordset
Dim rsmovcajadetDol     As New ADODB.Recordset
Dim rs_nummov           As New ADODB.Recordset
Dim idmovcaja           As String
Dim csql_insert         As String
Dim csql_insertDet      As String
Dim strcodcta_sol       As String
Dim strcodcta_dol       As String
Dim strtipomov          As String
Dim strRUCCliente       As String
Dim cserie              As String
Dim cnumdoc             As String
Dim ctipodoc            As String
Dim strnomcli           As String
Dim strnomdoc           As String
Dim nvaltotalFac        As Double
Dim ntipcam             As Double
Dim nvaltotalFacAcum    As Double
Dim nmov                As Integer
Dim nelemen             As Integer
Dim strpagoSol          As String
Dim cdirectorio         As String
Dim strestadocaja       As String
Dim CodCta_Sol          As Integer
Dim NumMov_Sol          As Integer
Dim CodCta_Dol          As Integer
Dim NumMov_Dol          As Integer
Dim idcaja              As String
Dim idGastosIngresos    As String
Dim indTrans            As Boolean
Dim strParamFPT         As String
Dim i                   As Integer
Dim strCodCtaCorr       As String
Dim rst                 As New ADODB.Recordset
Dim strParamFPT1        As String
Dim strParamFPT2        As String
Dim X                   As Integer
        
    X = 0
    strParamFPT1 = traerCampo("Parametros", "ValParametro", "GlsParametro", "CODIGO_FORMA_PAGO_TARJETA_1", True)
    strParamFPT2 = traerCampo("Parametros", "ValParametro", "GlsParametro", "CODIGO_FORMA_PAGO_TARJETA_2", True)
    
    indTrans = False
    
    Cn.BeginTrans
    indTrans = True
    
    nelemen = 0
    CMes = Format(dtpDiaTrans.Value, "mm")
    idmovcaja = traerCampo("movcajas", "idmovcaja", "FecCaja", Format(dtpDiaTrans.Value, "yyyy-mm-dd"), True, "  idsucursal = '" & txtCod_Sucursal.Text & "' and idUsuario = '" & txtCod_Usuario.Text & "'")
    idcaja = traerCampo("movcajas", "idcaja", "FecCaja", Format(dtpDiaTrans.Value, "yyyy-mm-dd"), True, "  idsucursal = '" & txtCod_Sucursal.Text & "' and idUsuario = '" & txtCod_Usuario.Text & "'")
    strestadocaja = traerCampo("movcajas", "indTransferido", "FecCaja", Format(dtpDiaTrans.Value, "yyyy-mm-dd"), True, "  idsucursal = '" & txtCod_Sucursal.Text & "' ")
  
    strcodcta_sol = traerCampo("Cajas", "ctacajasoles", "idcaja", idcaja, True)
    strcodcta_dol = traerCampo("Cajas", "ctacajadolar", "idcaja", idcaja, True)
    strtipomov = traerCampo("parametros", "valparametro", "glsparametro", "CODIGO_TIPMOV_INGRESO", True)
    
    '--- Cuenta Soles x Usuario
    CodCta_Sol = Val(traerCampo("movcajas", "CodCta_Sol", "idmovcaja", idmovcaja, True, "  idsucursal = '" & txtCod_Sucursal.Text & "' and idUsuario = '" & txtCod_Usuario.Text & "'"))
    NumMov_Sol = Val(traerCampo("movcajas", "NumMov_Sol", "idmovcaja", idmovcaja, True, "  idsucursal = '" & txtCod_Sucursal.Text & "' and idUsuario = '" & txtCod_Usuario.Text & "'"))
    
    '--- Cuenta Dolares x Usuario
    CodCta_Dol = Val(traerCampo("movcajas", "CodCta_Dol", "idmovcaja", idmovcaja, True, "  idsucursal = '" & txtCod_Sucursal.Text & "' and idUsuario = '" & txtCod_Usuario.Text & "'"))
    NumMov_Dol = Val(traerCampo("movcajas", "NumMov_Dol", "idmovcaja", idmovcaja, True, "  idsucursal = '" & txtCod_Sucursal.Text & "' and idUsuario = '" & txtCod_Usuario.Text & "'"))
    
    If Trim(idmovcaja) = "" Then
        StrMsgError = "No hay caja disponible de este día": GoTo Err
    Else
        If (OptSoles.Value = True And CodCta_Sol <> 0 And NumMov_Sol <> 0) Then
            StrMsgError = "La caja de este día ya fue transferido": GoTo Err
        End If
        
        If (OptDolares.Value = True And CodCta_Dol <> 0 And NumMov_Dol <> 0) Then
            StrMsgError = "La caja de este día ya fue transferido": GoTo Err
        End If
        
        If OptAmbos.Value = True And CodCta_Sol <> 0 And NumMov_Sol <> 0 And CodCta_Dol <> 0 And NumMov_Dol <> 0 Then
            StrMsgError = "La caja para soles o dolares ya fue transferida": GoTo Err
        End If
        
    End If
    
    ntipcam = traerCampo("tiposdecambio", "tcVenta", "DATE_FORMAT(fecha,'%d/%m/%Y')", Format(dtpDiaTrans.Value, "dd/mm/yyyy"), False)
         
    If OptAmbos.Value = True Or OptSoles.Value = True Then
         csql = "Select * From movcajasdet m " & _
               "Inner Join Formaspagos f     On f.idFormaPago = m.idFormadePago   And f.idEmpresa = m.idEmpresa " & _
               "Inner Join TipoFormaspago t  On t.idTipoFormaPago = f.idTipoFormaPago " & _
               "Where idMoneda = 'PEN' And idmovcaja = '" & idmovcaja & "' And t.idTipoFormaPago = '06090001' And m.idTipoMovCaja = '99990002' And m.estMovCajaDet = 'ACT' " & _
               "And m.idEmpresa = '" & glsEmpresa & "'  And  m.idSucursal ='" & Trim(txtCod_Sucursal.Text) & " ' Order By idmoneda,idMovCajaDet  "
        
        If rsmovcajadetSol.State = 1 Then rsmovcajadetSol.Close
        rsmovcajadetSol.Open csql, Cn, adOpenDynamic, adLockOptimistic
        
        If Not rsmovcajadetSol.EOF Then
            If rs_nummov.State = 1 Then rs_nummov.Close
             
            csql_mov = "Select idMovBancosCab from MovBancosCab where idEmpresa='" & glsEmpresa & "' and idCtaCorriente=" & strcodcta_sol & " and periodo='" & Format(dtpDiaTrans.Value, "yyyymm") & "' order by idMovBancosCab desc"
            rs_nummov.Open csql_mov, Cn, adOpenDynamic, adLockOptimistic
             
            If Not rs_nummov.EOF Then
                nmov = Val(rs_nummov.Fields("idMovBancosCab")) + 1
            Else
                nmov = 1
            End If
            rs_nummov.Close: Set rs_nummov = Nothing
            
            csql_insert = " INSERT INTO MovBancosCab (idCtaCorriente,idMovBancosCab,idOpeBancaria," & _
                          " Nro_Comp,Detalle,FecEmision,Destino,TotalCab,FechaReg,IdUsuarioReg, " & _
                          " IndPendiente,TipoCambio,Periodo,IdEmpresa,estMovBancos,idTalonario,Observacion1, Observacion2, idPendiente,idCliProv)  " & _
                          " VALUES (" & strcodcta_sol & "," & nmov & ",'" & strtipomov & "', " & _
                          " '" & Format(dtpDiaTrans.Value, "ddmmyy") & "','Cobranza Diaria', " & _
                          " '" & Format(dtpDiaTrans.Value, "yyyy-mm-dd") & "','I',0, " & _
                          " '" & Format(Now, "yyyy-mm-dd") & "', '" & glsUser & "'," & _
                          " '0'," & ntipcam & ",'" & Format(dtpDiaTrans.Value, "yyyymm") & "','" & glsEmpresa & "','ACT','0','','','','')"
            Cn.Execute (csql_insert)
            
            nelemen = 0
            nvaltotalFacAcum = 0
            cserie = ""
            cnumdoc = ""
            ctipodoc = ""
            strnomdoc = ""
            nvaltotalFac = 0
            nvalPago = 0
            nvaltotalVuelto = 0
            idGastosIngresos = ""
            
            strcodgto = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_SOLES", True)
            Do While Not rsmovcajadetSol.EOF
                nelemen = nelemen + 1
                
                cserie = rsmovcajadetSol.Fields("idSerie") & ""
                cnumdoc = rsmovcajadetSol.Fields("iddocventas") & ""
                ctipodoc = rsmovcajadetSol.Fields("iddocumento") & ""
                strnomdoc = traerCampo("documentos", "abreDocumento", "iddocumento", ctipodoc, False)
                nvaltotalFac = Trim(traerCampo("docventas", "TotalPrecioVenta", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & txtCod_Sucursal.Text & "'"))
                nvalPago = Val(Format(rsmovcajadetSol.Fields("valMonto"), "0.00"))
                nvaltotalVuelto = traerCampo("movcajasdet", "ValMonto", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & txtCod_Sucursal.Text & "' and  idTipoMovCaja='99990003' ")
                
                If nvalPago > Val(nvaltotalVuelto) Then
                    nvalPago = nvalPago - Format(Val(nvaltotalVuelto), "0.00")
                End If
                
                nvaltotalFacAcum = nvaltotalFacAcum + nvalPago
                strnomcli = Trim(traerCampo("docventas", "glsCliente", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & txtCod_Sucursal.Text & "'"))
                strRUCCliente = Trim(traerCampo("docventas", "RUCCliente", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & txtCod_Sucursal.Text & "'"))
                
                idGastosIngresos = traerCampo("GastosIngresos", "idGastosIngresos", "idCtaContable", "" & traerCampo("Documentos", "idCtaVtaSoles", "idDocumento", ctipodoc, False), True, "indTipo ='I'")
                
                csql_insertDet = " INSERT INTO MovBancosDet (IdCtaCorriente,IdMovBancosCab,Item," & _
                      " Periodo,idCtaContable,GlsDetalle,NroDocumento,ValMonto,TipoCambio,indDebHab,idTipoDocumento, " & _
                      " IdGastosIngresos,SerieDocumento,Ruc_Dni,IdEmpresa, idCentroCosto,idPag_MovBancos, NumMovCompras, idPag_Dcto, idCtaCorrienteRec, idMovBancosCabRec,FecEmision) " & _
                      " VALUES (" & strcodcta_sol & "," & nmov & "," & nelemen & ", " & _
                      " '" & Format(dtpDiaTrans.Value, "yyyymm") & "','" & strcodgto & "','" & strnomcli & "','" & cnumdoc & "', " & _
                      " " & Val(Format(nvalPago, "0.00")) & "," & ntipcam & ", " & _
                      " " & "'H','" & strnomdoc & "', " & _
                      " '" & idGastosIngresos & "', '" & cserie & "' ,'" & strRUCCliente & "','" & glsEmpresa & "','','','','','','','" & Format(dtpDiaTrans.Value, "yyyy-mm-dd") & "') "
                Cn.Execute (csql_insertDet)
                     
                rsmovcajadetSol.MoveNext
            Loop
            
            Cn.Execute ("update movcajas set indTransferido='S', CODCTA_SOL = " & strcodcta_sol & ", NUMMOV_SOL = " & nmov & "  where idmovcaja= '" & idmovcaja & "'  And idEmpresa ='" & glsEmpresa & "' And idSucursal ='" & Trim(txtCod_Sucursal.Text) & "'")
            Cn.Execute ("UPDATE MovBancosCab SET TotalCab= " & Val(Format(nvaltotalFacAcum, "0.00")) & " , TotalDet= " & Val(Format(nvaltotalFacAcum, "0.00")) & "  Where idEmpresa='" & glsEmpresa & "' and Periodo='" & Format(dtpDiaTrans.Value, "yyyymm") & "' and IdCtaCorriente = " & strcodcta_sol & " AND IdMovBancosCab = " & nmov & "  ")
        
        Else
            X = 1
        End If
        
       'Transfiere Ventas Canceladas con Tarjeta
       '----------------------------------------------------------------------------- -----------------------------------------------------------------------------  -----------------------------------------------------------------------------
       For i = 1 To 2
        strParamFPT = traerCampo("Parametros", "ValParametro", "GlsParametro", "CODIGO_FORMA_PAGO_TARJETA_" & i & "", True)
        csql = "Select f.idTipoFormaPago, f.GlsFormaPago From MovcajasDet m " & _
             "Inner Join Formaspagos f " & _
               "On f.idFormaPago = m.idFormadePago And f.idEmpresa = m.idEmpresa " & _
             "Inner Join TipoFormaspago t " & _
               "On t.idTipoFormaPago = f.idTipoFormaPago " & _
             "Inner Join Docventas d " & _
               "On m.idSerie = d.idSerie  And m.idDocventas = d.idDocventas   And m.idDocumento = d.idDocumento  And m.idEmpresa = d.idEmpresa  And m.idSucursal = d.idSucursal " & _
             "Where m.idMoneda = 'PEN' And m.idmovcaja = '" & idmovcaja & "' And t.idTipoFormaPago = '06090004'  And m.idTipoMovCaja = '99990002'  And m.estMovCajaDet = 'ACT'  And m.idEmpresa = '" & glsEmpresa & "'  And m.idSucursal = '" & glsSucursal & "'  And m.idFormadePago IN('" & strParamFPT & "') " & _
             "Group By m.idFormadePago"
         rst.Open csql, Cn, adOpenDynamic, adLockOptimistic
         If Not rst.EOF Then
         
                 strCodCtaCorr = traerCampo("Cajas", "CtaCajSolTarj" & i & " ", "idcaja", idcaja, True)
                 nmov = Val(traerCampo("MovbancosCab", "idMovBancosCab", "idCtaCorriente", strCodCtaCorr, True, " Periodo = '" & Format(dtpDiaTrans.Value, "yyyymm") & "' order by idMovBancosCab desc ")) + 1
                 
                 csql = "Call Spu_TransfiereVentas('" & glsEmpresa & "','" & glsSucursal & "','" & nmov & "','" & idmovcaja & "','" & strParamFPT & "','" & strCodCtaCorr & "','" & strtipomov & "','" & Format(dtpDiaTrans.Value, "yyyy-mm-dd") & "','" & Trim(txtCod_Usuario.Text) & "','" & ntipcam & "')"
                 Cn.Execute (csql)
          Else
             X = X + 1
          End If
          If rst.State = 1 Then rst.Close: Set rst = Nothing
        Next
       '----------------------------------------------------------------------------- -----------------------------------------------------------------------------  -----------------------------------------------------------------------------
        If X = 3 Then
            StrMsgError = "La caja no tiene movimientos": GoTo Err
        End If
    End If
    
    If OptAmbos.Value = True Or OptDolares.Value = True Then
        csql = "Select * From movcajasdet m " & _
               "Inner Join Formaspagos f  On f.idFormaPago = m.idFormadePago   And f.idEmpresa = m.idEmpresa " & _
               "Inner Join TipoFormaspago t  On t.idTipoFormaPago = f.idTipoFormaPago " & _
               "Where idMoneda = 'USD' And idmovcaja = '" & idmovcaja & "' And t.idTipoFormaPago = '06090001' And m.idTipoMovCaja = '99990002' And m.estMovCajaDet = 'ACT' " & _
               "And m.idEmpresa = '" & glsEmpresa & "'  And  m.idSucursal ='" & Trim(txtCod_Sucursal.Text) & " ' Order By idmoneda,idMovCajaDet  "
      
        If rsmovcajadetDol.State = 1 Then rsmovcajadetDol.Close
        rsmovcajadetDol.Open csql, Cn, adOpenDynamic, adLockOptimistic
        nmov = 0
        If Not rsmovcajadetDol.EOF Then
            
            If rs_nummov.State = 1 Then rs_nummov.Close
          
            csql_mov = "Select idMovBancosCab from MovBancosCab where idEmpresa='" & glsEmpresa & "' and idCtaCorriente=" & strcodcta_dol & " and periodo='" & Format(dtpDiaTrans.Value, "yyyymm") & "' order by idMovBancosCab desc"
            rs_nummov.Open csql_mov, Cn, adOpenDynamic, adLockOptimistic
            If Not rs_nummov.EOF Then
                nmov = Val(rs_nummov.Fields("idMovBancosCab")) + 1
            Else
                nmov = 1
            End If
            rs_nummov.Close: Set rs_nummov = Nothing
 
            csql_insert = " INSERT INTO MovBancosCab (idCtaCorriente,idMovBancosCab,idOpeBancaria," & _
                          " Nro_Comp,Detalle,FecEmision,Destino,TotalCab,FechaReg,IdUsuarioReg, " & _
                          " IndPendiente,TipoCambio,Periodo,IdEmpresa,estMovBancos,idTalonario,Observacion1, Observacion2, idPendiente,idCliProv)  " & _
                          " VALUES (" & strcodcta_dol & "," & nmov & ",'" & strtipomov & "', " & _
                          " '" & Format(dtpDiaTrans.Value, "ddmmyy") & "','Cobranza Diaria', " & _
                          " '" & Format(dtpDiaTrans.Value, "yyyy-mm-dd") & "','I',0, " & _
                          " '" & Format(Now, "yyyy-mm-dd") & "', '" & glsUser & "'," & _
                          " '0'," & ntipcam & ",'" & Format(dtpDiaTrans.Value, "yyyymm") & "','" & glsEmpresa & "','ACT','0','','','','') "
                          
            Cn.Execute (csql_insert)
            nelemen = 0
            nvaltotalFacAcum = 0
            nvaltotalVueltoSOL = 0
            nvaltotalVueltoDOL = 0
            nvalPago = 0
            nvaltotalFac = 0
            cserie = ""
            cnumdoc = ""
            ctipodoc = ""
            strnomdoc = ""
            idGastosIngresos = ""
            
            strcodgto = traerCampo("parametros", "valParametro", "glsParametro", "TRANS_CONTA_CTA_12_DOLARES", True)
            Do While Not rsmovcajadetDol.EOF
                nelemen = nelemen + 1
                cserie = rsmovcajadetDol.Fields("idSerie") & ""
                cnumdoc = rsmovcajadetDol.Fields("iddocventas") & ""
                ctipodoc = rsmovcajadetDol.Fields("iddocumento") & ""
                strnomdoc = traerCampo("documentos", "abreDocumento", "iddocumento", ctipodoc, False)
                nvaltotalFac = Trim(traerCampo("docventas", "TotalPrecioVenta", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & Trim(txtCod_Sucursal.Text) & "'"))
                nvalPago = Val(Format(rsmovcajadetDol.Fields("valMonto"), "0.00"))
                strpagoSol = traerCampo("movcajasdet", "idmoneda", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & Trim(txtCod_Sucursal.Text) & "' and  idTipoMovCaja='99990002' and idmoneda='PEN' ")
                
                nvaltotalVueltoSOL = traerCampo("movcajasdet", "ValMonto", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & Trim(txtCod_Sucursal.Text) & "' and  idTipoMovCaja='99990003' and idmoneda='PEN' ")
                nvaltotalVueltoDOL = traerCampo("movcajasdet", "ValMonto", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & Trim(txtCod_Sucursal.Text) & "' and  idTipoMovCaja='99990003' and idmoneda='USD' ")
                 
                If Val(Format(nvaltotalVueltoSOL, "0.00")) > 0 Then
                    nvalPago = nvalPago - (Val(Format(nvaltotalVueltoSOL, "0.00")) / Val(Format(ntipcam, "0.000")))
                End If
                
                If Val(Format(nvaltotalVueltoDOL, "0.00")) > 0 Then
                    nvalPago = nvalPago - Val(Format(nvaltotalVueltoDOL, "0.00"))
                End If
                
                nvaltotalFacAcum = nvaltotalFacAcum + nvalPago
                strnomcli = Trim(traerCampo("docventas", "glsCliente", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & Trim(txtCod_Sucursal.Text) & "'"))
                strRUCCliente = Trim(traerCampo("docventas", "RUCCliente", "idDocumento", ctipodoc, True, " idDocVentas = '" & cnumdoc & "' AND idSerie = '" & cserie & "' AND idSucursal = '" & Trim(txtCod_Sucursal.Text) & "'"))
                
                idGastosIngresos = traerCampo("GastosIngresos", "idGastosIngresos", "idCtaContable", "" & traerCampo("Documentos", "idCtaVtaDolares", "idDocumento", ctipodoc, False), True, "indTipo ='I'")
                
                csql_insertDet = " INSERT INTO MovBancosDet (IdCtaCorriente,IdMovBancosCab,Item," & _
                      " Periodo,idCtaContable,GlsDetalle,NroDocumento,ValMonto,TipoCambio,indDebHab,idTipoDocumento, " & _
                      " IdGastosIngresos,SerieDocumento,Ruc_Dni,IdEmpresa, idCentroCosto,idPag_MovBancos, NumMovCompras, idPag_Dcto, idCtaCorrienteRec, idMovBancosCabRec,FecEmision) " & _
                      " VALUES (" & strcodcta_dol & "," & nmov & "," & nelemen & ", " & _
                      " '" & Format(dtpDiaTrans.Value, "yyyymm") & "','" & strcodgto & "','" & strnomcli & "','" & cnumdoc & "', " & _
                      " " & Val(Format(nvalPago, "0.00")) & "," & ntipcam & ", " & _
                      " " & "'H','" & strnomdoc & "', " & _
                      " '" & idGastosIngresos & "', '" & cserie & "' ,'" & strRUCCliente & "','" & glsEmpresa & "','','','','','','','" & Format(dtpDiaTrans.Value, "yyyy-mm-dd") & "') "
                Cn.Execute (csql_insertDet)
                
                rsmovcajadetDol.MoveNext
            Loop
            
            Cn.Execute ("Update Movcajas set indTransferido = 'S', CODCTA_DOL = " & strcodcta_dol & ", NUMMOV_DOL = " & nmov & " Where idmovcaja= '" & idmovcaja & "' And idEmpresa ='" & glsEmpresa & "' And idSucursal ='" & Trim(txtCod_Sucursal.Text) & "' ")
            Cn.Execute ("Update MovBancosCab Set TotalCab= " & Val(Format(nvaltotalFacAcum, "0.00")) & "  , TotalDet= " & Val(Format(nvaltotalFacAcum, "0.00")) & " Where idEmpresa='" & glsEmpresa & "' And Periodo='" & Format(dtpDiaTrans.Value, "yyyymm") & "' And IdCtaCorriente = " & strcodcta_dol & " And IdMovBancosCab = " & nmov & "  ")
         
        Else
            StrMsgError = "La caja no tiene movimientos": GoTo Err
        End If
    End If
            
    Cn.CommitTrans
    
    MsgBox "Se transfiró correctamente", vbInformation, App.Title
    
    Exit Sub
    
Err:
    If indTrans = True Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub cmbAyudaUsuario_Click()
    
    mostrarAyuda "USUARIO", txtCod_Usuario, txtGls_Usuario

End Sub

Private Sub Form_Load()

    Me.top = 0
    Me.left = 0
    dtpDiaTrans.Value = Format(getFechaSistema, "dd/mm/yyyy")

End Sub
 
Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        KeyAscii = 0
    End If

End Sub

 Private Sub txtCod_Usuario_Change()
    
    txtGls_Usuario.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Usuario.Text, False)

End Sub

Private Sub txtCod_Usuario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "USUARIO", txtCod_Usuario, txtGls_Usuario
        KeyAscii = 0
        If txtCod_Usuario.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub
