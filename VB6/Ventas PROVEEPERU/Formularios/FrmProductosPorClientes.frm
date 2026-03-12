VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmProductosPorClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos por Cliente"
   ClientHeight    =   5190
   ClientLeft      =   4605
   ClientTop       =   2385
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4515
      Left            =   90
      TabIndex        =   12
      Top             =   45
      Width           =   7080
      Begin VB.CheckBox ChkMuestrasG 
         Caption         =   "Sólo Muestras Gratuitas"
         Height          =   240
         Left            =   225
         TabIndex        =   27
         Top             =   4095
         Width           =   2895
      End
      Begin VB.CommandButton cmbAyudaMoneda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6500
         Picture         =   "FrmProductosPorClientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1515
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6500
         Picture         =   "FrmProductosPorClientes.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   730
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaSucursal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6500
         Picture         =   "FrmProductosPorClientes.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   325
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   840
         Index           =   1
         Left            =   225
         TabIndex        =   14
         Top             =   2070
         Width           =   6645
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   4
            Top             =   345
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
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
            Format          =   174981121
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
            TabIndex        =   5
            Top             =   345
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
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
            Format          =   174981121
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3960
            TabIndex        =   16
            Top             =   375
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   15
            Top             =   375
            Width           =   465
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         ForeColor       =   &H00000000&
         Height          =   840
         Index           =   14
         Left            =   225
         TabIndex        =   13
         Top             =   3060
         Width           =   6645
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado"
            Height          =   240
            Left            =   4320
            TabIndex        =   7
            Top             =   345
            Width           =   1665
         End
         Begin VB.OptionButton OptGeneral 
            Caption         =   "Resumen"
            Height          =   240
            Left            =   1350
            TabIndex        =   6
            Top             =   345
            Value           =   -1  'True
            Width           =   2025
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   330
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
         Container       =   "FrmProductosPorClientes.frx":0A9E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2145
         TabIndex        =   18
         Top             =   330
         Width           =   4315
         _ExtentX        =   7620
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
         Container       =   "FrmProductosPorClientes.frx":0ABA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Tag             =   "TidMoneda"
         Top             =   735
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
         Container       =   "FrmProductosPorClientes.frx":0AD6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2145
         TabIndex        =   21
         Top             =   735
         Width           =   4315
         _ExtentX        =   7620
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
         Container       =   "FrmProductosPorClientes.frx":0AF2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Serie 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1125
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
         MaxLength       =   4
         Container       =   "FrmProductosPorClientes.frx":0B0E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   1515
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
         Container       =   "FrmProductosPorClientes.frx":0B2A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2145
         TabIndex        =   25
         Top             =   1515
         Width           =   4315
         _ExtentX        =   7620
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
         Container       =   "FrmProductosPorClientes.frx":0B46
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   26
         Top             =   1575
         Width           =   570
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   23
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   22
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   19
         Top             =   375
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   5010
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   400
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton cmdanexonc 
      Caption         =   "A&nexo N/C"
      Height          =   400
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton cmdanexond 
      Caption         =   "An&exo N/D"
      Height          =   400
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1200
   End
End
Attribute VB_Name = "FrmProductosPorClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaCliente_Click()

    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmdaceptar_Click()
    
    imprimir

End Sub

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Private Sub cmdanexonc_Click()
    
    ImprimirAnexos "07"

End Sub

Private Sub cmdanexond_Click()
    
    ImprimirAnexos "08"

End Sub

Private Sub Form_Load()

    Me.top = 0
    Me.left = 0

    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    txtCod_Moneda.Text = "PEN"
    
End Sub

Private Sub txt_Serie_LostFocus()

    txt_serie.Text = Format(txt_serie.Text, "0000")
    
End Sub

Private Sub txtCod_Cliente_Change()

    If txtCod_Cliente.Text <> "" Then
       If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
            If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
                txtGls_Cliente.Text = traerCampo("personas p Inner Join  clientes c On  p.idPersona = c.idCliente Inner Join personas v On c.idVendedorCampo = v.idPersona", "p.GlsPersona", "p.idPersona", txtCod_Cliente.Text, False, "c.idVendedorCampo ='" & glsUser & "' ")
            Else
                txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
            End If
        Else
                txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
        End If
     Else
        txtGls_Cliente.Text = "TODOS LOS CLIENTES"
     End If
     
End Sub

Private Sub txtCod_Moneda_Change()

    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
    
End Sub

Private Sub txtCod_Sucursal_Change()

    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    
End Sub

Private Sub ImprimirAnexos(ByRef Documento As String)
On Error GoTo Err
Dim GlsForm As String
Dim StrMsgError As String
Dim CMuestra                As String

    GlsForm = "Productos por Clientes"
    
    CMuestra = ""
    
    If ChkMuestrasG.Value = "1" Then
        
        CMuestra = "1"
    
    End If
    
    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If Documento = "07" Then
        GlsForm = "ANEXO DE NOTAS DE CREDITO"
    Else
        GlsForm = "ANEXO DE NOTAS DE DEBITO"
    End If

    If OptGeneral.Value = True Then
        mostrarReporte "ProductosClientesAnexos.rpt", "parEmpresa|parSucursal|parDocumento|parSerie|parMoneda|parFecDesde|parFecHasta|parCliente|parMuestras", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & Documento & "|" & Trim(txt_serie.Text) & "|" & txtCod_Moneda.Text & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & CMuestra, GlsForm, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Else
        mostrarReporte "ProductosClientesAnexosDet.rpt", "parEmpresa|parSucursal|parDocumento|parSerie|parMoneda|parFecDesde|parFecHasta|parCliente|parMuestras", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & Documento & "|" & Trim(txt_serie.Text) & "|" & txtCod_Moneda.Text & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & CMuestra, GlsForm, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
        
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub imprimir()
On Error GoTo Err
Dim GlsForm                 As String
Dim StrMsgError             As String
Dim CMuestra                As String

    GlsForm = "Productos por Clientes"
    
    CMuestra = ""
    
    If ChkMuestrasG.Value = "1" Then
        
        CMuestra = "1"
    
    End If
    
    If traerCampo("Parametros", "Valparametro", "GlsParametro", "FORMATOREPORTECLIENTESPRODUCTO", True) = "1" Then
        validaFormSQL Me, StrMsgError
        If StrMsgError <> "" Then GoTo Err
     
        If OptGeneral.Value = True Then
            mostrarReporte "ProductosClientes.rpt", "parEmpresa|parSucursal|parSerie|parMoneda|parFecDesde|parFecHasta|parCliente|parMuestras", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & Trim(txt_serie.Text) & "|" & txtCod_Moneda.Text & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & CMuestra, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Else
            mostrarReporte "ProductosClientesDet.rpt", "parEmpresa|parSucursal|parSerie|parMoneda|parFecDesde|parFecHasta|parCliente|parMuestras", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & Trim(txt_serie.Text) & "|" & txtCod_Moneda.Text & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & CMuestra, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
    Else
        If OptGeneral.Value = True Then
            mostrarReporte "ProductosClientesFormato2.rpt", "parEmpresa|parSucursal|parSerie|parMoneda|parFecDesde|parFecHasta|parCliente|parMuestras", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & Trim(txt_serie.Text) & "|" & txtCod_Moneda.Text & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & CMuestra, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Else
            mostrarReporte "ProductosClientesDetFormato2.rpt", "parEmpresa|parSucursal|parSerie|parMoneda|parFecDesde|parFecHasta|parCliente|parMuestras", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & Trim(txt_serie.Text) & "|" & txtCod_Moneda.Text & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & CMuestra, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
    End If
   
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
