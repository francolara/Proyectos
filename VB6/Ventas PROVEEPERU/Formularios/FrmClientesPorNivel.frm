VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmClientesPorNivel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Clientes por Nivel"
   ClientHeight    =   4680
   ClientLeft      =   4785
   ClientTop       =   2220
   ClientWidth     =   7665
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7665
   Begin VB.Frame Frame1 
      Height          =   3885
      Left            =   90
      TabIndex        =   11
      Top             =   45
      Width           =   7440
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
         Left            =   6800
         Picture         =   "FrmClientesPorNivel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1500
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaTipoDoc 
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
         Left            =   6800
         Picture         =   "FrmClientesPorNivel.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1080
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
         Left            =   6800
         Picture         =   "FrmClientesPorNivel.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   700
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
         Left            =   6800
         Picture         =   "FrmClientesPorNivel.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   350
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   1
         Left            =   270
         TabIndex        =   13
         Top             =   1935
         Width           =   6915
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   4
            Top             =   300
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   103940097
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
            TabIndex        =   5
            Top             =   300
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   103940097
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3915
            TabIndex        =   15
            Top             =   375
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   855
            TabIndex        =   14
            Top             =   360
            Width           =   465
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1300
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   330
         Width           =   950
         _ExtentX        =   1667
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
         Container       =   "FrmClientesPorNivel.frx":0E28
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2280
         TabIndex        =   17
         Top             =   330
         Width           =   4500
         _ExtentX        =   7938
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
         Container       =   "FrmClientesPorNivel.frx":0E44
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1300
         TabIndex        =   1
         Tag             =   "TidPerCliente"
         Top             =   705
         Width           =   950
         _ExtentX        =   1667
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
         Container       =   "FrmClientesPorNivel.frx":0E60
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2280
         TabIndex        =   20
         Tag             =   "TGlsCliente"
         Top             =   705
         Width           =   4500
         _ExtentX        =   7938
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
         Locked          =   -1  'True
         Container       =   "FrmClientesPorNivel.frx":0E7C
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Documento 
         Height          =   315
         Left            =   1300
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1095
         Width           =   950
         _ExtentX        =   1667
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
         Container       =   "FrmClientesPorNivel.frx":0E98
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Documento 
         Height          =   315
         Left            =   2280
         TabIndex        =   23
         Top             =   1095
         Width           =   4500
         _ExtentX        =   7938
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
         Container       =   "FrmClientesPorNivel.frx":0EB4
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1300
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   1485
         Width           =   950
         _ExtentX        =   1667
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
         Container       =   "FrmClientesPorNivel.frx":0ED0
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2280
         TabIndex        =   26
         Top             =   1485
         Width           =   4500
         _ExtentX        =   7938
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
         Container       =   "FrmClientesPorNivel.frx":0EEC
         Vacio           =   -1  'True
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   14
         Left            =   270
         TabIndex        =   12
         Top             =   2835
         Width           =   6915
         Begin VB.OptionButton OptSerie 
            Caption         =   "Detallado por Serie"
            Height          =   240
            Left            =   4590
            TabIndex        =   8
            Top             =   300
            Width           =   1800
         End
         Begin VB.OptionButton OptGeneral 
            Caption         =   "General"
            Height          =   240
            Left            =   765
            TabIndex        =   6
            Top             =   300
            Value           =   -1  'True
            Width           =   1530
         End
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado"
            Height          =   240
            Left            =   2745
            TabIndex        =   7
            Top             =   300
            Width           =   1350
         End
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   315
         TabIndex        =   27
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label lbldocumento 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   315
         TabIndex        =   24
         Top             =   1155
         Width           =   810
      End
      Begin VB.Label lbl_Cliente 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   315
         TabIndex        =   21
         Top             =   765
         Width           =   480
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   315
         TabIndex        =   18
         Top             =   375
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4095
      Width           =   1230
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4095
      Width           =   1230
   End
End
Attribute VB_Name = "FrmClientesPorNivel"
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

Private Sub cmbAyudaTipoDoc_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError     As String
Dim fIni            As String
Dim Ffin            As String
    
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    If OptGeneral.Value = True Then
        mostrarReporte "rptClientesXNivelGeneral.rpt", "parEmpresa|parSucursal|parCliente|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Cliente.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & fIni & "|" & Ffin, "Clientes por Nivel - General", StrMsgError
    ElseIf OptDetallado.Value = True Then
        mostrarReporte "rptClientesXNivelDetallado.rpt", "parEmpresa|parSucursal|parCliente|parDocumento|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Cliente.Text) & "|" & Trim(txtCod_Documento.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
    ElseIf OptSerie.Value = True Then
        mostrarReporte "rptNivelesXSerie.rpt", "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
    End If
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    lbldocumento.Enabled = False
    txtCod_Documento.Enabled = False
    txtGls_Documento.Enabled = False
    CmbAyudaTipoDoc.Enabled = False
    
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    txtGls_Documento.Text = "TODOS LOS DOCUMENTOS"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    txtCod_Moneda.Text = "PEN"

End Sub

Private Sub OptDetallado_Click()
    
    lbldocumento.Enabled = True
    txtCod_Documento.Enabled = True
    txtGls_Documento.Enabled = True
    CmbAyudaTipoDoc.Enabled = True

End Sub

Private Sub OptGeneral_Click()
    
    lbldocumento.Enabled = False
    txtCod_Documento.Enabled = False
    txtGls_Documento.Enabled = False
    CmbAyudaTipoDoc.Enabled = False

End Sub

Private Sub OptSerie_Click()
    
    lbldocumento.Enabled = True
    txtCod_Documento.Enabled = True
    txtGls_Documento.Enabled = True
    CmbAyudaTipoDoc.Enabled = True
    
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

Private Sub txtCod_Documento_Change()
    
    If txtCod_Cliente.Text <> "" Then
        txtGls_Documento.Text = traerCampo("Documentos", "GlsDocumento", "idDocumento", txtCod_Documento.Text, False)
    Else
        txtGls_Documento.Text = "TODOS LOS DOCUMENTOS"
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
