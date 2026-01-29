VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmrptReproductores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Venta de Reproductores"
   ClientHeight    =   4080
   ClientLeft      =   3870
   ClientTop       =   1950
   ClientWidth     =   7545
   DrawMode        =   12  'Nop
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7545
   Begin VB.CommandButton Command1 
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
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3510
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3885
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3510
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Height          =   3300
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   7350
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6680
         Picture         =   "FrmrptReproductores.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   750
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaCliente 
         Height          =   315
         Left            =   6680
         Picture         =   "FrmrptReproductores.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   350
         Width           =   390
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
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
         Height          =   810
         Left            =   225
         TabIndex        =   12
         Top             =   2250
         Width           =   6870
         Begin VB.OptionButton Optdet 
            Caption         =   "Detallado"
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
            Left            =   4455
            TabIndex        =   5
            Top             =   270
            Width           =   1095
         End
         Begin VB.OptionButton OptGen 
            Caption         =   "General"
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
            Left            =   1350
            TabIndex        =   4
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
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
         Height          =   810
         Index           =   1
         Left            =   225
         TabIndex        =   9
         Top             =   1260
         Width           =   6870
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   300
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
            Format          =   133955585
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4605
            TabIndex        =   3
            Top             =   300
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
            Format          =   133955585
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4140
            TabIndex        =   11
            Top             =   375
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   990
            TabIndex        =   10
            Top             =   375
            Width           =   465
         End
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1100
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
         Container       =   "FrmrptReproductores.frx":0714
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2025
         TabIndex        =   14
         Top             =   330
         Width           =   4635
         _ExtentX        =   8176
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
         Container       =   "FrmrptReproductores.frx":0730
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1100
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
         Container       =   "FrmrptReproductores.frx":074C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2025
         TabIndex        =   17
         Top             =   735
         Width           =   4635
         _ExtentX        =   8176
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
         Container       =   "FrmrptReproductores.frx":0768
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
         Left            =   270
         TabIndex        =   18
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   270
         TabIndex        =   15
         Top             =   375
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmrptReproductores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub Command1_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim fIni As String
Dim Ffin As String
Dim Tipo As String

    fIni = Format(dtpfInicio, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal, "yyyy-mm-dd")
    
    If OptGen.Value = True Then
        Tipo = "1"
        mostrarReporte "rptListaReproductores_Gen.rpt", "parEmpresa|parSucursal|parCliente|parFecIni|parFecFin|parTipo", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & txtCod_Cliente.Text & "|" & fIni & "|" & Ffin & "|" & Tipo, "Lista de Reproductores - General", StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Else
        Tipo = "2"
        mostrarReporte "rptListaReproductores_Det.rpt", "parEmpresa|parSucursal|parCliente|parFecIni|parFecFin|parTipo", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & txtCod_Cliente.Text & "|" & fIni & "|" & Ffin & "|" & Tipo, "Lista de Reproductores - Detallado", StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Command2_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    OptGen.Value = True
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"

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

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    End If
    Me.Caption = Me.Caption & " - " & txtGls_Sucursal.Text
    
End Sub

Private Sub txtCod_Sucursal_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Sucursal.Text = ""
    End If

End Sub

Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        KeyAscii = 0
    End If

End Sub
