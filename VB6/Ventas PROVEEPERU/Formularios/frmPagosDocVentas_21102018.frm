VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmPagosDocVentas 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagos"
   ClientHeight    =   7620
   ClientLeft      =   3030
   ClientTop       =   2805
   ClientWidth     =   10755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraGrabaImp 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   3000
      TabIndex        =   37
      Top             =   2760
      Width           =   5055
      Begin VB.CommandButton BtnCancelar 
         Caption         =   "&Cancelar"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1140
      End
      Begin VB.CommandButton BtnGTFac 
         Caption         =   "&Factura"
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1140
      End
      Begin VB.CommandButton BtnGTBov 
         Caption         =   "&Boleta"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Imprimir Ticket"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   225
         Left            =   75
         TabIndex        =   38
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame fraTotales 
      Appearance      =   0  'Flat
      Caption         =   " Totales "
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
      Height          =   1770
      Left            =   6165
      TabIndex        =   23
      Top             =   5445
      Width           =   4060
      Begin CATControls.CATTextBox txt_TotalRecibido 
         Height          =   315
         Left            =   2280
         TabIndex        =   24
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":0000
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Vuelto 
         Height          =   315
         Left            =   2280
         TabIndex        =   25
         Top             =   600
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":001C
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_FormaPago 
         Height          =   285
         Left            =   45
         TabIndex        =   35
         Top             =   1305
         Visible         =   0   'False
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "frmPagosDocVentas.frx":0038
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_FormaPago 
         Height          =   285
         Left            =   180
         TabIndex        =   36
         Top             =   1305
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":0054
         Vacio           =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Vuelto S/.:"
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
         Height          =   240
         Left            =   1020
         TabIndex        =   27
         Top             =   660
         Width           =   945
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Total Recibido:"
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
         Height          =   240
         Left            =   1020
         TabIndex        =   26
         Top             =   300
         Width           =   1125
      End
   End
   Begin VB.Frame fraVuelto 
      Appearance      =   0  'Flat
      Caption         =   " Vuelto - Entregado "
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
      Height          =   1770
      Left            =   555
      TabIndex        =   31
      Top             =   5445
      Width           =   5535
      Begin CATControls.CATTextBox txtVal_VueltoEntregado 
         Height          =   315
         Left            =   3720
         TabIndex        =   32
         Top             =   1380
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":0070
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gVuelto 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmPagosDocVentas.frx":008C
         TabIndex        =   33
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Total Vuelto Entregado S/.:"
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
         Height          =   240
         Left            =   1620
         TabIndex        =   34
         Top             =   1440
         Width           =   2025
      End
   End
   Begin VB.Frame fraDatos 
      Appearance      =   0  'Flat
      Caption         =   " Datos del Documento "
      Enabled         =   0   'False
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
      Height          =   1515
      Left            =   150
      TabIndex        =   5
      Top             =   675
      Width           =   10530
      Begin CATControls.CATTextBox txt_NumDoc 
         Height          =   315
         Left            =   8550
         TabIndex        =   7
         Top             =   200
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Arial"
         FontSize        =   9
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":2168
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Serie 
         Height          =   315
         Left            =   6210
         TabIndex        =   8
         Top             =   200
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Arial"
         FontSize        =   9
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":2184
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TipoCambio 
         Height          =   315
         Left            =   8550
         TabIndex        =   9
         Top             =   675
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":21A0
         Text            =   "0"
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalBruto 
         Height          =   315
         Left            =   900
         TabIndex        =   13
         Top             =   1050
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":21BC
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalIGV 
         Height          =   315
         Left            =   4725
         TabIndex        =   14
         Top             =   1050
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":21D8
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalNeto 
         Height          =   315
         Left            =   8550
         TabIndex        =   15
         Top             =   1050
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmPagosDocVentas.frx":21F4
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   900
         TabIndex        =   19
         Top             =   675
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   16777152
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
         Container       =   "frmPagosDocVentas.frx":2210
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   1875
         TabIndex        =   20
         Top             =   675
         Width           =   4530
         _ExtentX        =   7990
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
         Container       =   "frmPagosDocVentas.frx":222C
         Vacio           =   -1  'True
      End
      Begin VB.Label lblDoc 
         Appearance      =   0  'Flat
         Caption         =   "Boleta de Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         TabIndex        =   22
         Top             =   225
         Width           =   3765
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
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
         Left            =   165
         TabIndex        =   21
         Top             =   765
         Width           =   570
      End
      Begin VB.Label lbl_TotalBruto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Bruto"
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
         Left            =   165
         TabIndex        =   18
         Top             =   1095
         Width           =   390
      End
      Begin VB.Label lbl_TotalIGV 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "IGV"
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
         Left            =   3960
         TabIndex        =   17
         Top             =   1050
         Width           =   270
      End
      Begin VB.Label lbl_TotalNeto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   8055
         TabIndex        =   16
         Top             =   1095
         Width           =   345
      End
      Begin VB.Label lbl_Serie 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   225
         Left            =   5670
         TabIndex        =   12
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lbl_NumDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   225
         Left            =   7740
         TabIndex        =   11
         Top             =   270
         Width           =   675
      End
      Begin VB.Label lbl_TC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "T/C"
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
         Left            =   8055
         TabIndex        =   10
         Top             =   765
         Width           =   240
      End
   End
   Begin VB.Frame frmPagos 
      Appearance      =   0  'Flat
      Caption         =   " Pagos "
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
      Height          =   3165
      Left            =   150
      TabIndex        =   3
      Top             =   2250
      Width           =   10530
      Begin DXDBGRIDLibCtl.dxDBGrid gPagos 
         Height          =   2775
         Left            =   60
         OleObjectBlob   =   "frmPagosDocVentas.frx":2248
         TabIndex        =   4
         Top             =   270
         Width           =   10395
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   75
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":6F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":730E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":7760
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":7AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":7E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":822E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":85C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":8962
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":8CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":9096
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":9430
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosDocVentas.frx":A0F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   1164
      ButtonWidth     =   1984
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Grabar     "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin CATControls.CATTextBox txtCod_Caja 
      Height          =   315
      Left            =   1665
      TabIndex        =   28
      Top             =   7245
      Width           =   915
      _ExtentX        =   1614
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
      MaxLength       =   8
      Container       =   "frmPagosDocVentas.frx":A48C
      Estilo          =   1
      EnterTab        =   -1  'True
   End
   Begin CATControls.CATTextBox txtGls_Caja 
      Height          =   315
      Left            =   2610
      TabIndex        =   29
      Top             =   7245
      Width           =   7635
      _ExtentX        =   13467
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
      Container       =   "frmPagosDocVentas.frx":A4A8
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Caja Activa"
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
      Left            =   585
      TabIndex        =   30
      Top             =   7290
      Width           =   825
   End
End
Attribute VB_Name = "frmPagosDocVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strTD           As String
Private strNumDoc       As String
Private strSerie        As String
Private strCodCliente   As String
Private strEstDoc       As String
Private fecEmision      As Date
Private indInserta      As Boolean
Private cfechavenc      As String
Private cruccliente     As String
Private cdireccliente   As String
Private cnombrecliente  As String
Private StrFPago        As String
Dim strTCC              As Double
Dim CIndEnviaCaja       As Boolean
Dim CTmpDocumentos          As String
Dim CTmpGuiasNF             As String
Dim CTmpDocumentosGen       As String

Private Sub GrabaTicketFactura()
Dim StrMsgError     As String
Dim strCodPersona   As String
Dim strglspersona   As String
Dim strRUCPersona   As String

On Error GoTo Err
    strRUCPersona = ""
    strglspersona = ""
    
    FrmBusca_Entidad.MostrarForm strRUCPersona, strglspersona
    If StrMsgError <> "" Then GoTo Err
    
    If Trim(strRUCPersona) <> "" Then
     ' Si el RUC no existe creamos la Entidad y actualizamos la entidad al cliente
     If Trim(traerCampo("Personas", "GlsPersona", "RUC", Trim(strRUCPersona), False)) = "" Then
         strCodPersona = ""
         strCodPersona = GeneraCorrelativoAnoMes("personas", "idPersona", False)
         csql = "Insert Into Personas(idPersona, GlsPersona, apellidoPaterno, apellidoMaterno, nombres, tipoPersona, ruc, idDistrito, direccion, FechaNacimiento, Telefonos, mail, direccionEntrega, GlsContacto, f2codcli, ef2cod, f2codtra, f2coduser, IdPais, Linea_Credito) " & _
                "Values ('" & strCodPersona & "','" & IIf(Trim(strglspersona) = "", Trim(strRUCPersona), Trim(strglspersona)) & "','','','','01002','" & Trim(strRUCPersona) & "','150101','XXXYYY','" & Format(getFechaSistema, "yyyy-mm-dd") & "','','','XXXYYY/LIMA/LIMA/LIMA','','','','','','02001',0)"
         Cn.Execute csql
         
         'Actualizamos como cliente
         csql = "INSERT INTO clientes(idCliente,idEmpresa,idGrupoCliente) VALUES('" & strCodPersona & "','" & glsEmpresa & "', '12')"
         Cn.Execute csql
         
         'Actualizamos el Cliente al Ticket
         csql = "Update Docventas Set idPerCliente = '" & strCodPersona & "', GlsCliente='" & IIf(Trim(strglspersona) = "", Trim(strRUCPersona), Trim(strglspersona)) & "', RUCCliente = '" & Trim(strRUCPersona) & "', dirCliente = 'X/LIMA/LIMA/LIMA' " & _
                "Where idEmpresa  ='" & glsEmpresa & "' And idDocumento  ='" & strTD & "' And idSerie = '" & txt_serie.Text & "' And idDocventas = '" & txt_numdoc.Text & "' And idSucursal  ='" & glsSucursal & "'"
         Cn.Execute csql
         
         strCodCliente = strCodPersona
     End If
    
     Toolbar1_ButtonClick Toolbar1.Buttons(1)
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub btnCancelar_Click()
    'Cancela
    FraGrabaImp.Visible = False
End Sub

Private Sub BtnGTBov_Click()
    'Ticket Boleta
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    Unload Me
End Sub

Private Sub BtnGTBov_GotFocus()
    If strTD = "12" Then
       BtnGTBov.SetFocus
    End If
End Sub

Private Sub BtnGTFac_Click()
    'Ticket Factura
    GrabaTicketFactura
    Unload Me
    
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError     As String

    indInserta = False
    Txt_TipoCambio.Decimales = glsDecimalesTC
    txt_TotalBruto.Decimales = glsDecimalesCaja
    txt_TotalIGV.Decimales = glsDecimalesCaja
    txt_TotalNeto.Decimales = glsDecimalesCaja
    txt_TotalRecibido.Decimales = glsDecimalesCaja
    txt_Vuelto.Decimales = glsDecimalesCaja
    txtVal_VueltoEntregado.Decimales = glsDecimalesCaja
    
    CTmpDocumentos = ""
    CTmpGuiasNF = ""
    CTmpDocumentosGen = ""
    
    gPagos.Columns.ColumnByFieldName("MontoOri").DecimalPlaces = glsDecimalesCaja
    gPagos.Columns.ColumnByFieldName("MontoSoles").DecimalPlaces = glsDecimalesCaja
    
    gVuelto.Columns.ColumnByFieldName("Vuelto").DecimalPlaces = glsDecimalesCaja
     
    ConfGrid gPagos, True, False, False, False
    ConfGrid gVuelto, True, False, False, False
    If traerCampo("parametros", "valparametro", "GLSPARAMETRO", "MODIFICA_FORMA_PAGO", True) = "N" Then
        gPagos.Columns.ColumnByFieldName("idFormadePago").DisableEditor = True
        gPagos.Columns.ColumnByFieldName("glsFormadePago").DisableEditor = True
    End If
    
    FraGrabaImp.Visible = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim objDocVentas    As New clsDocVentas
Dim RsP             As New ADODB.Recordset
Dim StrCodigo       As String
Dim strMsg          As String
Dim strCodMovCaja   As String
Dim stridVendedor   As String
Dim strReferencia   As String
Dim idFormaPago     As String
Dim cabrev          As String
Dim cdocumento      As String
Dim Cta_Dcto        As String
Dim cselect         As String
Dim fecCaja         As String
Dim Pagos           As Integer
Dim indTrans        As Boolean
Dim StrFPago        As String
Dim strCheque       As String
Dim ntotpagado      As Double
Dim ntotsaldo       As Double
Dim ctipofp         As String, ctipo As String, cdirectorio As String, cruta As String
Dim cconex_empresa  As String, cbusca As String, cinsert  As String
Dim cnn_empresa     As New ADODB.Connection
Dim rsbusca         As New ADODB.Recordset
Dim rscta           As New ADODB.Recordset
Dim ncorrela        As Double, ntc As Double, NTotal As Double
Dim ccliente        As String, cfechaemi As String
Dim cndebhab        As String
Dim cinsert_ctadcto As String
Dim cinsert_ctamvto As String
Dim cmonpago        As String
Dim npago           As Double
Dim ntotpago        As Double
Dim ncorrela_dep    As Double
Dim sw_resta        As Boolean
Dim strCodCtaDcto     As String
Dim strAbreDcto       As String
Dim strNumComprobante As String
Dim TipoFormaPago     As String
Dim tipoProducto      As String
Dim StrCcosto         As String
Dim strconsumido        As Double
Dim strCodMoneda        As String
Dim strtipocambio       As Double
Dim strutilizado        As Double
Dim strlineacredito     As Double
Dim strsaldo            As Double
Dim strmonto            As Double
Dim strcodmonedaCTD     As String
Dim stracumulado        As String
Dim strDocReferencia    As New frmDocVentas
Dim strconsumidoTOT     As Double
Dim strconsumidoLCR     As Double
Dim rst                 As New ADODB.Recordset
Dim CIndVtaGratuita                     As String
Dim NDias                               As Integer

    indTrans = True
    Cn.BeginTrans
    
    StrFPago = gPagos.Columns.ColumnByFieldName("idFormadePago").Value
    strCheque = traerCampo("FormasPagos", "cheque", "idFormaPago", StrFPago, True)
    If strCheque = "S" And gPagos.Columns.ColumnByFieldName("glsNumCheque").Value = "" Then
        StrMsgError = "Falta Ingresar NumCheque Verifique."
        GoTo Err
    End If

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    eliminaNulosGrilla
    
    If gPagos.Count >= 1 Then
        If gPagos.Count = 1 And gPagos.Columns.ColumnByFieldName("idFormadePago").Value = "" Then
            StrMsgError = "Falta Ingresar Pagos"
            GoTo Err
        End If
    End If
    
    '--- PARAMETRO QUE EVALUA PAGOS A CUENTA
    If traerCampo("Parametros", "valparametro", "Glsparametro", "PAGOS_A_CUENTA", True) = "N" Then
        If strTD <> "91" Then
            If Val(Format(txt_TotalRecibido.Value, "0.00")) < Val(Format(txt_TotalNeto.Value, "0.00")) Then
                StrMsgError = "El Monto recibido es menor al Monto por pagar"
                GoTo Err
            End If
        End If
    End If
    
    strCodMovCaja = Trim(traerCampo("docventas", "idMovCaja", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))
    stridVendedor = Trim(traerCampo("docventas", "idPerVendedorCampo", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))
    strReferencia = Trim(traerCampo("docventas", "GlsDocReferencia", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))
    
    If strCodMovCaja = "" Then
        strCodMovCaja = CajaAperturadaUsuario(0, StrMsgError)
        If StrMsgError <> "" Then GoTo Err
    End If
    
    '--- Verificamos si el documento ha sido al credito y tiene pagos
    cabrev = traerCampo("documentos", "AbreDocumento", "idDocumento", strTD, False)
    'cdocumento = cabrev & Format(txt_Serie.Text, "000") & "/" & Format(txt_NumDoc.Text, "00000000")
    cdocumento = cabrev & txt_serie.Text & "/" & Format(txt_numdoc.Text, "00000000")
    Cta_Dcto = traerCampo("Cta_Dcto", "idCta_Dcto", "Nro_Comp", cdocumento, True)
    Pagos = Val(traerCampo("Cta_Mvto", "count(*)", "idCta_Dcto", Cta_Dcto, True))
    
    If gPagos.Columns.ColumnByFieldName("IdTipoFormaPagoTemp").Value <> "06090001" And gPagos.Columns.ColumnByFieldName("IdTipoFormaPagoTemp").Value <> "06090004" And Pagos <> 0 Then
       StrMsgError = "No se puede Modificar el Pago del Documento"
       GoTo Err
    End If
    objDocVentas.EjecutaSQLFormPagosDocVentas Me, StrMsgError, strTD, strNumDoc, strSerie, strCodMovCaja, gPagos, gVuelto, strCodCliente
    If StrMsgError <> "" Then GoTo Err
    
    '---------------------------------------------------
    '---- ACTUALIZA EN CUENTAS POR COBRAR, SI ES CREDITO
    '---------------------------------------------------
    cselect = "SELECT idTipoFormaPago,fecVctos  " & _
              "FROM pagosdocventas m " & _
              "WHERE m.idempresa ='" & glsEmpresa & "' AND m.idsucursal = '" & glsSucursal & _
              "' AND m.iddocumento = '" & strTD & "' AND idserie='" & txt_serie.Text & "' AND iddocventas = '" & txt_numdoc.Text & "'"
              
    ctipofp = ""
    ctipo = ""
    If RsP.State = adStateOpen Then RsP.Close
    RsP.Open cselect, Cn, adOpenKeyset, adLockOptimistic
    If Not RsP.EOF Then
        ctipofp = Trim(RsP.Fields("idTipoFormaPago") & "")
    End If
              
    If Len(Trim(ctipofp)) > 0 Then
        ctipo = traerCampo("tipoformaspago", "TipoFormaPago", "idTipoFormaPago", ctipofp, False)
        cfechavenc = RsP.Fields("fecVctos") & ""
        cdirectorio = traerCampo("empresas", "Carpeta", "idEmpresa", glsEmpresa, False)
        
        If cdirectorio <> "" And glsSistemaAccess = "S" Then 'Grabamos en la Version de Access
            sw_grabactas = False
            If glsGraba_Contado = "S" Then
                sw_grabactas = True
            Else
                If ctipo = "R" Then sw_grabactas = True
            End If
            
            If sw_grabactas = True Then
                cruta = glsRuta_Access & cdirectorio
                If cnn_empresa.State = adStateOpen Then cnn_empresa.Close
                cconex_empresa = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cruta & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
                cnn_empresa.Open cconex_empresa
                
                cabrev = traerCampo("documentos", "AbreDocumento", "idDocumento", strTD, False)
                'cdocumento = cabrev & Format(txt_Serie.Text, "000") & "/" & Format(txt_NumDoc.Text, "0000000")
                cdocumento = cabrev & txt_serie.Text & "/" & Format(txt_numdoc.Text, "0000000")
                
                cbusca = "SELECT * FROM CTA_DCTO WHERE NRO_COMP='" & cdocumento & "'"
                If rscta.State = adStateOpen Then rscta.Close
                rscta.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
                If rscta.EOF Then
                    '--- NUEVO
                    '--------------------------------
                    '--- CORRELA
                    ncorrela = 0
                    cbusca = "SELECT CORRELA FROM CTA_DCTO ORDER BY CORRELA DESC"
                    If rsbusca.State = adStateOpen Then rsbusca.Close
                    rsbusca.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
                    If Not rsbusca.EOF Then
                        ncorrela = Val(rsbusca.Fields("CORRELA") & "") + 1
                    Else
                        ncorrela = 1
                    End If
                    rsbusca.Close: Set rsbusca = Nothing
                    '--------------------------------
                    '--- CLIENTE
                    ccliente = ""
                    cbusca = "SELECT F2CODCLI FROM EF2CLIENTES WHERE F2NEWRUC = '" & cruccliente & "'"
                    If rsbusca.State = adStateOpen Then rsbusca.Close
                    rsbusca.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
                    If Not rsbusca.EOF Then
                        ccliente = Trim(rsbusca.Fields("F2CODCLI") & "")
                    End If
                    rsbusca.Close: Set rsbusca = Nothing
                    
                    '--------------------------------
                    '---- AGREGA CLIENTE
                    If Len(ccliente) = 0 Then
                        ccliente = ""
                        cbusca = "SELECT F2CODCLI FROM EF2CLIENTES ORDER BY F2CODCLI DESC"
                        If rsbusca.State = adStateOpen Then rsbusca.Close
                        rsbusca.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
                        If Not rsbusca.EOF Then
                            ccliente = Format(Val(rsbusca.Fields("F2CODCLI") & "") + 1, "0000")
                        End If
                        rsbusca.Close: Set rsbusca = Nothing
                        
                        cinsert = "INSERT INTO EF2CLIENTES " & _
                                  "(F2CODCLI,F2NOMCLI,F2NEWRUC,F2DIRCLI,F2TIPDOC) " & _
                                  "VALUES ('" & ccliente & "','" & cnombrecliente & "','" & cruccliente & "','" & cdireccliente & "','J')"
                        cnn_empresa.Execute (cinsert)
                    End If
                    
                    '--------------------------------
                    cfechaemi = fecEmision
                    cmoneda = IIf(txtCod_Moneda.Text = "PEN", "S", "D")
                    ntc = Val(Txt_TipoCambio.Text & "")
                    NTotal = Val(Format(txt_TotalNeto.Text & "", "0.00"))
                    cndebhab = IIf(strTD = "07", "H", "D")
                    
                    cinsert = "INSERT INTO CTA_DCTO " & _
                              "(Via_ingr,CORRELA,NRO_COMP,FCH_COMP,FCH_VCTO,CLIENTE,CLIENTEO,MONEDA,MONEDAO,TCAMBIO,TCAMBIOO,TOTAL,TOTALO,SALDO,DEB_HAB) " & _
                              "VALUES ('1'," & ncorrela & ",'" & cdocumento & "',CVDATE('" & cfechaemi & "'),CVDATE('" & cfechavenc & "'),'" & _
                              ccliente & "','" & ccliente & "','" & cmoneda & "','" & cmoneda & "'," & ntc & "," & ntc & "," & NTotal & "," & NTotal & "," & NTotal & ",'" & cndebhab & "')"
                    cnn_empresa.Execute (cinsert)
                
                Else
                    '--- MODIFICA
                    If Val(Format(rscta.Fields("TOTAL"), "0.00")) = Val(Format(rscta.Fields("SALDO"), "0.00")) Then
                        NTotal = Val(Format(txt_TotalNeto.Text & "", "0.00"))
                        ccliente = rscta.Fields("CLIENTE")
                        ncorrela = Val(rscta.Fields("CORRELA") & "")
                        cinsert = "UPDATE CTA_DCTO SET TOTAL=" & NTotal & ", TOTALO = " & NTotal & ", SALDO = " & NTotal & " where correla = " & ncorrela & " "
                        cnn_empresa.Execute (cinsert)
                    End If
                End If
            End If
            
            If ctipo <> "R" And glsGraba_Contado = "S" Then
                '------------------GRABAR PAGO
                If gPagos.Count > 0 Then
                    gPagos.Dataset.First
                    ntotpago = 0
                    Do While Not gPagos.Dataset.EOF
                        '--- CORRELA
                        ncorrela_dep = 0
                        cbusca = "SELECT CORRELA FROM CTA_DCTO ORDER BY CORRELA DESC"
                        If rsbusca.State = adStateOpen Then rsbusca.Close
                        rsbusca.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
                        If Not rsbusca.EOF Then
                            ncorrela_dep = Val(rsbusca.Fields("CORRELA") & "") + 1
                        Else
                            ncorrela_dep = 1
                        End If
                        rsbusca.Close: Set rsbusca = Nothing
                        '---------------------------------------------------
                        cmonpago = IIf(gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN", "S", "D")
                        If txtVal_VueltoEntregado.Value > 0 Then
                            If cmonpago = "S" Then
                                npago = Val(Format(gPagos.Columns.ColumnByFieldName("MontoOri").Value, "0.00")) - Val(Format(txtVal_VueltoEntregado.Value, "0.00"))
                            Else
                                npago = gPagos.Columns.ColumnByFieldName("MontoOri").Value
                            End If
                        Else
                            npago = gPagos.Columns.ColumnByFieldName("MontoOri").Value
                        End If
                        If IIf(txtCod_Moneda.Text = "PEN", "S", "D") = "S" Then
                            If cmonpago = "S" Then
                                ntotpago = ntotpago + npago
                            Else
                                ntotpago = ntotpago + Val(Format(npago, "0.00")) * Val(Format(Txt_TipoCambio.Text, "0.000"))
                            End If
                        Else
                            If cmonpago = "D" Then
                                ntotpago = ntotpago + npago
                            Else
                                ntotpago = Val(Format(ntotpago, "0.00")) + Val(Format(npago, "0.00")) / Val(Format(Txt_TipoCambio.Text, "0.000"))
                            End If
                            
                        End If
                        cndebhab = IIf(strTD = "07", "D", "H")
                        
                        cinsert_ctadcto = "INSERT INTO CTA_DCTO " & _
                                          "(Via_ingr,CORRELA,NRO_COMP,FCH_COMP,FCH_VCTO,CLIENTE,CLIENTEO,MONEDA,MONEDAO,TCAMBIO,TCAMBIOO,TOTAL,TOTALO,SALDO,DEB_HAB) " & _
                                          "VALUES ('2'," & ncorrela_dep & ",'" & "Efe" & Format(Now, "dd/mm/yyyy") & "' ,CVDATE('" & Format(Now, "dd/mm/yyyy") & "'),CVDATE('" & Format(Now, "dd/mm/yyyy") & "'),'" & _
                                          ccliente & "','" & ccliente & "','" & cmonpago & "','" & cmonpago & "'," & ntc & "," & ntc & "," & npago & "," & npago & ",0,'" & cndebhab & "')"
                        cnn_empresa.Execute (cinsert_ctadcto)
                        
                        cinsert_ctamvto = " INSERT INTO CTA_MVTO " & _
                                         " (Cliente,corr_comp,corr_dcto,imputaso,tcambio,imputado,ano_repo,nro_repo , fch_mvto,fch_repo) " & _
                                         " VALUES('" & ccliente & "'," & ncorrela_dep & "," & ncorrela & "," & IIf(cmonpago = "S", npago, 0) & "," & ntc & "," & IIf(cmonpago = "D", npago, 0) & "," & _
                                         " CVDATE('" & Format(Now, "yyyy") & "') ,0,CVDATE('" & Format(Now, "dd/mm/yyyy") & "'),CVDATE('" & Format(Now, "dd/mm/yyyy") & "') ) "
                        cnn_empresa.Execute (cinsert_ctamvto)
                        
                        gPagos.Dataset.Next
                    Loop
                    
                End If
                '--- ACTUALIZAR SALDO
                If Val(Format(ntotpago, "0.00")) >= Val(Format(txt_TotalNeto.Value, "0.00")) Then
                    cnn_empresa.Execute ("UPDATE CTA_DCTO SET SALDO =0.00 WHERE CORRELA= " & ncorrela & " ")
                End If
                
                rscta.Close: Set rscta = Nothing
                cnn_empresa.Close: Set cnn_empresa = Nothing
            End If
        
        Else 'Grabamos en la Version de MySQL
            fecCaja = traerCampo("movcajas", "fecCaja", "idMovCaja", strCodMovCaja, True, "idSucursal = '" & glsSucursal & "'")
            strAbreDcto = traerCampo("documentos", "AbreDocumento", "idDocumento", strTD, False)
            strNumComprobante = strAbreDcto & txt_serie.Text & "/" & txt_numdoc.Text

            gPagos.Dataset.First
            If Cta_Dcto = "" Then
                strCodCtaDcto = GeneraCorrelativoAnoMesNuevo("cta_dcto", "idCta_Dcto", Year(getFechaSistema) & Format(Month(getFechaSistema), "00"), True)
            
            Else
                Cn.Execute ("DELETE FROM Cta_Dcto where IdEmpresa = '" & glsEmpresa & "' And idCta_Dcto = '" & Cta_Dcto & "'")
                
                If gPagos.Columns.ColumnByFieldName("IdTipoFormaPagoTemp").Value = "06090001" Or gPagos.Columns.ColumnByFieldName("IdTipoFormaPagoTemp").Value = "06090004" Then
                    Cn.Execute ("DELETE FROM Cta_Mvto where IdEmpresa = '" & glsEmpresa & "' And idCta_Dcto = '" & Cta_Dcto & "'")
                End If
                strCodCtaDcto = Cta_Dcto
            End If
            
            TipoFormaPago = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", gPagos.Columns.ColumnByFieldName("idFormadePago").Value, True)
            tipoProducto = traerCampo("docventasdet", "glsProducto", "iddocumento", strTD, True, "idserie = '" & txt_serie.Text & "' and iddocventas = '" & txt_numdoc.Text & "'")
            StrCcosto = Trim("" & traerCampo("docventas", "idcentrocosto", "iddocumento", strTD, True, "idserie = '" & txt_serie.Text & "' and iddocventas = '" & txt_numdoc.Text & "'"))
            
            '--- Si la Forma de pago es diferente a efectivo
            If TipoFormaPago <> "06090001" And TipoFormaPago <> "06090004" Then
                
                CIndVtaGratuita = Trim(traerCampo("docventas", "indVtaGratuita", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))
                If CIndVtaGratuita = "0" Then
                    CIndVtaGratuita = Trim(traerCampo("docventas", "IndTransGratuitaMP", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))
                End If
                If CIndVtaGratuita = "0" Then
                    CIndVtaGratuita = Trim(traerCampo("docventas", "IndTransGratuita", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))
                End If
                
                csql = "INSERT INTO cta_dcto " & _
                       "(idEmpresa,idSucursal,idSucursalCobro,indViaIngreso,idCta_Dcto,Nro_Comp,Fec_Comp,Fec_Vcto,idCliente,idClienteOri,idMoneda,idMonedaOri,ValTipoCambio,ValTipoCambioOri,ValTotal,ValTotalOri,ValSaldo,indDeb_Hab,IdVendedor,IdCobrador,glsreferencia, tipoProducto,idCentroCosto) " & _
                       "VALUES ('" & glsEmpresa & "','" & glsSucursal & "','" & glsSucursal & "',1,'" & strCodCtaDcto & "','" & strNumComprobante & "','" & Format(fecEmision, "yyyy-mm-dd") & "','" & Format("" & RsP.Fields("fecVctos"), "yyyy-mm-dd") & "','" & _
                       strCodCliente & "','" & strCodCliente & "','" & txtCod_Moneda.Text & "','" & txtCod_Moneda.Text & "'," & Txt_TipoCambio.Value & "," & Txt_TipoCambio.Value & "," & IIf(CIndVtaGratuita = "1", 0, Val(Format(txt_TotalNeto.Text, "0.00"))) & "," & IIf(CIndVtaGratuita = "1", 0, Val(Format(txt_TotalNeto.Text, "0.00"))) & "," & IIf(CIndVtaGratuita = "1", 0, Val(Format(txt_TotalNeto.Text, "0.00"))) & ",'D','" & stridVendedor & "', '" & stridVendedor & "','" & strReferencia & "', '" & left(tipoProducto, 250) & "','" & StrCcosto & "')"
                
                Cn.Execute (csql)
            
            Else
                '--- PARAMETRO QUE EVALUA PAGOS A CUENTA
                If traerCampo("Parametros", "valparametro", "Glsparametro", "PAGOS_A_CUENTA", True) = "S" Then
                     
                    csql = "INSERT INTO cta_dcto " & _
                            "(idEmpresa,idSucursal,idSucursalCobro,indViaIngreso,idCta_Dcto,Nro_Comp,Fec_Comp,Fec_Vcto,idCliente,idClienteOri,idMoneda,idMonedaOri,ValTipoCambio,ValTipoCambioOri,ValTotal,ValTotalOri,ValSaldo,indDeb_Hab,IdVendedor,IdCobrador,glsreferencia, TipoProducto,idCentroCosto) " & _
                            "VALUES ('" & glsEmpresa & "','" & glsSucursal & "','" & glsSucursal & "',1,'" & strCodCtaDcto & "','" & strNumComprobante & "','" & Format(fecEmision, "yyyy-mm-dd") & "','" & Format(fecEmision, "yyyy-mm-dd") & "','" & _
                            strCodCliente & "','" & strCodCliente & "','" & txtCod_Moneda.Text & "','" & txtCod_Moneda.Text & "'," & Txt_TipoCambio.Value & "," & Txt_TipoCambio.Value & "," & Val(Format(txt_TotalNeto.Text, "0.00")) & "," & Val(Format(txt_TotalNeto.Text, "0.00")) & ", 0, 'D','" & stridVendedor & "', '" & stridVendedor & "','" & strReferencia & "','" & left(tipoProducto, 250) & "','" & StrCcosto & "')"
                    Cn.Execute (csql)
    
                    glsEtiqPago = "Efe" & Format(fecEmision, "yyyy-mm-dd")
                    ncorrela_dep2 = GeneraCorrelativoAnoMesNuevo("cta_dcto", "idCta_Dcto", Year(getFechaSistema) & Format(Month(getFechaSistema), "00"), True)
    
                    cinsert_ctadcto = "INSERT INTO CTA_DCTO " & _
                                      "(idEmpresa,idSucursal,idSucursalCobro,indViaIngreso,idCta_Dcto,NRO_COMP,FEC_COMP,FEC_VCTO,idCliente,idClienteOri,idMoneda,idMonedaOri,ValTipoCambio,ValTipoCambioOri,ValTotal,ValTotalOri,ValSaldo,IndDeb_Hab) " & _
                                      "VALUES ('" & glsEmpresa & "','" & glsSucursal & "','" & glsSucursal & "','2','" & ncorrela_dep2 & "','" & glsEtiqPago & "' ,'" & Format(fecEmision, "yyyy-mm-dd") & "','" & Format(fecEmision, "yyyy-mm-dd") & "','" & _
                                      strCodCliente & "','" & strCodCliente & "','" & txtCod_Moneda.Text & "','" & txtCod_Moneda.Text & "'," & Txt_TipoCambio.Value & "," & Txt_TipoCambio.Value & "," & Val(Format(txt_TotalNeto.Text, "0.00")) & "," & Val(Format(txt_TotalNeto.Text, "0.00")) & ",0,'H')"
                    Cn.Execute (cinsert_ctadcto)
    
                    cinsert_ctamvto = " INSERT INTO CTA_MVTO " & _
                                     " (idEmpresa,idSucursal,idCliente,idCta_Dcto,idCta_Comp,ValimputaSO,ValTipoCambio,ValimputaDO,Fec_Mvto,Fec_Repo) " & _
                                     " VALUES('" & glsEmpresa & "','" & glsSucursal & "','" & strCodCliente & "','" & strCodCtaDcto & "','" & ncorrela_dep2 & "'," & IIf(txtCod_Moneda.Text = "PEN", Val(Format(txt_TotalNeto.Text, "0.00")), 0) & "," & Txt_TipoCambio.Value & "," & IIf(txtCod_Moneda.Text = "USD", Val(Format(txt_TotalNeto.Text, "0.00")), 0) & "," & _
                                     " '" & Format(fecEmision, "yyyy-mm-dd") & "', '" & Format(fecEmision, "yyyy-mm-dd") & "' ) "
                    Cn.Execute (cinsert_ctamvto)
                    
                Else
                    csql = "INSERT INTO cta_dcto " & _
                            "(idEmpresa,idSucursal,idSucursalCobro,indViaIngreso,idCta_Dcto,Nro_Comp,Fec_Comp,Fec_Vcto,idCliente,idClienteOri,idMoneda,idMonedaOri,ValTipoCambio,ValTipoCambioOri,ValTotal,ValTotalOri,ValSaldo,indDeb_Hab,IdVendedor,IdCobrador,glsreferencia, TipoProducto,idCentroCosto) " & _
                            "VALUES ('" & glsEmpresa & "','" & glsSucursal & "','" & glsSucursal & "',1,'" & strCodCtaDcto & "','" & strNumComprobante & "','" & Format(fecEmision, "yyyy-mm-dd") & "','" & Format(fecEmision, "yyyy-mm-dd") & "','" & _
                            strCodCliente & "','" & strCodCliente & "','" & txtCod_Moneda.Text & "','" & txtCod_Moneda.Text & "'," & Txt_TipoCambio.Value & "," & Txt_TipoCambio.Value & "," & Val(Format(txt_TotalNeto.Text, "0.00")) & "," & Val(Format(txt_TotalNeto.Text, "0.00")) & ", 0, 'D','" & stridVendedor & "', '" & stridVendedor & "','" & strReferencia & "','" & left(tipoProducto, 250) & "','" & StrCcosto & "')"
                    Cn.Execute (csql)
    
                    glsEtiqPago = "Efe" & Format(fecEmision, "yyyy-mm-dd")
                    ncorrela_dep2 = GeneraCorrelativoAnoMesNuevo("cta_dcto", "idCta_Dcto", Year(getFechaSistema) & Format(Month(getFechaSistema), "00"), True)
    
                    cinsert_ctadcto = "INSERT INTO CTA_DCTO " & _
                                      "(idEmpresa,idSucursal,idSucursalCobro,indViaIngreso,idCta_Dcto,NRO_COMP,FEC_COMP,FEC_VCTO,idCliente,idClienteOri,idMoneda,idMonedaOri,ValTipoCambio,ValTipoCambioOri,ValTotal,ValTotalOri,ValSaldo,IndDeb_Hab) " & _
                                      "VALUES ('" & glsEmpresa & "','" & glsSucursal & "','" & glsSucursal & "','2','" & ncorrela_dep2 & "','" & glsEtiqPago & "' ,'" & Format(fecEmision, "yyyy-mm-dd") & "','" & Format(fecEmision, "yyyy-mm-dd") & "','" & _
                                      strCodCliente & "','" & strCodCliente & "','" & txtCod_Moneda.Text & "','" & txtCod_Moneda.Text & "'," & Txt_TipoCambio.Value & "," & Txt_TipoCambio.Value & "," & Val(Format(txt_TotalNeto.Text, "0.00")) & "," & Val(Format(txt_TotalNeto.Text, "0.00")) & ",0,'H')"
                    Cn.Execute (cinsert_ctadcto)
    
                    cinsert_ctamvto = " INSERT INTO CTA_MVTO " & _
                                     " (idEmpresa,idSucursal,idCliente,idCta_Dcto,idCta_Comp,ValimputaSO,ValTipoCambio,ValimputaDO,Fec_Mvto,Fec_Repo) " & _
                                     " VALUES('" & glsEmpresa & "','" & glsSucursal & "','" & strCodCliente & "','" & strCodCtaDcto & "','" & ncorrela_dep2 & "'," & IIf(txtCod_Moneda.Text = "PEN", Val(Format(txt_TotalNeto.Text, "0.00")), 0) & "," & Txt_TipoCambio.Value & "," & IIf(txtCod_Moneda.Text = "USD", Val(Format(txt_TotalNeto.Text, "0.00")), 0) & "," & _
                                     " '" & Format(fecEmision, "yyyy-mm-dd") & "', '" & Format(fecEmision, "yyyy-mm-dd") & "' ) "
                    Cn.Execute (cinsert_ctamvto)
                End If
            End If
        End If
    End If
  
    strtipocambio = Trim(Txt_TipoCambio.Text)
    strmonto = Val(Format(txt_TotalNeto.Text & "", "0.00"))
    
    ActualizaLineaCredito StrMsgError, strCodCliente
    If StrMsgError <> "" Then GoTo Err
    
    If leeParametro("FACTURACION_ENTRE_EMPRESAS") = "S" Then
        
        NDias = Val("" & traerCampo("FormasPagos", "DiasVcto", "IdFormaPago", StrFPago, True))
        
        csql = "Update DocVentas A " & _
               "Inner Join DocVentasRegisDocFE B " & _
                   "On A.IdEmpresa = B.IdEmpresa And A.IdDocumento = B.IdDocumento And A.IdSerie = B.IdSerie And A.IdDocVentas = B.IdDocVentas " & _
               "Inner Join RegisDoc C " & _
                   "On B.IdEmpresaCli = C.IdEmpresa And B.Annio_Mov = C.Annio_Mov And B.IdMesMov = C.IdMesMov And B.IdNumMov = C.IdNumMov " & _
               "Set C.IdForma_Pago = '" & StrFPago & "',C.Fec_Vcto = AddDate(Cast(C.FechaMov As Date)," & NDias & ") " & _
               "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdDocumento = '" & strTD & "' And A.IdSerie = '" & txt_serie.Text & "' " & _
               "And A.IdDocVentas = '" & txt_numdoc.Text & "'"
               
        Cn.Execute csql
    
    End If
    
    Cn.CommitTrans
    
    CIndEnviaCaja = False
    
    gPagos.Dataset.Edit
    gPagos.Columns.ColumnByFieldName("IdTipoFormaPagoTemp").Value = gPagos.Columns.ColumnByFieldName("IdTipoFormaPago").Value
    gPagos.Dataset.Post
    
    RsP.Close: Set RsP = Nothing
    strMsg = "Grabo"
    Set objDocVentas = Nothing
    
    If strTD <> "12" Then
        Unload Me
    End If
    
    Exit Sub

Err:
    Set objDocVentas = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
    Exit Sub
    Resume
End Sub

Private Sub nuevo(ByRef StrMsgError As String)
On Error GoTo Err
    
    listaPagos StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    'Solo si el TD es Ticket
    If strTD = "12" Then
        gPagos.Columns.FocusedIndex = gPagos.Columns.ColumnByFieldName("MontoOri").Index - 1
    Else
        gPagos.Columns.FocusedIndex = gPagos.Columns.ColumnByFieldName("idFormadePago").Index
    End If
    
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err
Dim StrMsgError                 As String
    
    If Len(Trim(CTmpDocumentos)) > 0 Then
        
        Cn.Execute "Drop Table If Exists " & CTmpDocumentos
    
    End If
    
    If Len(Trim(CTmpGuiasNF)) > 0 Then
        
        Cn.Execute "Drop Table If Exists " & CTmpGuiasNF
    
    End If
    
    If Len(Trim(CTmpDocumentosGen)) > 0 Then
        
        Cn.Execute "Drop Table If Exists " & CTmpDocumentosGen
    
    End If

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gPagos_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gPagos.Columns.ColumnByFieldName("item").Value = gPagos.Count
        If txt_TotalNeto.Value - txt_TotalRecibido.Value Then
            gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN"
            gPagos.Columns.ColumnByFieldName("GlsMoneda").Value = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
            gPagos.Columns.ColumnByFieldName("MontoSoles").Value = txt_TotalNeto.Value - txt_TotalRecibido.Value
            gPagos.Columns.ColumnByFieldName("MontoOri").Value = txt_TotalNeto.Value - txt_TotalRecibido.Value
        Else
            gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN"
            gPagos.Columns.ColumnByFieldName("GlsMoneda").Value = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
            gPagos.Columns.ColumnByFieldName("MontoSoles").Value = 0
            gPagos.Columns.ColumnByFieldName("MontoOri").Value = 0
        End If
        gPagos.Dataset.Post
    End If

End Sub

Private Sub gPagos_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If Action = daInsert Then
        If gPagos.Columns.ColumnByFieldName("idFormadePago").Value = "" And indInserta = False Then
            Allow = False
        Else
            gPagos.Columns.FocusedIndex = gPagos.Columns.ColumnByFieldName("idFormadePago").Index
        End If
    End If

End Sub

Private Sub gPagos_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrDes As String
Dim StrCod As String
Dim TipoFormaPago As String
    
    Select Case Column.Index
        Case gPagos.Columns.ColumnByFieldName("idFormadePago").Index
            StrCod = gPagos.Columns.ColumnByFieldName("idFormadePago").Value
            StrDes = gPagos.Columns.ColumnByFieldName("glsFormadePago").Value
            
            If Trim("" & traerCampo("parametros", "valparametro", "GLSPARAMETRO", "FORMA_PAGOS_X_CLIENTE", True)) = "S" Then
                mostrarAyuda "FORMASPAGOxCLIENTE", txtCod_FormaPago, txtGls_FormaPago, " and A.idcliente = '" & Trim("" & strCodCliente) & "' and a.idempresa = '" & glsEmpresa & "' "
                StrCod = txtCod_FormaPago.Text
                StrDes = txtGls_FormaPago.Text
            Else
                mostrarAyudaTexto "FORMASPAGO", StrCod, StrDes
            End If

            gPagos.Dataset.Edit
            gPagos.Columns.ColumnByFieldName("idFormadePago").Value = StrCod
            gPagos.Columns.ColumnByFieldName("glsFormadePago").Value = StrDes
            gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", StrCod, True)
            If gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = "06090002" Then
                gPagos.Columns.ColumnByFieldName("fecVctos").Value = Format(DateAdd("D", traerCampo("formaspagos", "diasVcto", "idFormaPago", StrCod, True), CDate(fecEmision)), "dd/mm/yyyy")
            Else
                gPagos.Columns.ColumnByFieldName("fecVctos").Value = Format(fecEmision, "dd/mm/yyyy")
            End If
            gPagos.Columns.ColumnByFieldName("glsNumCheque").Value = ""
            gPagos.Columns.ColumnByFieldName("idBanco").Value = ""
            gPagos.Columns.ColumnByFieldName("glsBanco").Value = ""
            gPagos.Dataset.Post
        
        Case gPagos.Columns.ColumnByFieldName("idMoneda").Index
            StrCod = gPagos.Columns.ColumnByFieldName("idMoneda").Value
            StrDes = gPagos.Columns.ColumnByFieldName("glsMoneda").Value
            
            mostrarAyudaTexto "MONEDA", StrCod, StrDes
            
            gPagos.Dataset.Edit
            gPagos.Columns.ColumnByFieldName("idMoneda").Value = StrCod
            gPagos.Columns.ColumnByFieldName("glsMoneda").Value = StrDes
            If StrCod = "PEN" Then 'soles
                gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value
            Else 'dolares
                gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * Txt_TipoCambio.Value
            End If
            gPagos.Dataset.Post
            
        Case gPagos.Columns.ColumnByFieldName("idBanco").Index
            StrCod = gPagos.Columns.ColumnByFieldName("idBanco").Value
            StrDes = gPagos.Columns.ColumnByFieldName("GlsBanco").Value
            
            mostrarAyudaTexto "BANCO", StrCod, StrDes
            
            gPagos.Dataset.Edit
            gPagos.Columns.ColumnByFieldName("idBanco").Value = StrCod
            gPagos.Columns.ColumnByFieldName("GlsBanco").Value = StrDes
            gPagos.Dataset.Post
    End Select
    calcularTotales
    
End Sub

Private Sub gPagos_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrCod      As String
    
    strTCC = Val(traerCampo("tiposdecambio", "tcCompraC", "Year(fecha)", Year(fecEmision), False, "month(fecha) = '" & Month(fecEmision) & "' And day(fecha)='" & Day(fecEmision) & "'"))
   
    If traerCampo("Parametros", "valparametro", "Glsparametro", "PAGOS_A_CUENTA", True) = "S" Then
        Select Case gPagos.Columns.FocusedColumn.Index
            Case gPagos.Columns.ColumnByFieldName("MontoOri").Index
                gPagos.Dataset.Edit
                
                If gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN" Then 'soles
                     gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value
                 
                ElseIf txtCod_Moneda.Text = "USD" And gPagos.Columns.ColumnByFieldName("idMoneda").Value = "USD" Then
                    gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * Txt_TipoCambio.Value
                                                 
                ElseIf txtCod_Moneda.Text = "PEN" And gPagos.Columns.ColumnByFieldName("idMoneda").Value = "USD" Then
                    gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * strTCC
                ElseIf txtCod_Moneda.Text = "USD" And gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN" Then
                    gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * Txt_TipoCambio.Value
                Else 'dolares
                     gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * Txt_TipoCambio.Value
                End If
                
                gPagos.Dataset.Post
                calcularTotales
                
            Case gPagos.Columns.ColumnByFieldName("idBanco").Index
                StrCod = "" & gPagos.Columns.ColumnByFieldName("idBanco").Value
                If StrCod <> "" Then
                      gPagos.Columns.ColumnByFieldName("glsBanco").Value = traerCampo("Bancos", "GlsBanco", "idBanco", StrCod, False)
                End If
        End Select
        
   Else
        Select Case gPagos.Columns.FocusedColumn.Index
             Case gPagos.Columns.ColumnByFieldName("MontoOri").Index
                 gPagos.Dataset.Edit
                 If gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN" Then 'soles
                     gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value
                 Else 'dolares
                     gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * Txt_TipoCambio.Value
                 End If
                 gPagos.Dataset.Post
                 calcularTotales
                 
             Case gPagos.Columns.ColumnByFieldName("idBanco").Index
                 StrCod = "" & gPagos.Columns.ColumnByFieldName("idBanco").Value
                 If StrCod <> "" Then
                       gPagos.Columns.ColumnByFieldName("glsBanco").Value = traerCampo("Bancos", "GlsBanco", "idBanco", StrCod, False)
                 End If
         End Select
    End If
    
End Sub

Private Sub gPagos_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer
 

    If KeyCode = 46 Then
        If gPagos.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gPagos.Count = 1 Then
                    gPagos.Dataset.Edit
                    gPagos.Columns.ColumnByFieldName("Item").Value = 1
                    gPagos.Columns.ColumnByFieldName("idFormadePago").Value = ""
                    gPagos.Columns.ColumnByFieldName("glsFormadePago").Value = ""
                    gPagos.Columns.ColumnByFieldName("idMoneda").Value = ""
                    gPagos.Columns.ColumnByFieldName("GlsMoneda").Value = ""
                    gPagos.Columns.ColumnByFieldName("MontoOri").Value = 0#
                    gPagos.Columns.ColumnByFieldName("MontoSoles").Value = 0#
                    gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = ""
                    gPagos.Columns.ColumnByFieldName("fecVctos").Value = ""
                    gPagos.Columns.ColumnByFieldName("glsNumCheque").Value = ""
                    gPagos.Columns.ColumnByFieldName("idBanco").Value = ""
                    gPagos.Columns.ColumnByFieldName("glsBanco").Value = ""
                    gPagos.Dataset.Post
                
                Else
                    gPagos.Dataset.Delete
                    gPagos.Dataset.First
                    Do While Not gPagos.Dataset.EOF
                        i = i + 1
                        gPagos.Dataset.Edit
                        gPagos.Columns.ColumnByFieldName("Item").Value = i
                        gPagos.Dataset.Post
                        gPagos.Dataset.Next
                    Loop
                    If gPagos.Dataset.State = dsEdit Or gPagos.Dataset.State = dsInsert Then
                        gPagos.Dataset.Post
                    End If
                End If
                calcularTotales
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gPagos.Dataset.State = dsEdit Or gPagos.Dataset.State = dsInsert Then
            gPagos.Dataset.Post
        End If
        
        If strTD = "12" Then
            Select Case gPagos.Columns.FocusedColumn.Index
                 Case gPagos.Columns.ColumnByFieldName("FecVctos").Index
                    If MsgBox("Agregar Nuevo Item", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                        gPagos.Dataset.Insert
                    Else
                        FraGrabaImp.Visible = True
                        BtnGTBov.SetFocus
                    End If
                  End Select
        End If
    End If
 
End Sub

Private Sub gPagos_OnKeyPress(Key As Integer)
Dim StrCod As String
Dim StrDes As String
    
    If traerCampo("parametros", "valparametro", "GLSPARAMETRO", "MODIFICA_FORMA_PAGO", True) = "S" Then
        Exit Sub
    Else
        If Key <> 9 And Key <> 13 And Key <> 27 Then
            Select Case gPagos.Columns.FocusedColumn.Index
                Case gPagos.Columns.ColumnByFieldName("idFormadePago").Index
                    StrCod = gPagos.Columns.ColumnByFieldName("idFormadePago").Value
                    StrDes = gPagos.Columns.ColumnByFieldName("glsFormadePago").Value
                    
                    mostrarAyudaKeyasciiTexto Key, "FORMASPAGO", StrCod, StrDes
                    Key = 0
                    gPagos.Dataset.Edit
                    gPagos.Columns.ColumnByFieldName("idFormadePago").Value = StrCod
                    gPagos.Columns.ColumnByFieldName("glsFormadePago").Value = StrDes
                    gPagos.Columns.ColumnByFieldName("idTipoForm|aPago").Value = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", StrCod, True)
                    If gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = "06090002" Then 'CREDITO
                        gPagos.Columns.ColumnByFieldName("fecVctos").Value = Format(DateAdd("D", traerCampo("formaspagos", "diasVcto", "idFormaPago", StrCod, True), CDate(fecEmision)), "dd/mm/yyyy")
                    Else
                        gPagos.Columns.ColumnByFieldName("fecVctos").Value = ""
                    End If
                    gPagos.Dataset.Post
                    calcularTotales
                
                Case gPagos.Columns.ColumnByFieldName("idMoneda").Index
                    StrCod = gPagos.Columns.ColumnByFieldName("idMoneda").Value
                    StrDes = gPagos.Columns.ColumnByFieldName("glsMoneda").Value
                    
                    mostrarAyudaKeyasciiTexto Key, "MONEDA", StrCod, StrDes
                    Key = 0
                    gPagos.Dataset.Edit
                    gPagos.Columns.ColumnByFieldName("idMoneda").Value = StrCod
                    gPagos.Columns.ColumnByFieldName("glsMoneda").Value = StrDes
                    If StrCod = "PEN" Then 'soles
                        gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value
                    Else 'dolares
                        gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * Txt_TipoCambio.Value
                    End If
                    gPagos.Dataset.Post
                    calcularTotales
               
                Case gPagos.Columns.ColumnByFieldName("idBanco").Index
                    StrCod = gPagos.Columns.ColumnByFieldName("idBanco").Value
                    StrDes = gPagos.Columns.ColumnByFieldName("glsBanco").Value
                    
                    mostrarAyudaTexto "BANCO", StrCod, StrDes
                    Key = 0
                    gPagos.Dataset.Edit
                    gPagos.Columns.ColumnByFieldName("idBanco").Value = StrCod
                    gPagos.Columns.ColumnByFieldName("glsBanco").Value = StrDes
                    gPagos.Dataset.Post
                    calcularTotales
            End Select
        End If
    End If
    
End Sub

Private Sub gvuelto_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer
    
    If Action = daInsert Then
        gVuelto.Columns.ColumnByFieldName("item").Value = gVuelto.Count
        
        gVuelto.Dataset.Post
    End If

End Sub

Private Sub gvuelto_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If Action = daInsert Then
        If gVuelto.Columns.ColumnByFieldName("idMoneda").Value = "" And indInserta = False Then
            Allow = False
        Else
            gVuelto.Columns.FocusedIndex = gVuelto.Columns.ColumnByFieldName("idMoneda").Index
        End If
    End If

End Sub

Private Sub gvuelto_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrCod      As String
Dim StrDes      As String

    strTCC = Val(traerCampo("tiposdecambio", "tcCompraC", "Year(fecha)", Year(fecEmision), False, "month(fecha) = '" & Month(fecEmision) & "' And day(fecha)='" & Day(fecEmision) & "'"))
    
    Select Case Column.Index
        Case gVuelto.Columns.ColumnByFieldName("idMoneda").Index
            StrCod = gVuelto.Columns.ColumnByFieldName("idMoneda").Value
            StrDes = gVuelto.Columns.ColumnByFieldName("glsMoneda").Value
            mostrarAyudaTexto "MONEDA", StrCod, StrDes
            
            If existeEnGrilla(gVuelto, "idMoneda", StrCod) = False Then
                gVuelto.Dataset.Edit
                gVuelto.Columns.ColumnByFieldName("idMoneda").Value = StrCod
                gVuelto.Columns.ColumnByFieldName("glsMoneda").Value = StrDes
            Else
                MsgBox "La Moneda ya fue ingresada", vbInformation, App.Title
                Exit Sub
            End If
            If gVuelto.Count = 1 Then txtVal_VueltoEntregado.Text = 0
            
            'regresa valores para recalcular
            gVuelto.Columns.ColumnByFieldName("Vuelto").Value = 0#
            gVuelto.Dataset.Post
            calcularTotalVueltoEntregado
            gVuelto.Dataset.Edit
            
            '--- PARA LOS VUELTOS CALCULA CON EL TIPO DE CAMBIO COMPRA COMERCIAL(strTCC)
            'calcula saldo
            If StrCod = "PEN" Then
                gVuelto.Columns.ColumnByFieldName("Vuelto").Value = (txt_Vuelto.Value - txtVal_VueltoEntregado.Value)
            Else
                gVuelto.Columns.ColumnByFieldName("Vuelto").Value = ((txt_Vuelto.Value - txtVal_VueltoEntregado.Value) / IIf(strTCC = 0, 1, strTCC))
            End If
            gVuelto.Dataset.Post
            gVuelto.Columns.FocusedIndex = gVuelto.Columns.ColumnByFieldName("Vuelto").Index
            
            calcularTotalVueltoEntregado
    End Select
    
End Sub

Private Sub gvuelto_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    Select Case gVuelto.Columns.FocusedColumn.Index
        Case gVuelto.Columns.ColumnByFieldName("Vuelto").Index
                calcularTotalVueltoEntregado
    End Select

End Sub

Private Sub gvuelto_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gVuelto.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gVuelto.Count = 1 Then
                    gVuelto.Dataset.Edit
                    gVuelto.Columns.ColumnByFieldName("Item").Value = 1
                    gVuelto.Columns.ColumnByFieldName("idMoneda").Value = ""
                    gVuelto.Columns.ColumnByFieldName("GlsMoneda").Value = ""
                    gVuelto.Columns.ColumnByFieldName("Vuelto").Value = 0#
                    gVuelto.Dataset.Post
                
                Else
                    gVuelto.Dataset.Delete
                    gVuelto.Dataset.First
                    Do While Not gVuelto.Dataset.EOF
                        i = i + 1
                        gVuelto.Dataset.Edit
                        gVuelto.Columns.ColumnByFieldName("Item").Value = i
                        gVuelto.Dataset.Post
                        gVuelto.Dataset.Next
                    Loop
                    If gVuelto.Dataset.State = dsEdit Or gVuelto.Dataset.State = dsInsert Then
                        gVuelto.Dataset.Post
                    End If
                End If
                calcularTotales
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gVuelto.Dataset.State = dsEdit Or gVuelto.Dataset.State = dsInsert Then
              gVuelto.Dataset.Post
        End If
    End If

End Sub

Private Sub gvuelto_OnKeyPress(Key As Integer)
Dim StrCod As String
Dim StrDes As String

    If Key <> 9 And Key <> 13 And Key <> 27 Then
        Select Case gVuelto.Columns.FocusedColumn.Index
            Case gVuelto.Columns.ColumnByFieldName("idMoneda").Index
                StrCod = gVuelto.Columns.ColumnByFieldName("idMoneda").Value
                StrDes = gVuelto.Columns.ColumnByFieldName("glsMoneda").Value
                
                mostrarAyudaKeyasciiTexto Key, "MONEDA", StrCod, StrDes
                Key = 0
                gVuelto.Dataset.Edit
                gVuelto.Columns.ColumnByFieldName("idMoneda").Value = StrCod
                gVuelto.Columns.ColumnByFieldName("glsMoneda").Value = StrDes
                gVuelto.Dataset.Post
        End Select
    End If

End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2 'Cancelar
            Me.Hide
        Case 3 'Salir
            Me.Hide
    End Select

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
    Resume
End Sub

Private Sub listaPagos(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim strFormaPagoCliente As String
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim rsv As New ADODB.Recordset

    '--- TRAE EL LISTADO DE FORMAS DE PAGO DEL DOCUMENTO Y LO ALMACENA EN UN RECORSET
    csql = "SELECT p.item, p.idFormadePago,f.GlsFormaPago, p.idMoneda,m.GlsMoneda, p.MontoOri, p.MontoSoles, " & _
                  "P.idTipoFormaPago , P.fecVctos , P.glsNumCheque,p.IdBanco,p.GlsBanco " & _
           "FROM pagosdocventas p,formaspagos f,monedas m " & _
           "WHERE p.idEmpresa = '" & glsEmpresa & "' AND p.idSucursal = '" & glsSucursal & "' " & _
             "AND f.idEmpresa = '" & glsEmpresa & "' " & _
             "AND P.idFormadePago = f.idFormaPago " & _
             "AND p.idMoneda = m.idMoneda " & _
             "AND p.idDocumento = '" & strTD & "' AND  p.idDocVentas = '" & strNumDoc & "' AND p.idSerie = '" & strSerie & "' ORDER BY p.item"
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idFormadePago", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "glsFormadePago", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idMoneda", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsMoneda", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "MontoOri", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "MontoSoles", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idTipoFormaPago", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "fecVctos", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "glsNumCheque", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idBanco", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "glsBanco", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "IdTipoFormaPagoTemp", adVarChar, 8, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        
        strFormaPagoCliente = Trim("" & traerCampo("clientesformapagos", "idFormaPago", "idCliente", strCodCliente, True, " indestado = 1 "))
        If strFormaPagoCliente = "" Then strFormaPagoCliente = glsFormaPagoVentas
        
        rsg.Fields("idFormadePago") = strFormaPagoCliente
        rsg.Fields("glsFormadePago") = traerCampo("formaspagos", "GlsFormaPago", "idFormaPago", strFormaPagoCliente, True)
        rsg.Fields("idMoneda") = txtCod_Moneda.Text
        rsg.Fields("GlsMoneda") = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
        rsg.Fields("MontoOri") = txt_TotalNeto.Value
        
        If txtCod_Moneda.Text = "PEN" Then
            rsg.Fields("MontoSoles") = txt_TotalNeto.Value
        Else
            rsg.Fields("MontoSoles") = txt_TotalNeto.Value * Txt_TipoCambio.Text
        End If
        
        rsg.Fields("idTipoFormaPago") = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", strFormaPagoCliente, True)
        If rsg.Fields("idTipoFormaPago") = "06090002" Then
            rsg.Fields("fecVctos") = Format(DateAdd("D", traerCampo("formaspagos", "diasVcto", "idFormaPago", strFormaPagoCliente, True), CDate(fecEmision)), "dd/mm/yyyy")
        Else
            rsg.Fields("fecVctos") = Format(fecEmision, "dd/mm/yyyy")
        End If
        rsg.Fields("glsNumCheque") = ""
        rsg.Fields("idBanco") = ""
        rsg.Fields("glsBanco") = ""
        rsg.Fields("idTipoFormaPagoTemp") = ""
        
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = rst.Fields("Item")
            rsg.Fields("idFormadePago") = "" & rst.Fields("idFormadePago")
            rsg.Fields("glsFormadePago") = "" & rst.Fields("GlsFormaPago")
            rsg.Fields("idMoneda") = "" & rst.Fields("idMoneda")
            rsg.Fields("GlsMoneda") = "" & rst.Fields("GlsMoneda")
            rsg.Fields("MontoOri") = "" & rst.Fields("MontoOri")
            rsg.Fields("MontoSoles") = "" & rst.Fields("MontoSoles")
            rsg.Fields("idTipoFormaPago") = "" & rst.Fields("idTipoFormaPago")
            rsg.Fields("fecVctos") = "" & rst.Fields("fecVctos")
            rsg.Fields("glsNumCheque") = "" & rst.Fields("glsNumCheque")
            rsg.Fields("idBanco") = "" & rst.Fields("idBanco")
            rsg.Fields("glsBanco") = "" & rst.Fields("glsBanco")
            rsg.Fields("IdTipoFormaPagoTemp") = "" & rst.Fields("idTipoFormaPago")
            rst.MoveNext
        Loop
    End If
      
    mostrarDatosGridSQL gPagos, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    '--- TRAE LOS VUELTOS REGISTRADOS
    csql = "SELECT m.idMoneda,o.GlsMoneda, m.ValMonto " & _
           "FROM movcajasdet m,monedas o " & _
           "WHERE m.idMoneda = o.idMoneda " & _
             "AND m.idDocumento = '" & strTD & "' AND  m.idDocVentas = '" & strNumDoc & "' AND m.idSerie = '" & strSerie & "' " & _
             "AND m.idTipoMovCaja = '99990003' ORDER BY m.idMoneda"
             
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
    rsv.Fields.Append "Item", adInteger, , adFldRowID
    rsv.Fields.Append "idMoneda", adVarChar, 8, adFldIsNullable
    rsv.Fields.Append "GlsMoneda", adVarChar, 185, adFldIsNullable
    rsv.Fields.Append "Vuelto", adDouble, 14, adFldIsNullable
    rsv.Open
    
    If rst.RecordCount = 0 Then
        rsv.AddNew
        rsv.Fields("Item") = 1
        rsv.Fields("idMoneda") = ""
        rsv.Fields("GlsMoneda") = ""
        rsv.Fields("Vuelto") = 0
        
    Else
        Do While Not rst.EOF
            rsv.AddNew
            rsv.Fields("Item") = rst.RecordCount
            rsv.Fields("idMoneda") = "" & rst.Fields("idMoneda")
            rsv.Fields("GlsMoneda") = "" & rst.Fields("GlsMoneda")
            rsv.Fields("Vuelto") = "" & rst.Fields("ValMonto")
            rst.MoveNext
        Loop
    End If
    
    mostrarDatosGridSQL gVuelto, rsv, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    calcularTotales
    Me.Refresh
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_RUC_KeyPress(KeyAscii As Integer)
Dim StrMsgError As String
On Error GoTo Err

    If KeyAscii = "13" Then
       txtGls_RazonSocial.Text = traerCampo("Personas", "GlsPersona", "RUC", Trim(txt_RUC.Text), False)
       txtGls_RazonSocial.SetFocus
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub


Private Sub txtCod_Caja_Change()
    
    txtGls_Caja.Text = traerCampo("cajas", "GlsCaja", "idCaja", txtCod_Caja.Text, True)

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Private Sub calcularTotales()
Dim intFila         As Integer
Dim dblTotalNeto    As Double
Dim dblTotale       As Double
Dim indEval         As String

    strTCC = Val(traerCampo("tiposdecambio", "tcCompraC", "Year(fecha)", Year(fecEmision), False, "month(fecha) = '" & Month(fecEmision) & "' And day(fecha)='" & Day(fecEmision) & "'"))
    intFila = gPagos.Dataset.RecNo
    intFila = gPagos.Dataset.RecNo
    intFila = gPagos.Dataset.RecNo
    
    txt_TotalRecibido.Text = 0#
    txt_Vuelto.Text = 0#
    
    If traerCampo("Parametros", "valparametro", "Glsparametro", "PAGOS_A_CUENTA", True) = "S" Then
    gPagos.Dataset.First
        Do While Not gPagos.Dataset.EOF
            txt_TotalRecibido.Text = txt_TotalRecibido.Value + gPagos.Columns.ColumnByFieldName("MontoSoles").Value
            
            If txtCod_Moneda.Text = "USD" And Trim(gPagos.Columns.ColumnByFieldName("idMoneda").Value) = "USD" Then
                dblTotale = dblTotale + gPagos.Columns.ColumnByFieldName("MontoOri").Value * Txt_TipoCambio.Value
                indEval = "1"
                
            ElseIf txtCod_Moneda.Text = "USD" And gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN" Then
                dblTotale = dblTotale + gPagos.Columns.ColumnByFieldName("MontoSoles").Value
                indEval = "2"
                
            ElseIf (txtCod_Moneda.Text = "PEN" And gPagos.Columns.ColumnByFieldName("idMoneda").Value = "USD") Then
                dblTotale = dblTotale + gPagos.Columns.ColumnByFieldName("MontoOri").Value * strTCC
                indEval = "3"
                
            ElseIf (txtCod_Moneda.Text = "PEN" And gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN") Then
                dblTotale = dblTotale + gPagos.Columns.ColumnByFieldName("MontoOri").Value
                indEval = "4"
            Else
                dblTotale = dblTotale + gPagos.Columns.ColumnByFieldName("MontoOri").Value
            End If
            gPagos.Dataset.Next
        Loop
        gPagos.Dataset.RecNo = intFila
        
        If txtCod_Moneda.Text = "PEN" Then
            dblTotalNeto = Val(txt_TotalNeto.Value)
        Else
            If indEval = "1" Then
               dblTotalNeto = Val(txt_TotalNeto.Value) * Val(Txt_TipoCambio.Value)
               
            ElseIf indEval = "2" Then
                dblTotalNeto = Val(txt_TotalNeto.Value) * Val(Txt_TipoCambio.Value)
                
            ElseIf indEval = "3" Then
                dblTotalNeto = Val(txt_TotalNeto.Value) * strTCC
                
            ElseIf indEval = "4" Then
                dblTotalNeto = Val(txt_TotalNeto.Value) * Val(Txt_TipoCambio.Value)
                
            Else
                dblTotalNeto = Val(txt_TotalNeto.Value) * strTCC
            End If
        End If
        
        If Val(txt_TotalRecibido.Value) > dblTotalNeto Then
            If txtCod_Moneda.Text = "PEN" Then
                txt_Vuelto.Text = dblTotale - Val(txt_TotalNeto.Value)
            Else
                If indEval = "1" Then
                    txt_Vuelto.Text = dblTotale - (Val(txt_TotalNeto.Value) * Val(Txt_TipoCambio.Text))
                    
                ElseIf indEval = "2" Then
                    txt_Vuelto.Text = dblTotale - (Val(txt_TotalNeto.Value) * Val(Txt_TipoCambio.Text))
                    
                ElseIf indEval = "3" Then
                    txt_Vuelto.Text = dblTotale - (Val(txt_TotalNeto.Value) * strTCC)
                
                ElseIf indEval = "4" Then
                    txt_Vuelto.Text = dblTotale - (Val(txt_TotalNeto.Value) * Val(Txt_TipoCambio.Text))
                    
                Else
                    txt_Vuelto.Text = dblTotale - (Val(txt_TotalNeto.Value) * strTCC)
                End If
            End If
        End If
       
    Else
        gPagos.Dataset.First
        Do While Not gPagos.Dataset.EOF
            txt_TotalRecibido.Text = txt_TotalRecibido.Value + gPagos.Columns.ColumnByFieldName("MontoSoles").Value
            gPagos.Dataset.Next
        Loop
        
        gPagos.Dataset.RecNo = intFila
        
        If txtCod_Moneda.Text = "PEN" Then
            dblTotalNeto = Val(txt_TotalNeto.Value)
        Else
            dblTotalNeto = Val(txt_TotalNeto.Value) * Val(Txt_TipoCambio.Value)
        End If
        
        If Val(txt_TotalRecibido.Value) > dblTotalNeto Then
            txt_Vuelto.Text = Val(txt_TotalRecibido.Value) - dblTotalNeto
        End If
    End If
    
    gVuelto.Dataset.Edit
    If txt_Vuelto.Value > 0 Then
        fraVuelto.Enabled = True
        gVuelto.Columns.ColumnByFieldName("idMoneda").Value = "PEN"
        gVuelto.Columns.ColumnByFieldName("GlsMoneda").Value = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
        gVuelto.Columns.ColumnByFieldName("Vuelto").Value = txt_Vuelto.Value
    Else
        gVuelto.Columns.ColumnByFieldName("idMoneda").Value = ""
        gVuelto.Columns.ColumnByFieldName("GlsMoneda").Value = ""
        gVuelto.Columns.ColumnByFieldName("Vuelto").Value = 0#
        fraVuelto.Enabled = False
    End If
    
    gVuelto.Dataset.Post
    calcularTotalVueltoEntregado

End Sub

Private Sub calcularTotalVueltoEntregado()
Dim dblVuelto As Double
Dim intFila As Integer
    
    intFila = gVuelto.Dataset.RecNo
    intFila = gVuelto.Dataset.RecNo
    intFila = gVuelto.Dataset.RecNo

    txtVal_VueltoEntregado.Text = 0#
    
    gVuelto.Dataset.First
    Do While Not gVuelto.Dataset.EOF
        '--- PARA LOS VUELTOS CALCULA CON EL TIPO DE CAMBIO COMPRA COMERCIAL(strTCC)
        If gVuelto.Columns.ColumnByFieldName("idMoneda").Value <> "" Then
            If gVuelto.Columns.ColumnByFieldName("idMoneda").Value = "PEN" Then
                dblVuelto = dblVuelto + Val("" & gVuelto.Columns.ColumnByFieldName("Vuelto").Value)
            Else
                dblVuelto = dblVuelto + (Val("" & gVuelto.Columns.ColumnByFieldName("Vuelto").Value) * strTCC)
            End If
        End If
        gVuelto.Dataset.Next
    Loop
    
    txtVal_VueltoEntregado.Text = dblVuelto
    gVuelto.Dataset.RecNo = intFila
    
End Sub

Private Sub eliminaNulosGrilla()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer
    
    indWhile = True
    Do While indWhile = True
        If gPagos.Count >= 1 Then
            gPagos.Dataset.First
            indEntro = False
            Do While Not gPagos.Dataset.EOF
                If Trim(gPagos.Columns.ColumnByFieldName("idFormadePago").Value) = "" Then
                    gPagos.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gPagos.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gPagos.Count >= 1 Then
        gPagos.Dataset.First
        i = 0
        Do While Not gPagos.Dataset.EOF
            i = i + 1
            gPagos.Dataset.Edit
            gPagos.Columns.ColumnByFieldName("item").Value = i
            If gPagos.Dataset.State = dsEdit Then gPagos.Dataset.Post
            gPagos.Dataset.Next
        Loop
    Else
        indInserta = True
        gPagos.Dataset.Append
        indInserta = False
    End If
    
End Sub

Public Sub MostrarForm(ByVal strVarTipoDoc As String, ByVal strVarNumDoc As String, ByVal strVarSerie As String, ByVal fecVarEmision As Date, ByRef StrMsgError As String, PIndEnviarCaja As Boolean)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim strCodCaja      As String
Dim strCodMovCaja   As String

    strTD = strVarTipoDoc
    strNumDoc = strVarNumDoc
    strSerie = strVarSerie
    fecEmision = fecVarEmision
    CIndEnviaCaja = PIndEnviarCaja
   
     
    
    strCodMovCaja = Trim(traerCampo("docventas", "idMovCaja", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))
    If strCodMovCaja = "" Then
        txtCod_Caja.Text = CajaAperturadaUsuario(1, StrMsgError)
        If StrMsgError <> "" Then GoTo Err
    Else
        strCodMovCaja = traerCampo("movcajas", "idCaja", "idMovCaja", strCodMovCaja, True, "idSucursal = '" & glsSucursal & "'")
        txtCod_Caja.Text = strCodMovCaja
    End If
    
    txt_serie.Text = strSerie
    txt_numdoc.Text = strNumDoc
    lblDoc.Caption = traerCampo("documentos", "GlsDocumento", "idDocumento", strTD, False)
    
    csql = "SELECT d.idFormaPago,d.idPerCliente,d.idMoneda, d.TipoCambio, d.TotalValorVenta, d.TotalIGVVenta, d.TotalPrecioVenta, d.estDocVentas, " & _
           "d.FecPago, d.RUCCliente, d.dirCliente, d.GlsCliente " & _
           "FROM docventas d " & _
           "WHERE d.idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "'"
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rst.EOF Then
        strCodCliente = "" & rst.Fields("idPerCliente")
        
        cfechavenc = rst.Fields("FecPago")
        cruccliente = rst.Fields("RUCCliente") & ""
        cdireccliente = rst.Fields("dirCliente") & ""
        cnombrecliente = rst.Fields("GlsCliente") & ""
        
        txtCod_Moneda.Text = "" & rst.Fields("idMoneda")
        Txt_TipoCambio.Text = "" & rst.Fields("TipoCambio")
        strEstDoc = "" & rst.Fields("estDocVentas")
        txt_TotalBruto.Text = "" & rst.Fields("TotalValorVenta")
        txt_TotalIGV.Text = "" & rst.Fields("TotalIGVVenta")
        txt_TotalNeto.Text = "" & rst.Fields("TotalPrecioVenta")
        
        StrFPago = Trim("" & rst.Fields("idFormaPago"))
        If strEstDoc = "IMP" Then
            frmPagos.Enabled = False
            Toolbar1.Buttons(1).Visible = False
        Else
            frmPagos.Enabled = True
            Toolbar1.Buttons(1).Visible = True
        End If
        
    End If

    nuevo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If Trim("" & traerCampo("parametros", "valparametro", "GLSPARAMETRO", "FORMA_PAGOS_X_CLIENTE", True)) = "S" Then
        gPagos.Dataset.Edit
        
        If Len(Trim("" & StrFPago)) = 0 Then
            gPagos.Columns.ColumnByFieldName("idFormadePago").Value = Trim("" & traerCampo("clientesformapagos", "idFormaPago", "idCliente", strCodCliente, True))
        Else
            gPagos.Columns.ColumnByFieldName("idFormadePago").Value = StrFPago
        End If

        gPagos.Columns.ColumnByFieldName("glsFormadePago").Value = Trim("" & traerCampo("formaspagos", "GlsFormaPago", "idFormaPago", StrFPago, True))
        gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", StrFPago, True)
        If gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = "06090002" Then
            gPagos.Columns.ColumnByFieldName("fecVctos").Value = Format(DateAdd("D", traerCampo("formaspagos", "diasVcto", "idFormaPago", StrFPago, True), CDate(fecEmision)), "dd/mm/yyyy")
        Else
            gPagos.Columns.ColumnByFieldName("fecVctos").Value = Format(fecEmision, "dd/mm/yyyy")
        End If
        gPagos.Dataset.Post
    End If
    
    frmPagosDocVentas.Show 1
    
    PIndEnviarCaja = CIndEnviaCaja
    Unload frmPagosDocVentas
    
    If rst.State = 1 Then rst.Close: Set rst = Nothing

    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub habilitaBotones(indexBoton)
    
    Select Case indexBoton
        Case 1
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = True
        Case 2
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = True
    End Select
 
End Sub

Private Sub ActualizaLineaCredito(StrMsgError As String, PIdCliente As String)
On Error GoTo Err
Dim CSqlC                       As String
Dim CPC                         As String
        
    If Len(Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "MONEDA_LINEA_CREDITO", True))) > 0 Then 'Si tiene asignada la Moneda se entiende que debe controlar
        
        If Len(Trim(CTmpDocumentos)) = 0 Then
        
            CPC = ComputerName
            CPC = Replace(CPC, "-", "")
            CPC = Trim(CPC)
            
            CTmpDocumentos = "TmpDocumentosV" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & CPC)
            
        End If
        
        If Len(Trim(CTmpGuiasNF)) = 0 Then
    
            CPC = ComputerName
            CPC = Replace(CPC, "-", "")
            CPC = Trim(CPC)
            
            CTmpGuiasNF = "TmpGuiasNFV" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & CPC)
            
        End If
        
        If Len(Trim(CTmpDocumentosGen)) = 0 Then
    
            CPC = ComputerName
            CPC = Replace(CPC, "-", "")
            CPC = Trim(CPC)
            
            CTmpDocumentosGen = "TmpDocumentosGenV" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & CPC)
            
        End If
        
        CSqlC = "Call Spu_VerificaMorosos('" & glsEmpresa & "','" & PIdCliente & "','0','" & CTmpDocumentos & "','" & CTmpGuiasNF & "','" & CTmpDocumentosGen & "')"
                
        Cn.Execute CSqlC
                
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
