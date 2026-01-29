VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form ayuda_proveedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de Proveedores"
   ClientHeight    =   5730
   ClientLeft      =   5895
   ClientTop       =   5400
   ClientWidth     =   9735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   45
      TabIndex        =   22
      Top             =   945
      Width           =   9600
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4395
         Left            =   180
         OleObjectBlob   =   "ayuda_proveedores.frx":0000
         TabIndex        =   1
         Top             =   180
         Width           =   9375
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   0
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
            Picture         =   "ayuda_proveedores.frx":2020
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":23BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":280C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":2F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":32DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":3674
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":3A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":3DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":4142
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":44DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ayuda_proveedores.frx":519E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   45
      TabIndex        =   20
      Top             =   180
      Width           =   9600
      Begin VB.TextBox txtbusqueda 
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
         Height          =   315
         Left            =   1035
         TabIndex        =   0
         Top             =   270
         Width           =   8430
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
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
         Left            =   180
         TabIndex        =   21
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6990
      Left            =   150
      TabIndex        =   23
      Top             =   780
      Width           =   8160
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   17070
         Picture         =   "ayuda_proveedores.frx":5538
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   5040
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaTipoPersona 
         Height          =   315
         Left            =   7500
         Picture         =   "ayuda_proveedores.frx":58C2
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   630
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaPais 
         Height          =   315
         Left            =   7500
         Picture         =   "ayuda_proveedores.frx":5C4C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2090
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaDepa 
         Height          =   315
         Left            =   7500
         Picture         =   "ayuda_proveedores.frx":5FD6
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2460
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaProv 
         Height          =   315
         Left            =   7500
         Picture         =   "ayuda_proveedores.frx":6360
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2835
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaDistrito 
         Height          =   315
         Left            =   7500
         Picture         =   "ayuda_proveedores.frx":66EA
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3210
         Width           =   390
      End
      Begin VB.CheckBox chk_Cliente 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6750
         TabIndex        =   19
         Top             =   5775
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmbAyudaVendedorCampo 
         Height          =   315
         Left            =   16980
         Picture         =   "ayuda_proveedores.frx":6A74
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4275
         Width           =   390
      End
      Begin MSComCtl2.DTPicker dtpFechaNac 
         Height          =   315
         Left            =   1650
         TabIndex        =   16
         Tag             =   "FFechaNacimiento"
         Top             =   4920
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   132317185
         CurrentDate     =   38638
      End
      Begin CATControls.CATTextBox txtCod_Persona 
         Height          =   315
         Left            =   6975
         TabIndex        =   30
         Tag             =   "TidPersona"
         Top             =   240
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
         Container       =   "ayuda_proveedores.frx":6DFE
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Paterno 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Tag             =   "TapellidoPaterno"
         Top             =   1020
         Width           =   6240
         _ExtentX        =   11007
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
         MaxLength       =   80
         Container       =   "ayuda_proveedores.frx":6E1A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Materno 
         Height          =   315
         Left            =   1650
         TabIndex        =   3
         Tag             =   "TapellidoMaterno"
         Top             =   1365
         Width           =   6240
         _ExtentX        =   11007
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
         MaxLength       =   80
         Container       =   "ayuda_proveedores.frx":6E36
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nombres 
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Tag             =   "Tnombres"
         Top             =   1730
         Width           =   6240
         _ExtentX        =   11007
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
         MaxLength       =   80
         Container       =   "ayuda_proveedores.frx":6E52
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Direccion 
         Height          =   315
         Left            =   1650
         TabIndex        =   10
         Tag             =   "Tdireccion"
         Top             =   3600
         Width           =   6240
         _ExtentX        =   11007
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
         MaxLength       =   255
         Container       =   "ayuda_proveedores.frx":6E6E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_RUC 
         Height          =   315
         Left            =   1650
         TabIndex        =   11
         Tag             =   "Truc"
         Top             =   3970
         Width           =   2115
         _ExtentX        =   3731
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
         MaxLength       =   11
         Container       =   "ayuda_proveedores.frx":6E8A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Telefono 
         Height          =   315
         Left            =   5940
         TabIndex        =   12
         Tag             =   "TTelefonos"
         Top             =   3975
         Width           =   1965
         _ExtentX        =   3466
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
         MaxLength       =   85
         Container       =   "ayuda_proveedores.frx":6EA6
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Email 
         Height          =   315
         Left            =   1650
         TabIndex        =   17
         Tag             =   "Tmail"
         Top             =   5280
         Width           =   6240
         _ExtentX        =   11007
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
         MaxLength       =   80
         Container       =   "ayuda_proveedores.frx":6EC2
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_TipoPersona 
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Tag             =   "TtipoPersona"
         Top             =   645
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
         Container       =   "ayuda_proveedores.frx":6EDE
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TipoPersona 
         Height          =   315
         Left            =   2580
         TabIndex        =   31
         Top             =   645
         Width           =   4890
         _ExtentX        =   8625
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
         Container       =   "ayuda_proveedores.frx":6EFA
      End
      Begin CATControls.CATTextBox txtCod_Pais 
         Height          =   315
         Left            =   1650
         TabIndex        =   6
         Tag             =   "TidPais"
         Top             =   2100
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
         Container       =   "ayuda_proveedores.frx":6F16
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Pais 
         Height          =   315
         Left            =   2580
         TabIndex        =   32
         Top             =   2100
         Width           =   4890
         _ExtentX        =   8625
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
         Container       =   "ayuda_proveedores.frx":6F32
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Depa 
         Height          =   315
         Left            =   1650
         TabIndex        =   7
         Top             =   2475
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
         Container       =   "ayuda_proveedores.frx":6F4E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Depa 
         Height          =   315
         Left            =   2580
         TabIndex        =   33
         Top             =   2475
         Width           =   4890
         _ExtentX        =   8625
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
         Container       =   "ayuda_proveedores.frx":6F6A
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Prov 
         Height          =   315
         Left            =   1650
         TabIndex        =   8
         Top             =   2850
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
         Container       =   "ayuda_proveedores.frx":6F86
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Prov 
         Height          =   315
         Left            =   2580
         TabIndex        =   34
         Top             =   2850
         Width           =   4890
         _ExtentX        =   8625
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
         Container       =   "ayuda_proveedores.frx":6FA2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Distrito 
         Height          =   315
         Left            =   1650
         TabIndex        =   9
         Tag             =   "TidDistrito"
         Top             =   3210
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
         Container       =   "ayuda_proveedores.frx":6FBE
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Distrito 
         Height          =   315
         Left            =   2580
         TabIndex        =   35
         Top             =   3210
         Width           =   4890
         _ExtentX        =   8625
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
         Container       =   "ayuda_proveedores.frx":6FDA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtGlsPersona 
         Height          =   285
         Left            =   750
         TabIndex        =   36
         Tag             =   "TGlsPersona"
         Top             =   210
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   -2147483633
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
         Container       =   "ayuda_proveedores.frx":6FF6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_DireccionFiscal 
         Height          =   315
         Left            =   1650
         TabIndex        =   14
         Tag             =   "TdireccionEntrega"
         Top             =   4380
         Width           =   6240
         _ExtentX        =   11007
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
         MaxLength       =   255
         Container       =   "ayuda_proveedores.frx":7012
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_VendedorCampo 
         Height          =   315
         Left            =   11130
         TabIndex        =   13
         Top             =   4275
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
         Container       =   "ayuda_proveedores.frx":702E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_VendedorCampo 
         Height          =   315
         Left            =   12060
         TabIndex        =   37
         Top             =   4275
         Width           =   4890
         _ExtentX        =   8625
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
         Container       =   "ayuda_proveedores.frx":704A
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtNro_Brevete 
         Height          =   315
         Left            =   1650
         TabIndex        =   18
         Top             =   5655
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
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
         Container       =   "ayuda_proveedores.frx":7066
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtcod_obliga 
         Height          =   315
         Left            =   11220
         TabIndex        =   15
         Top             =   5040
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
         Container       =   "ayuda_proveedores.frx":7082
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtdescobliga 
         Height          =   315
         Left            =   12150
         TabIndex        =   57
         Top             =   5040
         Width           =   4890
         _ExtentX        =   8625
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
         Container       =   "ayuda_proveedores.frx":709E
         Vacio           =   -1  'True
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Grupo Obligacion"
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
         Left            =   9705
         TabIndex        =   58
         Top             =   5040
         Width           =   1245
      End
      Begin VB.Label lblFecNacimiento 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha Nac."
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
         Left            =   135
         TabIndex        =   54
         Top             =   4920
         Width           =   825
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
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
         Left            =   135
         TabIndex        =   53
         Top             =   4095
         Width           =   450
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos"
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
         Left            =   5115
         TabIndex        =   52
         Top             =   4020
         Width           =   720
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "E Mail"
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
         Left            =   135
         TabIndex        =   51
         Top             =   5325
         Width           =   405
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Left            =   135
         TabIndex        =   50
         Top             =   3675
         Width           =   675
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
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
         Left            =   135
         TabIndex        =   49
         Top             =   2820
         Width           =   660
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
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
         Left            =   135
         TabIndex        =   48
         Top             =   2445
         Width           =   1005
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "País"
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
         Left            =   135
         TabIndex        =   47
         Top             =   2160
         Width           =   300
      End
      Begin VB.Label lblApePaterno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
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
         Left            =   135
         TabIndex        =   46
         Top             =   1110
         Width           =   1170
      End
      Begin VB.Label lblApeMaterno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
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
         Left            =   135
         TabIndex        =   45
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblNombres 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
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
         Left            =   135
         TabIndex        =   44
         Top             =   1770
         Width           =   645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Entidad"
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
         Left            =   135
         TabIndex        =   43
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   6300
         TabIndex        =   42
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
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
         Left            =   135
         TabIndex        =   41
         Top             =   3270
         Width           =   495
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Direc. Entrega"
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
         Left            =   135
         TabIndex        =   40
         Top             =   4425
         Width           =   1020
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor Campo"
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
         Left            =   9615
         TabIndex        =   39
         Top             =   4320
         Width           =   1260
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nº Brevete"
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
         Left            =   135
         TabIndex        =   38
         Top             =   5730
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1164
      ButtonWidth     =   2858
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "            Nuevo          "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Lista"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "ayuda_proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw_limpia   As Boolean
Private strCodCliente As String
Private indCopiaDireccion As Boolean
Private indEditando As Boolean
Dim chk As String
Dim gsw As String

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo As String
Dim strMsg      As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaHomonimia "personas", "GlsPersona", "idPersona", txtGlsPersona.Text, txtCod_Persona.Text, False, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If Trim(txt_RUC.Text) <> "" Then
        validaHomonimia "personas", "ruc", "idPersona", txt_RUC.Text, txtCod_Persona.Text, False, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    strCodCliente = ""
    
    If txtCod_Persona.Text = "" Then 'graba
        txtCod_Persona.Text = GeneraCorrelativoAnoMes("personas", "idPersona", False)
        EjecutaSQLForm_1 Me, 0, False, "personas", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        actualizaCliente StrMsgError
        strMsg = "Grabo"
        strCodCliente = txtCod_Persona.Text
    
    Else 'modifica
        EjecutaSQLForm_1 Me, 1, False, "personas", StrMsgError, "idpersona"
        If StrMsgError <> "" Then GoTo Err
        actualizaCliente StrMsgError
        strMsg = "Modifico"
    End If
    fraGeneral.Enabled = False
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub cmbAyudaDepa_Click()
    
    mostrarAyuda "DEPARTAMENTO", txtCod_Depa, txtGls_Depa, " AND idPais = '" & txtCod_Pais.Text & "'"
    If txtCod_Depa.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaDistrito_Click()
    
    mostrarAyuda "DISTRITO", txtCod_Distrito, txtGls_Distrito, " AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Depa.Text & "' and idProv = '" + txtCod_Prov.Text + "'"
    If txtCod_Distrito.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaPais_Click()
    
    mostrarAyuda "PAIS", txtCod_Pais, txtGls_Pais
    If txtCod_Pais.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaProv_Click()
    
    mostrarAyuda "PROVINCIA", txtCod_Prov, txtGls_Prov, " AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Depa.Text + "'"
    If txtCod_Prov.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaTipoPersona_Click()

    mostrarAyuda "TIPOPERSONA", txtCod_TipoPersona, txtGls_TipoPersona
    txtCod_TipoPersona_LostFocus
    If txtCod_TipoPersona.Text <> "" Then SendKeys "{Tab}"

End Sub

Private Sub cmbAyudaVendedorCampo_Click()
    
    mostrarAyuda "VENDEDORCAMPO", txtCod_VendedorCampo, txtGls_VendedorCampo

End Sub

'Private Sub Command1_Click()
'
'    Ayuda_GrupoObliga.Show 1
'    If Len(wcod_obliga) <> 0 Then
'        txtcod_obliga.Text = "" & Trim(wcod_obliga)
'        txtdescobliga.Text = "" & Trim(wdes_obliga)
'    End If
'
'End Sub

Private Sub dtpFechaNac_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub fill()
Dim csql    As String
Dim rsdatos     As New ADODB.Recordset

'    csql = "SELECT p.idPersona cod ,p.GlsPersona des, p.ruc, p.Telefonos,concat(p.direccion,' ',u.glsUbigeo,' ',d.glsUbigeo) as direccion " & _
'           "FROM personas p INNER JOIN proveedores pro ON pro.idEmpresa = '" & glsEmpresa & "' AND p.idPersona = pro.idproveedor " & _
'           "LEFT JOIN ubigeo u ON P.idDistrito = u.idDistrito And P.idPais = u.idPais " & _
'           "LEFT JOIN ubigeo d ON left(u.idDistrito,2) = d.idDpto AND d.idProv = '00' AND d.idDist = '00' And   u.idPais = d.idPais  " & _
'           "where p.ruc LIKE '%" & txtbusqueda.Text & "%' OR " & " p.glspersona LIKE '%" & txtbusqueda.Text & "%'"
           
    csql = "SELECT p.idPersona cod ,p.GlsPersona des, p.ruc, p.Telefonos,concat(p.direccion,' ',ISNull(glsUbigeo1,''),' ',ISNull(A.glsUbigeo2,'')) as direccion " & _
           "FROM personas p  " & _
           "INNER JOIN proveedores pro " & _
              "ON p.idPersona = pro.idproveedor AND pro.idEmpresa = '" & glsEmpresa & "' " & _
           "LEFT JOIN ( " & _
              "SELECT u.glsUbigeo As glsUbigeo1, d.glsUbigeo As glsUbigeo2,u.idDistrito,u.idPais " & _
              "FROM ubigeo u " & _
              "INNER JOIN ubigeo d " & _
                "ON left(u.idDistrito,2) = d.idDpto AND u.idPais = d.idPais AND d.idProv = '00' AND d.idDist = '00') A " & _
              "ON P.idDistrito = A.idDistrito  AND P.idPais = A.idPais " & _
            "WHERE pro.idEmpresa  = '" & glsEmpresa & "' And  (p.ruc LIKE '%" & txtbusqueda.Text & "%' OR  p.glspersona LIKE '%" & txtbusqueda.Text & "%')"
           
If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
    
Set dxDBGrid1.DataSource = rsdatos

'    With dxDBGrid1
'         .DefaultFields = False
'         .Dataset.ADODataset.ConnectionString = strcn
'         .Dataset.ADODataset.CursorLocation = clUseClient
'         .Dataset.Active = False
'         .Dataset.ADODataset.CommandText = csql
'         .Dataset.DisableControls
'         .Dataset.Active = True
'         .KeyField = "cod"
'    End With

End Sub

Private Sub dxDBGrid1_OnDblClick()
    
    cod_Prov = "" & dxDBGrid1.Columns.ColumnByFieldName("cod").Value
    des_prov = "" & dxDBGrid1.Columns.ColumnByFieldName("des").Value
    ruc_prov = "" & dxDBGrid1.Columns.ColumnByFieldName("ruc").Value
    dir_prov = "" & dxDBGrid1.Columns.ColumnByFieldName("direccion").Value
     
    sw_limpia = True
    txtbusqueda.Text = ""
    sw_limpia = False
    
    Me.Hide
    
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.m.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
    End With

End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    
    Select Case KeyCode
        Case 13:
            dxDBGrid1_OnDblClick
        Case 27:
            cod_Prov = ""
            des_prov = ""
            ruc_prov = ""
            dir_prov = ""
            Me.Hide
    End Select
       
End Sub

Private Sub Form_Activate()

    If sw_proveedor = False Then
        dxDBGrid1.Dataset.Close
        dxDBGrid1.OptionEnabled = 0
        dxDBGrid1.Dataset.Refresh
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.SetFocus
        
        dxDBGrid1.OptionEnabled = 1
        
        TxtBusqueda_KeyPress 13
        
        dxDBGrid1.Dataset.First
        txtbusqueda.SetFocus
        
        ConfGrid dxDBGrid1, False, False, False, False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    dxDBGrid1.Dataset.Close
    
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.SetFocus
    End If

End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
        If Len(Trim(txtbusqueda.Text)) > 0 Then
            fill
        End If
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    If sw_proveedor = False Then
        fraGeneral.Visible = False
        Toolbar1.Visible = False
        Frame1.Visible = True
        Frame2.Visible = True
        Me.Height = 6135
        Me.Width = 9780
    Else
        fraGeneral.Visible = True
        Toolbar1.Visible = True
        Frame1.Visible = False
        Frame2.Visible = False
        Me.Height = 8205
        Me.Width = 8235
        txt_RUC.Text = "" & Trim(wcodruc_Nuevo)
       
        If gsw = "2" Then
            txtNro_Brevete.Visible = True
            Label11.Visible = True
        End If
    End If
 
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    
End Sub

Private Sub nuevo()
    
    limpiaForm Me
    chk_Cliente.Enabled = True
    chk_Cliente.Value = 1
    txtCod_Pais.Text = "02001"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            nuevo
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
            indCopiaDireccion = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            fraGeneral.Enabled = True
            indCopiaDireccion = False
        Case 5 'Imprimir
            ShellEx App.Path & "\Temporales\Mantenimiento_Personas_Rapido.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 7 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub habilitaBotones(indexBoton As Integer)
Dim indHabilitar As Boolean

    Select Case indexBoton
        Case 1, 2, 3 'Nuevo, Grabar, Modificar
            If indexBoton = 2 Then indHabilitar = True
            Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
            Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
            Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
            Toolbar1.Buttons(5).Visible = indHabilitar 'Imprimir
        Case 4, 6 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = True
            Toolbar1.Buttons(6).Visible = False
    End Select

End Sub

Private Sub txtCod_Depa_Change()
    
    txtGls_Depa.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", txtCod_Depa.Text, False, " idProv = '00'")
    txtCod_Prov.Text = ""
    txtGls_Prov.Text = ""
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""

End Sub

Private Sub txtCod_Depa_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DEPARTAMENTO", txtCod_Depa, txtGls_Depa
        KeyAscii = 0
        If txtCod_Depa.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Distrito_Change()
    
    txtGls_Distrito.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", txtCod_Distrito.Text, False)
    If indCopiaDireccion Then
        indEditando = True
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & " " & txtGls_Distrito.Text & " " & txtGls_Prov.Text
        indEditando = False
    End If

End Sub

Private Sub txtCod_Distrito_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DISTRITO", txtCod_Distrito, txtGls_Distrito, "AND idDpto = '" & txtCod_Depa.Text & "' and idProv = '" + txtCod_Prov.Text + "'"
        KeyAscii = 0
        If txtCod_Distrito.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Pais_Change()
    
    txtGls_Pais.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_Pais.Text, False)
    txtCod_Depa.Text = ""
    txtGls_Depa.Text = ""
    txtCod_Prov.Text = ""
    txtGls_Prov.Text = ""
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""

End Sub

Private Sub txtCod_Pais_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PAIS", txtCod_Pais, txtGls_Pais
        KeyAscii = 0
        If txtCod_Pais.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Prov_Change()
    
    txtGls_Prov.Text = traerCampo("Ubigeo", "GlsUbigeo", "idProv", txtCod_Prov.Text, False, " idDpto = '" & txtCod_Depa.Text & "' and idProv <> '00' and idDist = '00'")
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""
    If indCopiaDireccion Then
        indEditando = True
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & " " & txtGls_Distrito.Text & " " & txtGls_Prov.Text
        indEditando = False
    End If

End Sub

Private Sub txtCod_prov_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PROVINCIA", txtCod_Prov, txtGls_Prov, "AND idDpto = '" & txtCod_Depa.Text + "'"
        KeyAscii = 0
        If txtCod_Prov.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_TipoPersona_Change()
    
    txtGls_TipoPersona.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_TipoPersona.Text, False)

End Sub

Private Sub txtCod_TipoPersona_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "TIPOPERSONA", txtCod_TipoPersona, txtGls_TipoPersona
        KeyAscii = 0
        If txtCod_TipoPersona.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_TipoPersona_LostFocus()

    If txtCod_TipoPersona.Text = "01001" Then  'natural
        lblApePaterno.Caption = "Apellido Paterno:"
        lblApeMaterno.Enabled = True
        lblNombres.Enabled = True
        txtGls_Materno.Enabled = True
        txtGls_Nombres.Enabled = True
        txtGls_Materno.BackColor = &H80000005
        txtGls_Nombres.BackColor = &H80000005
        txtGls_Materno.Vacio = False
        txtGls_Nombres.Vacio = False
        txt_RUC.Vacio = True
        lblFecNacimiento.Enabled = True
        dtpFechaNac.Enabled = True
        
    Else 'juridica
        lblApePaterno.Caption = "Razon Social:"
        lblApeMaterno.Enabled = False
        lblNombres.Enabled = False
        txtGls_Materno.Enabled = False
        txtGls_Nombres.Enabled = False
        txtGls_Materno.BackColor = &H80000018
        txtGls_Nombres.BackColor = &H80000018
        txtGls_Materno.Vacio = True
        txtGls_Nombres.Vacio = True
        txt_RUC.Vacio = False
        lblFecNacimiento.Enabled = False
        dtpFechaNac.Enabled = False
    End If

End Sub

Private Sub txtGls_Direccion_Change()

    If indCopiaDireccion Then
        indEditando = True
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & " " & txtGls_Distrito.Text & " " & txtGls_Prov.Text
        indEditando = False
    End If

End Sub

Private Sub txtGls_DireccionFiscal_Change()
    
    If indEditando = False Then indCopiaDireccion = False

End Sub

Private Sub txtGls_Nombres_Change()

    txtGlsPersona.Text = Trim(Trim(txtGls_Paterno.Text) & " " & Trim(txtGls_Materno.Text) & " " & Trim(txtGls_Nombres.Text))

End Sub

Private Sub txtGls_Paterno_Change()
    
    txtGlsPersona.Text = Trim(Trim(txtGls_Paterno.Text) & " " & Trim(txtGls_Materno.Text) & " " & Trim(txtGls_Nombres.Text))

End Sub

Private Sub txtGls_Materno_Change()
    
    txtGlsPersona.Text = Trim(Trim(txtGls_Paterno.Text) & " " & Trim(txtGls_Materno.Text) & " " & Trim(txtGls_Nombres.Text))

End Sub

Private Sub mostrarPersona(strCodPer As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT idPersona,GlsPersona,apellidoPaterno,apellidoMaterno,nombres," & _
            "tipoPersona , ruc, direccion, p.iddistrito, u.idDpto, u.idProv, telefonos, mail, FechaNacimiento, direccionEntrega " & _
            "FROM personas p,ubigeo u " & _
            "WHERE p.iddistrito = u.iddistrito and idpersona = '" & strCodPer & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    txtCod_Depa.Tag = "TidDpto"
    txtCod_Prov.Tag = "TidProv"
    txtCod_Pais.Text = "02001"
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If traerCampo("clientes", "idCliente", "idCliente", txtCod_Persona.Text, True) = "" Then
        chk_Cliente.Value = 0
        chk_Cliente.Enabled = True
    Else
        chk_Cliente.Value = 1
        chk_Cliente.Enabled = False
    End If
    txtCod_Depa.Tag = ""
    txtCod_Prov.Tag = ""
    indCopiaDireccion = False
    Me.Refresh

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub actualizaCliente(ByRef StrMsgError As String)
On Error GoTo Err
Dim cdirectorio As String
    
    If chk <> "" Then
        chk_Cliente.Value = "0"
    End If
    
    If chk_Cliente.Value Then
        If traerCampo("proveedores", "idproveedor", "idproveedor", txtCod_Persona.Text, True) = "" Then
           csql = "INSERT INTO proveedores(idproveedor,idEmpresa,idgrupoproveedor) VALUES('" & txtCod_Persona.Text & "','" & glsEmpresa & "','" & txtcod_obliga.Text & "')"
           Cn.Execute csql
        End If
        
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Public Sub MostrarForm(ByRef codCliente As String, Optional CHKCLIENTE As String, Optional sw As String, Optional nomfrm As String)

    If CHKCLIENTE <> "" Then
        chk = CHKCLIENTE
    Else
        chk = ""
    End If
    
    If sw <> "" Then
        gsw = sw
    Else
        gsw = ""
    End If
    
    nuevo
    fraGeneral.Visible = True
    fraGeneral.Enabled = True
    habilitaBotones 1
    indCopiaDireccion = True

    Load Me
    Me.Show 1

    codCliente = strCodCliente
    Unload Me

End Sub
