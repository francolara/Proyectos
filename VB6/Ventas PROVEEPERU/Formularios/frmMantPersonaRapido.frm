VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmMantPersonaRapido 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Clientes"
   ClientHeight    =   8430
   ClientLeft      =   10770
   ClientTop       =   1005
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   8145
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   240
      Top             =   1110
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
            Picture         =   "frmMantPersonaRapido.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersonaRapido.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
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
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7620
      Left            =   45
      TabIndex        =   14
      Top             =   660
      Width           =   8040
      Begin VB.CommandButton CmbTipoDocIdentidad 
         Height          =   315
         Left            =   7500
         Picture         =   "frmMantPersonaRapido.frx":3518
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   4275
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaFormasPago 
         Height          =   315
         Left            =   6165
         Picture         =   "frmMantPersonaRapido.frx":38A2
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   6795
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaVendedorCampo 
         Height          =   315
         Left            =   6180
         Picture         =   "frmMantPersonaRapido.frx":3C2C
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   6390
         Width           =   390
      End
      Begin VB.CheckBox chk_Cliente 
         Appearance      =   0  'Flat
         Caption         =   "Cliente"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1620
         TabIndex        =   47
         Top             =   7245
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmbAyudaDistrito 
         Height          =   315
         Left            =   6195
         Picture         =   "frmMantPersonaRapido.frx":3FB6
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3420
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaProv 
         Height          =   315
         Left            =   6195
         Picture         =   "frmMantPersonaRapido.frx":4340
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3045
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaDepa 
         Height          =   315
         Left            =   6195
         Picture         =   "frmMantPersonaRapido.frx":46CA
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2670
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaPais 
         Height          =   315
         Left            =   6195
         Picture         =   "frmMantPersonaRapido.frx":4A54
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2295
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaTipoPersona 
         Height          =   315
         Left            =   6240
         Picture         =   "frmMantPersonaRapido.frx":4DDE
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   660
         Width           =   390
      End
      Begin MSComCtl2.DTPicker dtpFechaNac 
         Height          =   315
         Left            =   1650
         TabIndex        =   12
         Tag             =   "FFechaNacimiento"
         Top             =   5610
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107741185
         CurrentDate     =   38638
      End
      Begin CATControls.CATTextBox txtCod_Persona 
         Height          =   285
         Left            =   6975
         TabIndex        =   15
         Tag             =   "TidPersona"
         Top             =   150
         Width           =   915
         _ExtentX        =   1614
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
         Container       =   "frmMantPersonaRapido.frx":5168
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Paterno 
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Tag             =   "TapellidoPaterno"
         Top             =   1110
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   556
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
         MaxLength       =   80
         Container       =   "frmMantPersonaRapido.frx":5184
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Materno 
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Tag             =   "TapellidoMaterno"
         Top             =   1500
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   556
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
         MaxLength       =   80
         Container       =   "frmMantPersonaRapido.frx":51A0
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nombres 
         Height          =   315
         Left            =   1650
         TabIndex        =   3
         Tag             =   "Tnombres"
         Top             =   1875
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   556
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
         MaxLength       =   80
         Container       =   "frmMantPersonaRapido.frx":51BC
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Direccion 
         Height          =   315
         Left            =   1650
         TabIndex        =   8
         Tag             =   "Tdireccion"
         Top             =   3840
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   556
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
         MaxLength       =   255
         Container       =   "frmMantPersonaRapido.frx":51D8
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_RUC 
         Height          =   315
         Left            =   1650
         TabIndex        =   9
         Tag             =   "Truc"
         Top             =   4770
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
         MaxLength       =   11
         Container       =   "frmMantPersonaRapido.frx":51F4
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Telefono 
         Height          =   315
         Left            =   5925
         TabIndex        =   10
         Tag             =   "TTelefonos"
         Top             =   4770
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
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
         MaxLength       =   85
         Container       =   "frmMantPersonaRapido.frx":5210
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Email 
         Height          =   315
         Left            =   1650
         TabIndex        =   13
         Tag             =   "Tmail"
         Top             =   6030
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   556
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
         MaxLength       =   80
         Container       =   "frmMantPersonaRapido.frx":522C
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_TipoPersona 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Tag             =   "TtipoPersona"
         Top             =   690
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":5248
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TipoPersona 
         Height          =   315
         Left            =   2670
         TabIndex        =   31
         Top             =   690
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":5264
      End
      Begin CATControls.CATTextBox txtCod_Pais 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Tag             =   "TIdPais"
         Top             =   2280
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":5280
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Pais 
         Height          =   315
         Left            =   2625
         TabIndex        =   33
         Top             =   2280
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":529C
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Depa 
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Top             =   2655
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":52B8
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Depa 
         Height          =   315
         Left            =   2625
         TabIndex        =   35
         Top             =   2655
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":52D4
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Prov 
         Height          =   315
         Left            =   1650
         TabIndex        =   6
         Top             =   3030
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":52F0
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Prov 
         Height          =   315
         Left            =   2625
         TabIndex        =   37
         Top             =   3030
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":530C
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Distrito 
         Height          =   315
         Left            =   1650
         TabIndex        =   7
         Tag             =   "TidDistrito"
         Top             =   3390
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":5328
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Distrito 
         Height          =   315
         Left            =   2625
         TabIndex        =   39
         Top             =   3405
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":5344
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtGlsPersona 
         Height          =   285
         Left            =   750
         TabIndex        =   41
         Tag             =   "TGlsPersona"
         Top             =   210
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   16711680
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
         Container       =   "frmMantPersonaRapido.frx":5360
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_DireccionFiscal 
         Height          =   315
         Left            =   1650
         TabIndex        =   11
         Tag             =   "TdireccionEntrega"
         Top             =   5235
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   556
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
         MaxLength       =   255
         Container       =   "frmMantPersonaRapido.frx":537C
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_VendedorCampo 
         Height          =   315
         Left            =   1635
         TabIndex        =   50
         Top             =   6420
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":5398
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_VendedorCampo 
         Height          =   315
         Left            =   2595
         TabIndex        =   51
         Top             =   6420
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
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
         Container       =   "frmMantPersonaRapido.frx":53B4
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_FormasPago 
         Height          =   285
         Left            =   1635
         TabIndex        =   54
         Top             =   6795
         Width           =   915
         _ExtentX        =   1614
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
         Container       =   "frmMantPersonaRapido.frx":53D0
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_FormasPago 
         Height          =   285
         Left            =   2595
         TabIndex        =   55
         Top             =   6795
         Width           =   3555
         _ExtentX        =   6271
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
         Container       =   "frmMantPersonaRapido.frx":53EC
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox TxtCodTipoDocIdentidad 
         Height          =   315
         Left            =   1650
         TabIndex        =   58
         Tag             =   "TIdTipoDocIdentidad"
         Top             =   4290
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
         Container       =   "frmMantPersonaRapido.frx":5408
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsTipoDocIdentidad 
         Height          =   315
         Left            =   2625
         TabIndex        =   59
         Top             =   4290
         Width           =   4845
         _ExtentX        =   8546
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
         Container       =   "frmMantPersonaRapido.frx":5424
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Documento de Identidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         TabIndex        =   60
         Top             =   4275
         Width           =   1425
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         Caption         =   "Forma de Pago:"
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
         Left            =   135
         TabIndex        =   56
         Top             =   6840
         Width           =   1485
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor Campo:"
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
         TabIndex        =   52
         Top             =   6480
         Width           =   1305
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Direc. Entrega:"
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
         Left            =   150
         TabIndex        =   48
         Top             =   5235
         Width           =   1065
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Distrito:"
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
         Left            =   150
         TabIndex        =   40
         Top             =   3450
         Width           =   540
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6075
         TabIndex        =   28
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Entidad:"
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
         Left            =   150
         TabIndex        =   27
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label lblNombres 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nombres:"
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
         Left            =   150
         TabIndex        =   26
         Top             =   1950
         Width           =   690
      End
      Begin VB.Label lblApeMaterno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno:"
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
         Left            =   150
         TabIndex        =   25
         Top             =   1575
         Width           =   1245
      End
      Begin VB.Label lblApePaterno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno:"
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
         Left            =   150
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
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
         Left            =   150
         TabIndex        =   23
         Top             =   2250
         Width           =   345
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Departamento:"
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
         Left            =   150
         TabIndex        =   22
         Top             =   2625
         Width           =   1050
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Provincia:"
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
         Left            =   150
         TabIndex        =   21
         Top             =   3000
         Width           =   705
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
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
         Left            =   150
         TabIndex        =   20
         Top             =   3870
         Width           =   720
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "E Mail:"
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
         Left            =   150
         TabIndex        =   19
         Top             =   6105
         Width           =   450
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Telefonos:"
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
         Left            =   4575
         TabIndex        =   18
         Top             =   4845
         Width           =   765
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "R.U.C.:"
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
         Left            =   150
         TabIndex        =   17
         Top             =   4845
         Width           =   495
      End
      Begin VB.Label lblFecNacimiento 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha Nac.:"
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
         Left            =   150
         TabIndex        =   16
         Top             =   5685
         Width           =   870
      End
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      Caption         =   "Listado"
      ForeColor       =   &H00C00000&
      Height          =   7305
      Left            =   45
      TabIndex        =   42
      Top             =   675
      Width           =   8040
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   60
         TabIndex        =   44
         Top             =   240
         Width           =   7815
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   285
            Left            =   1500
            TabIndex        =   45
            Top             =   210
            Width           =   6240
            _ExtentX        =   11007
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
            MaxLength       =   255
            Container       =   "frmMantPersonaRapido.frx":5440
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Busqueda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   210
            Width           =   765
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   5985
         Left            =   90
         OleObjectBlob   =   "frmMantPersonaRapido.frx":545C
         TabIndex        =   43
         Top             =   1110
         Width           =   7860
      End
   End
End
Attribute VB_Name = "frmMantPersonaRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strCodCliente As String
Private indCopiaDireccion As Boolean
Private indEditando As Boolean

Private Sub Grabar(ByRef StrMsgError As String)
Dim StrCodigo As String
Dim strMsg      As String
On Error GoTo Err

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
    
    EjecutaSQLForm Me, 0, False, "personas", StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    actualizaCliente
    
    strMsg = "Grabo"
    
    strCodCliente = txtCod_Persona.Text
Else 'modifica

    EjecutaSQLForm Me, 1, False, "personas", StrMsgError, "idpersona"
    If StrMsgError <> "" Then GoTo Err
    
    actualizaCliente
    
    strMsg = "Modifico"
End If

Me.Hide
'MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title

FraGeneral.Enabled = False

listaPersona StrMsgError
If StrMsgError <> "" Then GoTo Err
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

Private Sub cmbAyudaFormasPago_Click()
    mostrarAyuda "FORMASPAGO", txtCod_FormasPago, txtGls_FormasPago
    If txtCod_FormasPago.Text <> "" Then SendKeys "{tab}"
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
    mostrarAyuda "VENDEDOR", txtCod_VendedorCampo, txtGls_VendedorCampo
End Sub

Private Sub CmbTipoDocIdentidad_Click()

    mostrarAyuda "TIPODOCUMENTOIDENTIDAD", TxtCodTipoDocIdentidad, TxtGlsTipoDocIdentidad
    
End Sub

Private Sub dtpFechaNac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

ConfGrid GLista, False, False, False, False

'listaPersona strMsgError
'If strMsgError <> "" Then GoTo err

'fraListado.Visible = True
'fraGeneral.Visible = False
'habilitaBotones 6
'nuevo
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub nuevo()
    limpiaForm Me
    chk_Cliente.Enabled = True
    chk_Cliente.Value = 1
End Sub

Private Sub gLista_OnDblClick()
Dim StrMsgError As String
On Error GoTo Err
mostrarPersona GLista.Columns.ColumnByName("idPersona").Value, StrMsgError
If StrMsgError <> "" Then GoTo Err

txtCod_TipoPersona_LostFocus

FraListado.Visible = False
FraGeneral.Visible = True
FraGeneral.Enabled = False
habilitaBotones 2
Exit Sub
Err:
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
On Error GoTo Err
Select Case Button.Index
    Case 1 'Nuevo
        nuevo
        FraListado.Visible = False
        FraGeneral.Visible = True
        FraGeneral.Enabled = True
        indCopiaDireccion = True
    Case 2 'Grabar
        Grabar StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Case 3 'Modificar
        FraGeneral.Enabled = True
        indCopiaDireccion = False
    Case 4, 6 'Cancelar
        FraListado.Visible = True
        FraGeneral.Visible = False
        FraGeneral.Enabled = False
    Case 5 'Imprimir
    
'    Case 6 'Lista
'        fraListado.Visible = True
'        fraGeneral.Visible = False
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
        Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
        Toolbar1.Buttons(5).Visible = indHabilitar 'Imprimir
        Toolbar1.Buttons(6).Visible = indHabilitar 'Lista
    Case 4, 6 'Cancelar, Lista
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(3).Visible = False
        Toolbar1.Buttons(4).Visible = False
        Toolbar1.Buttons(5).Visible = False
        Toolbar1.Buttons(6).Visible = False
End Select

End Sub

Private Sub txt_TextoBuscar_Change()
Dim StrMsgError As String
On Error GoTo Err
listaPersona StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then GLista.SetFocus
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

Private Sub txtCod_FormasPago_Change()
txtGls_FormasPago.Text = traerCampo("formaspagos", "GlsFormaPago", "idFormaPago", txtCod_FormasPago.Text, True)
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

Private Sub txtCod_Prov_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PROVINCIA", txtCod_Prov, txtGls_Prov, "AND idDpto = '" & txtCod_Depa.Text + "'"
        KeyAscii = 0
        If txtCod_Prov.Text <> "" Then SendKeys "{tab}"
        End If
End Sub

Private Sub txtCod_TipoPersona_Change()
    
    txtGls_TipoPersona.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_TipoPersona.Text, False)
    
    If txtCod_TipoPersona.Text = "01001" Then
        
        TxtCodTipoDocIdentidad.Text = "1"
    
    Else
        
        TxtCodTipoDocIdentidad.Text = "6"
    
    End If
    
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

Private Sub txtCod_VendedorCampo_Change()
    txtGls_VendedorCampo.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_VendedorCampo.Text, False)
End Sub

Private Sub TxtCodTipoDocIdentidad_Change()
On Error GoTo Err
Dim StrMsgError                 As String

    If TxtCodTipoDocIdentidad.Text <> "" Then
        
        TxtGlsTipoDocIdentidad.Text = traerCampo("TiposDocIdentidad", "GlsTipoDocIdentidad", "IdTipoDocIdentidad", TxtCodTipoDocIdentidad.Text, False)
        txt_RUC.MaxLength = Val("" & traerCampo("TiposDocIdentidad", "Longitud", "IdTipoDocIdentidad", TxtCodTipoDocIdentidad.Text, False))
    
    Else
        
        TxtGlsTipoDocIdentidad.Text = ""
        txt_RUC.MaxLength = 11
        
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
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


Private Sub listaPersona(ByRef StrMsgError As String)
Dim strCond As String
On Error GoTo Err
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND GlsPersona LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT idPersona ,GlsPersona ," & _
                  "if(tipoPersona = '01001','Natural','Juridica') as TipoPersona,ruc,concat(direccion,', ',GlsUbigeo) as Direccion " & _
           "FROM personas p,ubigeo u " & _
           "WHERE p.iddistrito = u.iddistrito " & strCond & _
           "ORDER BY idPersona"
    With GLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn  ''Cn
        
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idPersona"
End With
Me.Refresh
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub



Private Sub mostrarPersona(strCodPer As String, ByRef StrMsgError As String)
Dim rst As New ADODB.Recordset
On Error GoTo Err
    csql = "SELECT idPersona,GlsPersona,apellidoPaterno,apellidoMaterno,nombres," & _
                   "tipoPersona , ruc, direccion, p.iddistrito, u.idDpto, u.idProv, telefonos, mail, FechaNacimiento, direccionEntrega " & _
           "FROM personas p,ubigeo u " & _
           "WHERE p.iddistrito = u.iddistrito and idpersona = '" & strCodPer & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    txtCod_Depa.Tag = "TidDpto"
    txtCod_Prov.Tag = "TidProv"
    txtCod_Pais.Tag = "TidPais"
    
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


Private Sub actualizaCliente()

If chk_Cliente.Value Then
    If traerCampo("clientes", "idCliente", "idCliente", txtCod_Persona.Text, True) = "" Then
       csql = "INSERT INTO clientes(idCliente,idEmpresa,idVendedorCampo) VALUES('" & txtCod_Persona.Text & "','" & glsEmpresa & "','" & txtCod_VendedorCampo.Text & "')"
       Cn.Execute csql
       
       csql = "INSERT INTO clientesformapagos(idempresa, idcliente, idFormaPago, FecRegistro, idUsuario, indEstado) VALUES('" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & txtCod_FormasPago.Text & "','" & Format(Trim("" & getFechaHoraSistema), "yyyy-mm-dd hh:mm:ss") & "','" & glsUser & "',1)"
       Cn.Execute csql
    End If
Else 'No Cliente
    
End If
End Sub

Public Sub MostrarForm(ByRef codCliente As String)

    nuevo
    FraListado.Visible = False
    FraGeneral.Visible = True
    FraGeneral.Enabled = True
    
    habilitaBotones 1
    
    indCopiaDireccion = True

    Load Me
    Me.Show 1

    codCliente = strCodCliente

    Unload Me

End Sub

