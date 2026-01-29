VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmMantPersona_Consul 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Entidades"
   ClientHeight    =   9435
   ClientLeft      =   1590
   ClientTop       =   855
   ClientWidth     =   11250
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
   Picture         =   "frmMantPersona_Consul.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   11250
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   -45
      Top             =   1665
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
            Picture         =   "frmMantPersona_Consul.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":06DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":0B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":0EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":1262
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":15FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":1996
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":1D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":20CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":2464
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":27FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPersona_Consul.frx":34C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   1164
      ButtonWidth     =   1931
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "      Nuevo      "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Caption         =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "       Lista       "
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8760
      Left            =   45
      TabIndex        =   17
      Top             =   630
      Width           =   11145
      Begin VB.CheckBox chk_Proveedor 
         Appearance      =   0  'Flat
         Caption         =   "Proveedores"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7650
         TabIndex        =   61
         Top             =   7020
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chk_Cliente 
         Appearance      =   0  'Flat
         Caption         =   "Cliente"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3600
         TabIndex        =   60
         Top             =   7020
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.Frame FraEmpleado 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   4400
         TabIndex        =   59
         Top             =   8010
         Visible         =   0   'False
         Width           =   2850
         Begin VB.OptionButton OptEmpleadoTipo 
            Caption         =   "Personal"
            Height          =   210
            Index           =   0
            Left            =   270
            TabIndex        =   66
            Top             =   270
            Width           =   1050
         End
         Begin VB.OptionButton OptEmpleadoTipo 
            Caption         =   "Socio"
            Height          =   210
            Index           =   1
            Left            =   1530
            TabIndex        =   65
            Top             =   270
            Value           =   -1  'True
            Width           =   1050
         End
      End
      Begin VB.CheckBox ChkPersonal 
         Appearance      =   0  'Flat
         Caption         =   "Empleado"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2745
         TabIndex        =   58
         Top             =   8190
         Width           =   1305
      End
      Begin VB.CheckBox ChkRelacionada 
         Appearance      =   0  'Flat
         Caption         =   "Relacionada"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2745
         TabIndex        =   57
         Top             =   7560
         Width           =   1305
      End
      Begin VB.CommandButton CmbTipoDocIdentidad 
         Enabled         =   0   'False
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
         Left            =   10515
         Picture         =   "frmMantPersona_Consul.frx":385A
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4185
         Width           =   390
      End
      Begin VB.Frame FraRelacionada 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   4395
         TabIndex        =   50
         Top             =   7380
         Visible         =   0   'False
         Width           =   4335
         Begin VB.OptionButton OptRelacionadaTipo 
            Caption         =   "Matriz"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   64
            Top             =   225
            Width           =   915
         End
         Begin VB.OptionButton OptRelacionadaTipo 
            Caption         =   "SubSidiaria"
            Height          =   210
            Index           =   1
            Left            =   1530
            TabIndex        =   63
            Top             =   225
            Width           =   1185
         End
         Begin VB.OptionButton OptRelacionadaTipo 
            Caption         =   "Asociada"
            Height          =   210
            Index           =   2
            Left            =   3015
            TabIndex        =   62
            Top             =   225
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin VB.CommandButton cmbAyudaDistrito 
         Enabled         =   0   'False
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
         Left            =   10515
         Picture         =   "frmMantPersona_Consul.frx":3BE4
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3420
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaProv 
         Enabled         =   0   'False
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
         Left            =   10515
         Picture         =   "frmMantPersona_Consul.frx":3F6E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3045
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaDepa 
         Enabled         =   0   'False
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
         Left            =   10515
         Picture         =   "frmMantPersona_Consul.frx":42F8
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2670
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaPais 
         Enabled         =   0   'False
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
         Left            =   10515
         Picture         =   "frmMantPersona_Consul.frx":4682
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2295
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaTipoPersona 
         Enabled         =   0   'False
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
         Left            =   10515
         Picture         =   "frmMantPersona_Consul.frx":4A0C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   750
         Width           =   390
      End
      Begin MSComCtl2.DTPicker dtpFechaNac 
         Height          =   315
         Left            =   2775
         TabIndex        =   15
         Tag             =   "FFechaNacimiento"
         Top             =   5805
         Width           =   1245
         _ExtentX        =   2196
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
         Format          =   109182977
         CurrentDate     =   38638
      End
      Begin CATControls.CATTextBox txtCod_Persona 
         Height          =   315
         Left            =   9495
         TabIndex        =   18
         Tag             =   "TidPersona"
         Top             =   330
         Width           =   960
         _ExtentX        =   1693
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
         Container       =   "frmMantPersona_Consul.frx":4D96
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Paterno 
         Height          =   315
         Left            =   2775
         TabIndex        =   3
         Tag             =   "TapellidoPaterno"
         Top             =   1110
         Width           =   7725
         _ExtentX        =   13626
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
         Container       =   "frmMantPersona_Consul.frx":4DB2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Materno 
         Height          =   315
         Left            =   2775
         TabIndex        =   4
         Tag             =   "TapellidoMaterno"
         Top             =   1500
         Width           =   7725
         _ExtentX        =   13626
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
         Container       =   "frmMantPersona_Consul.frx":4DCE
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nombres 
         Height          =   315
         Left            =   2775
         TabIndex        =   5
         Tag             =   "Tnombres"
         Top             =   1875
         Width           =   7725
         _ExtentX        =   13626
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
         Container       =   "frmMantPersona_Consul.frx":4DEA
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Direccion 
         Height          =   315
         Left            =   2775
         TabIndex        =   10
         Tag             =   "Tdireccion"
         Top             =   3795
         Width           =   7725
         _ExtentX        =   13626
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
         Container       =   "frmMantPersona_Consul.frx":4E06
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_RUC 
         Height          =   315
         Left            =   2775
         TabIndex        =   11
         Tag             =   "Truc"
         Top             =   4590
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
         Container       =   "frmMantPersona_Consul.frx":4E22
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Telefono 
         Height          =   315
         Left            =   6465
         TabIndex        =   12
         Tag             =   "TTelefonos"
         Top             =   4590
         Width           =   4035
         _ExtentX        =   7117
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
         Container       =   "frmMantPersona_Consul.frx":4E3E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Email 
         Height          =   315
         Left            =   2775
         TabIndex        =   16
         Tag             =   "Tmail"
         Top             =   6210
         Width           =   7725
         _ExtentX        =   13626
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
         Container       =   "frmMantPersona_Consul.frx":4E5A
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_TipoPersona 
         Height          =   315
         Left            =   2775
         TabIndex        =   2
         Tag             =   "TtipoPersona"
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
         Container       =   "frmMantPersona_Consul.frx":4E76
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TipoPersona 
         Height          =   315
         Left            =   3735
         TabIndex        =   34
         Top             =   735
         Width           =   6735
         _ExtentX        =   11880
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
         Container       =   "frmMantPersona_Consul.frx":4E92
      End
      Begin CATControls.CATTextBox txtCod_Pais 
         Height          =   315
         Left            =   2775
         TabIndex        =   6
         Tag             =   "TidPais"
         Top             =   2280
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
         Container       =   "frmMantPersona_Consul.frx":4EAE
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Pais 
         Height          =   315
         Left            =   3750
         TabIndex        =   36
         Top             =   2280
         Width           =   6735
         _ExtentX        =   11880
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
         Container       =   "frmMantPersona_Consul.frx":4ECA
      End
      Begin CATControls.CATTextBox txtCod_Depa 
         Height          =   315
         Left            =   2775
         TabIndex        =   7
         Top             =   2655
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
         Container       =   "frmMantPersona_Consul.frx":4EE6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Depa 
         Height          =   315
         Left            =   3750
         TabIndex        =   38
         Top             =   2655
         Width           =   6735
         _ExtentX        =   11880
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
         Container       =   "frmMantPersona_Consul.frx":4F02
      End
      Begin CATControls.CATTextBox txtCod_Prov 
         Height          =   315
         Left            =   2775
         TabIndex        =   8
         Top             =   3030
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
         Container       =   "frmMantPersona_Consul.frx":4F1E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Prov 
         Height          =   315
         Left            =   3750
         TabIndex        =   40
         Top             =   3030
         Width           =   6735
         _ExtentX        =   11880
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
         Container       =   "frmMantPersona_Consul.frx":4F3A
      End
      Begin CATControls.CATTextBox txtCod_Distrito 
         Height          =   315
         Left            =   2775
         TabIndex        =   9
         Tag             =   "TidDistrito"
         Top             =   3405
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
         Container       =   "frmMantPersona_Consul.frx":4F56
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Distrito 
         Height          =   315
         Left            =   3750
         TabIndex        =   42
         Top             =   3405
         Width           =   6735
         _ExtentX        =   11880
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
         Container       =   "frmMantPersona_Consul.frx":4F72
      End
      Begin CATControls.CATTextBox txtGlsPersona 
         Height          =   285
         Left            =   750
         TabIndex        =   44
         Tag             =   "TGlsPersona"
         Top             =   210
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   12640511
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
         Container       =   "frmMantPersona_Consul.frx":4F8E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_DireccionFiscal 
         Height          =   315
         Left            =   2775
         TabIndex        =   13
         Tag             =   "TdireccionEntrega"
         Top             =   4995
         Width           =   7725
         _ExtentX        =   13626
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
         Container       =   "frmMantPersona_Consul.frx":4FAA
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Contacto 
         Height          =   315
         Left            =   2775
         TabIndex        =   14
         Tag             =   "TGlsContacto"
         Top             =   5400
         Width           =   7725
         _ExtentX        =   13626
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
         Container       =   "frmMantPersona_Consul.frx":4FC6
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtCodTipoDocIdentidad 
         Height          =   315
         Left            =   2775
         TabIndex        =   52
         Tag             =   "TIdTipoDocIdentidad"
         Top             =   4200
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
         Container       =   "frmMantPersona_Consul.frx":4FE2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsTipoDocIdentidad 
         Height          =   315
         Left            =   3750
         TabIndex        =   53
         Top             =   4200
         Width           =   6735
         _ExtentX        =   11880
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
         Container       =   "frmMantPersona_Consul.frx":4FFE
      End
      Begin CATControls.CATTextBox TxtGlsNombreComercial 
         Height          =   315
         Left            =   2775
         TabIndex        =   55
         Tag             =   "TGlsNombreComercial"
         Top             =   6615
         Width           =   7725
         _ExtentX        =   13626
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
         Container       =   "frmMantPersona_Consul.frx":501A
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nombre Comercial"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   56
         Top             =   6690
         Width           =   1305
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Documento de Identidad"
         Height          =   420
         Left            =   765
         TabIndex        =   54
         Top             =   4185
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   49
         Top             =   5460
         Width           =   645
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dirección de Entrega"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   48
         Top             =   5055
         Width           =   1500
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   43
         Top             =   3450
         Width           =   495
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   8865
         TabIndex        =   31
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Entidad"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   30
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label lblNombres 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   29
         Top             =   1950
         Width           =   645
      End
      Begin VB.Label lblApeMaterno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   28
         Top             =   1575
         Width           =   1200
      End
      Begin VB.Label lblApePaterno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   27
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "País"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   26
         Top             =   2250
         Width           =   300
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   25
         Top             =   2625
         Width           =   1005
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   24
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   23
         Top             =   3900
         Width           =   675
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "E Mail"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   22
         Top             =   6285
         Width           =   405
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5610
         TabIndex        =   21
         Top             =   4665
         Width           =   720
      End
      Begin VB.Label lblRUC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Número de Documento"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   20
         Top             =   4665
         Width           =   1635
      End
      Begin VB.Label lblFecNacimiento 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Nacimiento"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   19
         Top             =   5880
         Width           =   1500
      End
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8745
      Left            =   45
      TabIndex        =   45
      Top             =   630
      Width           =   11145
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   150
         TabIndex        =   46
         Top             =   150
         Width           =   10875
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   990
            TabIndex        =   0
            Top             =   255
            Width           =   9750
            _ExtentX        =   17198
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
            Container       =   "frmMantPersona_Consul.frx":5036
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Búsqueda"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   165
            TabIndex        =   47
            Top             =   315
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   7605
         Left            =   135
         OleObjectBlob   =   "frmMantPersona_Consul.frx":5052
         TabIndex        =   1
         Top             =   1020
         Width           =   10920
      End
   End
End
Attribute VB_Name = "frmMantPersona_Consul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private indCopiaDireccion               As Boolean
Private indEditando                     As Boolean
Dim CIdContacto                         As String

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim StrCodigo                           As String
Dim strMsg                              As String
Dim CSqlC                               As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaHomonimia "personas", "GlsPersona", "idPersona", txtGlsPersona.Text, txtCod_Persona.Text, False, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If ChkRelacionada.Value = 1 And OptRelacionadaTipo(0).Value Then
    
        If Len(Trim("" & traerCampo("EmpresasRelacionadas", "IndMatriz", "IndMatriz", "1", True, "IdPersona <> '" & txtCod_Persona.Text & "'"))) > 0 Then
            
            StrMsgError = "Ya existe una Empresa Relacionada Matriz. Verificar."
        
        End If
    
    End If
    
    If Trim(txt_RUC.Text) <> "" Then
        validaHomonimia "personas", "ruc", "idPersona", txt_RUC.Text, txtCod_Persona.Text, False, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    If txtCod_Persona.Text = "" Then
        txtCod_Persona.Text = GeneraCorrelativoAnoMes("personas", "idPersona", False)
        EjecutaSQLForm Me, 0, False, "personas", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        actualizaCliente
        strMsg = "Grabo"
    
    Else 'modifica
        EjecutaSQLForm Me, 1, False, "personas", StrMsgError, "idpersona"
        If StrMsgError <> "" Then GoTo Err
        actualizaCliente
        strMsg = "Modifico"
    End If
    
    If Len(Trim("" & txtGls_Contacto.Text)) > 0 Then
        
        validaHomonimia "Personas", "GlsPersona", "IdPersona", txtGls_Contacto.Text, CIdContacto, False, StrMsgError
        If StrMsgError <> "" Then GoTo Err
                
        If Len(Trim("" & CIdContacto)) > 0 Then
            
            CSqlC = "Update Personas " & _
                    "Set GlsPersona = '" & txtGls_Contacto.Text & "' " & _
                    "Where IdPersona = '" & CIdContacto & "'"
            
            Cn.Execute CSqlC
            
        Else
        
            CIdContacto = GeneraCorrelativoAnoMes("Personas", "IdPersona", False)
            
            CSqlC = "Insert Into Personas(IdPersona,GlsPersona,ApellidoPaterno,ApellidoMaterno,Nombres,TipoPersona,Ruc,IdDistrito,Direccion," & _
                    "FechaNacimiento,Telefonos,Mail,DireccionEntrega,GlsContacto,IdPais,Linea_Credito)" & _
                    "Select '" & CIdContacto & "','" & txtGls_Contacto.Text & "','" & txtGls_Contacto.Text & "','" & txtGls_Contacto.Text & "'," & _
                    "'" & txtGls_Contacto.Text & "','01001','" & CIdContacto & "',IdDistrito,Direccion,FechaNacimiento,'','',DireccionEntrega,''," & _
                    "IdPais,Linea_Credito " & _
                    "From Personas " & _
                    "Where IdPersona = '" & txtCod_Persona.Text & "'"
            
            Cn.Execute CSqlC
            
        End If
    
        If chk_Cliente.Value = 1 Then
                
            CSqlC = "Delete From ContactosClientes " & _
                    "Where IdEmpresa = '" & glsEmpresa & "' And IdCliente = '" & txtCod_Persona.Text & "' And IdContacto = '" & CIdContacto & "'"
            
            Cn.Execute CSqlC
            
            CSqlC = "Insert Into ContactosClientes(IdEmpresa,IdCliente,IdContacto)Values(" & _
                    "'" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & CIdContacto & "')"
            
            Cn.Execute CSqlC
            
        End If
        
        If chk_Proveedor.Value = 1 Then
            
            CSqlC = "Delete From ContactosProveedores " & _
                    "Where IdEmpresa = '" & glsEmpresa & "' And IdProveedor = '" & txtCod_Persona.Text & "' And IdContacto = '" & CIdContacto & "'"
            
            Cn.Execute CSqlC
            
            CSqlC = "Insert Into ContactosProveedores(IdEmpresa,IdProveedor,IdContacto)Values(" & _
                    "'" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & CIdContacto & "')"
            
            Cn.Execute CSqlC
            
        End If
        
        CSqlC = "Update Personas " & _
                "Set IdContacto = '" & CIdContacto & "' " & _
                "Where IdPersona = '" & txtCod_Persona.Text & "'"
        
        Cn.Execute CSqlC
        
    End If
    
    CSqlC = "Delete From EmpresasRelacionadas " & _
            "Where IdEmpresa = '" & glsEmpresa & "' And IdPersona = '" & txtCod_Persona.Text & "'"

    Cn.Execute CSqlC
        
    If ChkRelacionada.Value = 1 Then
        
        CSqlC = "Insert Into EmpresasRelacionadas(IdEmpresa,IdPersona,IndMatriz,IndSubSidiaria,IndAsociada)Values(" & _
                "'" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & IIf(OptRelacionadaTipo(0).Value, "1", "0") & "'," & _
                "'" & IIf(OptRelacionadaTipo(1).Value, "1", "0") & "','" & IIf(OptRelacionadaTipo(2).Value, "1", "0") & "')"
                
        Cn.Execute CSqlC
        
    End If
    
    CSqlC = "Delete From Personal " & _
            "Where IdEmpresa = '" & glsEmpresa & "' And IdPersona = '" & txtCod_Persona.Text & "'"

    Cn.Execute CSqlC
        
    If ChkPersonal.Value = 1 Then
        
        CSqlC = "Insert Into Personal(IdEmpresa,IdPersona,IndPersonal,IndSocio)Values(" & _
                "'" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & IIf(OptEmpleadoTipo(0).Value, "1", "0") & "'," & _
                "'" & IIf(OptEmpleadoTipo(1).Value, "1", "0") & "')"
                
        Cn.Execute CSqlC
        
    End If
    
    MsgBox "Se " & strMsg & " satisfactoriamente", vbInformation, App.Title
    
    fraGeneral.Enabled = False
    listaPersona StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub ChkPersonal_Click()
On Error GoTo Err
Dim StrMsgError                             As String
    
    If ChkPersonal.Value = 1 Then
        
        ChkRelacionada.Value = 0
        FraEmpleado.Visible = True
        
    Else
        
        FraEmpleado.Visible = False
        
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ChkRelacionada_Click()
On Error GoTo Err
Dim StrMsgError                             As String
    
    If ChkRelacionada.Value = 1 Then
        
        ChkPersonal.Value = 0
        FraRelacionada.Visible = True
        
    Else
        
        FraRelacionada.Visible = False
        
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaDepa_Click()
    
    mostrarAyuda "DEPARTAMENTO", txtCod_Depa, txtGls_Depa, " AND idPais = '" & txtCod_Pais.Text & "'"
    If txtCod_Depa.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaDistrito_Click()
    
    mostrarAyuda "DISTRITO", txtCod_Distrito, txtGls_Distrito, "AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Depa.Text & "' and idProv = '" + txtCod_Prov.Text + "'"
    If txtCod_Distrito.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaPais_Click()
    
    mostrarAyuda "PAIS", txtCod_Pais, txtGls_Pais
    If txtCod_Pais.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaProv_Click()
    
    mostrarAyuda "PROVINCIA", txtCod_Prov, txtGls_Prov, "AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Depa.Text + "'"
    If txtCod_Prov.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaTipoPersona_Click()
    
    mostrarAyuda "TIPOPERSONA", txtCod_TipoPersona, txtGls_TipoPersona
    txtCod_TipoPersona_LostFocus
'    If txtCod_TipoPersona.Text <> "" Then SendKeys "{Tab}"

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
On Error GoTo Err
Dim StrMsgError As String
    
    CIdContacto = ""
    
    ConfGrid gLista, False, False, False, False
    fraListado.Visible = True
    fraGeneral.Visible = False
    
    If leeParametro("VISUALIZA_NOMBRE_COMERCIAL") = "S" Then
        
        gLista.Columns.ColumnByFieldName("GlsNombreComercial").Visible = True
    
    End If
    
    habilitaBotones 7
    
    nuevo
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub nuevo()
    
    limpiaForm Me
    
    CIdContacto = ""
    
    chk_Cliente.Enabled = True
    chk_Cliente.Value = 0
    
    chk_Proveedor.Enabled = True
    chk_Proveedor.Value = 0
    
    ChkRelacionada.Value = 0
    ChkPersonal.Value = 0
    
    txtCod_Pais.Text = "02001"

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarPersona gLista.Columns.ColumnByName("idPersona").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    txtCod_TipoPersona_LostFocus
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            nuevo
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
            indCopiaDireccion = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            fraGeneral.Enabled = True
            indCopiaDireccion = True
        Case 4, 7 'Cancelar
            If Button.Index = 7 Then
                listaPersona StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 6 'Imprimir
        Case 8 'Salir
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
            'If indexBoton = 2 Then indHabilitar = True
            'Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
            'Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
            'Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
            'Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
            'Toolbar1.Buttons(5).Visible = indHabilitar 'Eliminar
            'Toolbar1.Buttons(6).Visible = indHabilitar 'Imprimir
            Toolbar1.Buttons(7).Visible = True 'Lista
        Case 4, 7 'Cancelar, Lista
            'Toolbar1.Buttons(1).Visible = True
            'Toolbar1.Buttons(2).Visible = False
            'Toolbar1.Buttons(3).Visible = False
            'Toolbar1.Buttons(4).Visible = False
            'Toolbar1.Buttons(5).Visible = False
            'Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = True
    End Select

End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    If glsEnterAyudaClientes = False Then
        listaPersona StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub txt_TextoBuscar_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyAscii = 13 Then
        listaPersona StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txtCod_Depa_Change()
    
    txtGls_Depa.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", txtCod_Depa.Text, False, " idProv = '00' And idPais = '" & txtCod_Pais.Text & "' ")
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

    txtGls_Distrito.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", txtCod_Distrito.Text, False, "idPais = '" & txtCod_Pais.Text & "'")
    If indCopiaDireccion Then
        indEditando = True
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & "/" & txtGls_Distrito.Text & "/" & txtGls_Prov.Text & "/" & txtGls_Depa.Text
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
On Error GoTo Err
Dim StrMsgError                 As String

    txtGls_Pais.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_Pais.Text, False)
    txtCod_Depa.Text = ""
    txtGls_Depa.Text = ""
    txtCod_Prov.Text = ""
    txtGls_Prov.Text = ""
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""
    
    If txtGls_Pais.Text = "" Or traerCampo("Datos", "IdSunat", "IdDato", txtCod_Pais.Text, False) = "9589" Then
        txtCod_Depa.Vacio = False
        txtGls_Depa.Vacio = False
        txtCod_Prov.Vacio = False
        txtGls_Prov.Vacio = False
        txtCod_Distrito.Vacio = False
        txtGls_Distrito.Vacio = False
    Else
        txtCod_Depa.Vacio = True
        txtGls_Depa.Vacio = True
        txtCod_Prov.Vacio = True
        txtGls_Prov.Vacio = True
        txtCod_Distrito.Vacio = True
        txtGls_Distrito.Vacio = True
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Pais_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PAIS", txtCod_Pais, txtGls_Pais
        KeyAscii = 0
        If txtCod_Pais.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Prov_Change()
    
    txtGls_Prov.Text = traerCampo("Ubigeo", "GlsUbigeo", "idProv", txtCod_Prov.Text, False, " idDpto = '" & txtCod_Depa.Text & "' and idProv <> '00' and idDist = '00' And idPais = '" & txtCod_Pais.Text & "' ")
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""
    If indCopiaDireccion Then
        indEditando = True
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & "/" & txtGls_Distrito.Text & "/" & txtGls_Prov.Text & "/" & txtGls_Depa.Text
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

    If txtCod_TipoPersona.Text = "01001" Then  '--- Persona Natural
        lblApePaterno.Caption = "Apellido Paterno"
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
'        lblRUC.Caption = "D.N.I."
        
    Else '--- Persona Juridica
        lblApePaterno.Caption = "Razón Social"
        lblApeMaterno.Enabled = False
        lblNombres.Enabled = False
        txtGls_Materno.Enabled = False
        txtGls_Nombres.Enabled = False
        txtGls_Materno.BackColor = &HC0FFFF
        txtGls_Nombres.BackColor = &HC0FFFF
        txtGls_Materno.Vacio = True
        txtGls_Nombres.Vacio = True
        txt_RUC.Vacio = False
        lblFecNacimiento.Enabled = False
        dtpFechaNac.Enabled = False
'        lblRUC.Caption = "R.U.C."
    End If

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
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & "/" & txtGls_Distrito.Text & "/" & txtGls_Prov.Text & "/" & txtGls_Depa.Text
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
On Error GoTo Err
Dim strCond                     As String
Dim CSqlC                       As String

    'PQS 07-01-15
    
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
    
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = "Where (GlsPersona Like'%" & strCond & "%' Or IdPersona Like'%" & strCond & "%' Or TipoPersona Like'%" & strCond & "%' " & _
                  "Or Direccion Like'%" & strCond & "%' Or Ruc Like'%" & strCond & "%' Or GlsNombreComercial Like'%" & strCond & "%') "
                    
    End If
    
    CSqlC = "Select IdPersona,GlsPersona,If(TipoPersona = '01001','Natural','Juridica') TipoPersona,Ruc," & _
            "ConCat(Direccion,', ',IfNull(U.GlsUbigeo, '')) Direccion,GlsNombreComercial,If(IfNull(B.IdProveedor,'') <> '','1','0') IndProveedor," & _
            "If(IfNull(C.IdCliente,'') <> '','1','0') IndCliente " & _
            "From Personas P " & _
            "Left Join Ubigeo U " & _
                "On P.IdDistrito = U.IdDistrito And P.IdPais = U.IdPais " & _
            "Left Join Proveedores B " & _
                "On '" & glsEmpresa & "' = B.IdEmpresa And P.IdPersona = B.IdProveedor " & _
            "Left Join Clientes C " & _
                "On '" & glsEmpresa & "' = C.IdEmpresa And P.IdPersona = C.IdCliente " & _
            " " & strCond & _
            "Order By IdPersona"
            
    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = CSqlC
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "IdPersona"
    End With
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub mostrarPersona(strCodPer As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst                         As New ADODB.Recordset
Dim CSqlC                       As String
    
    'PQS 07-01-15
    
    CSqlC = "Select IdPersona,GlsPersona,ApellidoPaterno,ApellidoMaterno,Nombres,TipoPersona,Ruc,Direccion,P.IdPais,P.IdDistrito,U.IdDpto,U.IdProv,Telefonos," & _
            "Mail,FechaNacimiento,DireccionEntrega,GlsContacto,IdTipoDocIdentidad,GlsNombreComercial,IdContacto " & _
            "From Personas P " & _
            "Left Join Ubigeo U " & _
                "On P.IdDistrito = U.IdDistrito And P.IdPais = U.IdPais " & _
            "Where IdPersona = '" & strCodPer & "'"
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    
    txtCod_Depa.Tag = "TidDpto"
    txtCod_Prov.Tag = "TidProv"
    txtCod_Pais.Tag = "TidPais"
    
    CIdContacto = Trim("" & rst.Fields("IdContacto"))
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If traerCampo("Clientes", "IdCliente", "IdCliente", txtCod_Persona.Text, True) = "" Then
        chk_Cliente.Value = 0
        chk_Cliente.Enabled = True
    Else
        chk_Cliente.Value = 1
        chk_Cliente.Enabled = False
    End If
 
    If traerCampo("Proveedores", "IdProveedor", "IdProveedor", txtCod_Persona.Text, True) = "" Then
        chk_Proveedor.Value = 0
        chk_Proveedor.Enabled = True
    Else
        chk_Proveedor.Value = 1
        chk_Proveedor.Enabled = False
    End If
    
    txtCod_Depa.Tag = ""
    txtCod_Prov.Tag = ""
    indCopiaDireccion = False
    
    CSqlC = "Select IndMatriz,IndSubSidiaria,IndAsociada " & _
            "From EmpresasRelacionadas " & _
            "Where IdEmpresa = '" & glsEmpresa & "' And IdPersona = '" & strCodPer & "'"
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
    
        ChkRelacionada.Value = 1
        
        OptRelacionadaTipo(0).Value = IIf(Trim("" & rst.Fields("IndMatriz")) = "1", True, False)
        OptRelacionadaTipo(1).Value = IIf(Trim("" & rst.Fields("IndSubSidiaria")) = "1", True, False)
        OptRelacionadaTipo(2).Value = IIf(Trim("" & rst.Fields("IndAsociada")) = "1", True, False)
    
    Else
        
        ChkRelacionada.Value = 0
    
    End If
    
    rst.Close: Set rst = Nothing
    
    CSqlC = "Select IndPersonal,IndSocio " & _
            "From Personal " & _
            "Where IdEmpresa = '" & glsEmpresa & "' And IdPersona = '" & strCodPer & "'"
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
    
        ChkPersonal.Value = 1
        
        OptEmpleadoTipo(0).Value = IIf(Trim("" & rst.Fields("IndPersonal")) = "1", True, False)
        OptEmpleadoTipo(1).Value = IIf(Trim("" & rst.Fields("IndSocio")) = "1", True, False)
        
    Else
        
        ChkPersonal.Value = 0
    
    End If
    
    rst.Close: Set rst = Nothing
    
    Me.Refresh
    
    Exit Sub
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub actualizaCliente()

    If chk_Cliente.Value Then
        If traerCampo("clientes", "idCliente", "idCliente", txtCod_Persona.Text, True) = "" Then
           csql = "INSERT INTO clientes(idCliente,idEmpresa,idGrupoCliente) VALUES('" & txtCod_Persona.Text & "','" & glsEmpresa & "','12')"
           Cn.Execute csql
        End If
    End If
    
    If chk_Proveedor.Value Then
        If traerCampo("proveedores", "idproveedor", "idproveedor", txtCod_Persona.Text, True) = "" Then
           csql = "INSERT INTO proveedores(idproveedor,idEmpresa,idgrupoproveedor) VALUES('" & txtCod_Persona.Text & "','" & glsEmpresa & "','01')"
           Cn.Execute csql
        End If
    End If

End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim StrCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("Está seguro(a) de eliminar el registro?" & vbCrLf & "Se eliminarán todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    StrCodigo = Trim(txtCod_Persona.Text)
    
    csql = "SELECT idDocVentas FROM docventas WHERE idPerCliente = '" & StrCodigo & "' " & _
            "OR idPerVendedor = '" & StrCodigo & "' " & _
            "OR idPerChofer = '" & StrCodigo & "' " & _
            "OR idPerEmpTrans = '" & StrCodigo & "' " & _
            "OR idPerVendedorCampo = '" & StrCodigo & "'"
                                            
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Ventas)."
        GoTo Err
    End If
    
    csql = "SELECT idValesCab FROM valescab WHERE idProvCliente = '" & StrCodigo & "'"
    If rsValida.State = 1 Then rsValida.Close
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Vales)."
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando el registro
    csql = "DELETE FROM clientes WHERE idCliente = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM choferes WHERE idChofer = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM proveedores WHERE idProveedor = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM sucursales WHERE idSucursal = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM sucursalesempresa WHERE idSucursal = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM usuarios WHERE idUsuario = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM seriexusuario WHERE idUsuario = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM vendedores WHERE idVendedor = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM emptrans WHERE idEmpTrans = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM EmpresasRelacionadas WHERE idPersona = '" & StrCodigo & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM personas WHERE idPersona = '" & StrCodigo & "'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    '--- Nuevo
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub

Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub
