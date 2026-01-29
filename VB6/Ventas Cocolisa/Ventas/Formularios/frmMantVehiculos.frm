VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantVehiculos 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Vehículos"
   ClientHeight    =   6330
   ClientLeft      =   1710
   ClientTop       =   1890
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   7470
      Top             =   45
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
            Picture         =   "frmMantVehiculos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantVehiculos.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   5685
      Left            =   45
      TabIndex        =   14
      Top             =   600
      Width           =   9330
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   165
         TabIndex        =   15
         Top             =   150
         Width           =   8985
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   960
            TabIndex        =   0
            Top             =   255
            Width           =   7815
            _ExtentX        =   13785
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
            Container       =   "frmMantVehiculos.frx":3518
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   4530
         Left            =   165
         OleObjectBlob   =   "frmMantVehiculos.frx":3534
         TabIndex        =   1
         Top             =   1020
         Width           =   8985
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   45
      TabIndex        =   10
      Top             =   600
      Width           =   9315
      Begin VB.CommandButton Cmd_Chofer 
         Height          =   315
         Left            =   8500
         Picture         =   "frmMantVehiculos.frx":4C64
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3870
         Width           =   390
      End
      Begin VB.TextBox Txt_GlsChofer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3870
         Width           =   5850
      End
      Begin VB.CommandButton cmbAyudaEmpTrans 
         Height          =   315
         Left            =   8500
         Picture         =   "frmMantVehiculos.frx":4FEE
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3465
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaMarcaVehi 
         Height          =   315
         Left            =   8500
         Picture         =   "frmMantVehiculos.frx":5378
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1830
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Vehiculo 
         Height          =   315
         Left            =   7995
         TabIndex        =   11
         Tag             =   "TidVehiculo"
         Top             =   315
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
         Container       =   "frmMantVehiculos.frx":5702
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vehiculo 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Tag             =   "TglsVehiculo"
         Top             =   1080
         Width           =   6150
         _ExtentX        =   10848
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
         Container       =   "frmMantVehiculos.frx":571E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Placa 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Tag             =   "TGlsPlaca"
         Top             =   1470
         Width           =   1500
         _ExtentX        =   2646
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
         MaxLength       =   16
         Container       =   "frmMantVehiculos.frx":573A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Modelo 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Tag             =   "TGlsModelo"
         Top             =   2235
         Width           =   4500
         _ExtentX        =   7938
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
         MaxLength       =   128
         Container       =   "frmMantVehiculos.frx":5756
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_MarcaVehi 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Tag             =   "TidMarcaVehi"
         Top             =   1845
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
         Container       =   "frmMantVehiculos.frx":5772
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_MarcaVehi 
         Height          =   315
         Left            =   2640
         TabIndex        =   19
         Top             =   1845
         Width           =   5850
         _ExtentX        =   10319
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
         Container       =   "frmMantVehiculos.frx":578E
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_CodInscrip 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Tag             =   "TGlsCodInscripcion"
         Top             =   3075
         Width           =   4500
         _ExtentX        =   7938
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
         MaxLength       =   128
         Container       =   "frmMantVehiculos.frx":57AA
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Color 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Tag             =   "TGlsColor"
         Top             =   2655
         Width           =   4500
         _ExtentX        =   7938
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
         MaxLength       =   128
         Container       =   "frmMantVehiculos.frx":57C6
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_EmpTrans 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Tag             =   "TidEmpTrans"
         Top             =   3465
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
         Container       =   "frmMantVehiculos.frx":57E2
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_EmpTrans 
         Height          =   315
         Left            =   2640
         TabIndex        =   26
         Top             =   3465
         Width           =   5850
         _ExtentX        =   10319
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
         Container       =   "frmMantVehiculos.frx":57FE
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox Txt_IdChofer 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Tag             =   "TIdChofer"
         Top             =   3870
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
         Container       =   "frmMantVehiculos.frx":581A
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Chofer"
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
         Left            =   495
         TabIndex        =   29
         Top             =   3915
         Width           =   495
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Emp. Transp."
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
         Left            =   495
         TabIndex        =   27
         Top             =   3510
         Width           =   945
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cod. Inscrip."
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
         Left            =   495
         TabIndex        =   24
         Top             =   3165
         Width           =   900
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Placa"
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
         Left            =   495
         TabIndex        =   23
         Top             =   1590
         Width           =   390
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
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
         Left            =   495
         TabIndex        =   22
         Top             =   2355
         Width           =   510
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Color"
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
         Left            =   495
         TabIndex        =   21
         Top             =   2730
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Left            =   495
         TabIndex        =   20
         Top             =   1935
         Width           =   450
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
         Left            =   7350
         TabIndex        =   13
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
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
         Left            =   495
         TabIndex        =   12
         Top             =   1140
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   2170
      ButtonWidth     =   2540
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Nuevo         "
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
            Caption         =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
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
End
Attribute VB_Name = "frmMantVehiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaEmpTrans_Click()

    mostrarAyuda "EMPTRANS", txtCod_EmpTrans, txtGls_EmpTrans

End Sub

Private Sub cmbAyudaMarcaVehi_Click()
    
    mostrarAyuda "MARCAVEHI", txtCod_MarcaVehi, txtGls_MarcaVehi

End Sub

Private Sub Cmd_Chofer_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "CHOFER", Txt_IdChofer, Txt_GlsChofer
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLista, False, False, False, False
    listaVehiculo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    nuevo
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo As String
Dim strMsg      As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaHomonimia "Vehiculos", "GlsVehiculo", "idVehiculo", txtGls_Vehiculo.Text, txtCod_Vehiculo.Text, True, StrMsgError, " or GlsPlaca = '" & txtGls_Placa.Text & "'"
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_Vehiculo.Text = "" Then 'graba
        txtCod_Vehiculo.Text = GeneraCorrelativoAnoMes("Vehiculos", "idVehiculo")
        
        EjecutaSQLForm Me, 0, True, "Vehiculos", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabo"
    
    Else 'modifica
        EjecutaSQLForm Me, 1, True, "Vehiculos", StrMsgError, "idVehiculo"
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modifico"
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    fraGeneral.Enabled = False
    
    listaVehiculo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub nuevo()
    
    limpiaForm Me

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarVehiculo gLista.Columns.ColumnByName("idVehiculo").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = Fals
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
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            fraGeneral.Enabled = True
        Case 4, 7 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
            
            listaVehiculo StrMsgError
            If StrMsgError <> "" Then GoTo Err
    
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
            If indexBoton = 2 Then indHabilitar = True
            Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
            Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
            Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
            Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
            Toolbar1.Buttons(5).Visible = indHabilitar 'Eliminar
            Toolbar1.Buttons(6).Visible = indHabilitar 'Imprimir
            Toolbar1.Buttons(7).Visible = indHabilitar 'Lista
        Case 4, 7 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
    End Select

End Sub

Private Sub Txt_IdChofer_Change()
On Error GoTo Err
Dim StrMsgError     As String
    
    Txt_GlsChofer.Text = traerCampo("Choferes A Inner Join Personas B On A.IdChofer = B.IdPersona", "B.GlsPersona", "A.IdChofer", Txt_IdChofer.Text, True)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdChofer_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    If KeyCode = 113 Then
        mostrarAyuda "CHOFER", Txt_IdChofer, Txt_GlsChofer
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaVehiculo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaVehiculo(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND v.GlsVehiculo LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT v.idVehiculo ,v.GlsVehiculo " & _
           "FROM Vehiculos v WHERE v.idEmpresa = '" & glsEmpresa & "'"
    If strCond <> "" Then csql = csql & strCond
    csql = csql & " ORDER BY v.idVehiculo"
    
If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
    
Set gLista.DataSource = rsdatos

'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idVehiculo"
'    End With
    
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarVehiculo(strCodVehi As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "Select V.IdVehiculo,V.GlsVehiculo,V.GlsPlaca,V.IdMarcaVehi,V.GlsModelo,V.GlsColor,V.GlsCodInscripcion,V.IdEmpTrans,V.IdChofer " & _
           "From Vehiculos V " & _
           "Where V.IdVehiculo = '" & strCodVehi & "' And V.IdEmpresa = '" & glsEmpresa & "'"
           
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_EmpTrans_Change()
    
    If txtCod_EmpTrans.Text <> "" Then
        txtGls_EmpTrans.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_EmpTrans.Text, False)
    Else
        txtGls_EmpTrans.Text = ""
    End If

End Sub

Private Sub txtCod_MarcaVehi_Change()
    
    txtGls_MarcaVehi.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_MarcaVehi.Text, False)

End Sub

Private Sub txtCod_MarcaVehi_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MARCAVEHI", txtCod_MarcaVehi, txtGls_MarcaVehi
        KeyAscii = 0
        If txtCod_MarcaVehi.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub

    strCodigo = Trim(txtCod_Vehiculo.Text)
    
    '--- Validando si existe en ventas
    csql = "SELECT idVehiculo FROM docventas WHERE idVehiculo = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Ventas)"
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando el registro
    csql = "DELETE FROM vehiculos WHERE idvehiculo = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
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
