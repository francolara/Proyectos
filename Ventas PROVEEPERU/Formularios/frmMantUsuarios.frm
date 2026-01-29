VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantUsuarios 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Usuarios"
   ClientHeight    =   8220
   ClientLeft      =   5040
   ClientTop       =   2220
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   180
      Top             =   7260
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
            Picture         =   "frmMantUsuarios.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUsuarios.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   1164
      ButtonWidth     =   2461
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Nuevo        "
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
      Height          =   7425
      Left            =   90
      TabIndex        =   9
      Top             =   675
      Width           =   8325
      Begin VB.CheckBox chkIndJefe 
         Caption         =   "Jefe"
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
         Left            =   7470
         TabIndex        =   7
         Tag             =   "NindJefe"
         Top             =   4275
         Width           =   690
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   " Sucursal por Empresa "
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
         Height          =   2595
         Left            =   180
         TabIndex        =   37
         Top             =   4665
         Width           =   7965
         Begin DXDBGRIDLibCtl.dxDBGrid gSucursales 
            Height          =   2190
            Left            =   105
            OleObjectBlob   =   "frmMantUsuarios.frx":3518
            TabIndex        =   8
            Top             =   285
            Width           =   7740
         End
      End
      Begin VB.CommandButton cmbAyudaPersona 
         Height          =   315
         Left            =   7710
         Picture         =   "frmMantUsuarios.frx":5B6D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   475
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaPerfil 
         Height          =   315
         Left            =   7710
         Picture         =   "frmMantUsuarios.frx":5EF7
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3490
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Persona 
         Height          =   315
         Left            =   1700
         TabIndex        =   18
         Tag             =   "TidUsuario"
         Top             =   495
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
         Container       =   "frmMantUsuarios.frx":6281
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Persona 
         Height          =   315
         Left            =   2610
         TabIndex        =   12
         Top             =   495
         Width           =   5055
         _ExtentX        =   8916
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
         Container       =   "frmMantUsuarios.frx":629D
      End
      Begin CATControls.CATTextBox txtGls_Direccion 
         Height          =   315
         Left            =   1695
         TabIndex        =   13
         Top             =   2370
         Width           =   5985
         _ExtentX        =   10557
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
         Container       =   "frmMantUsuarios.frx":62B9
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Pais 
         Height          =   315
         Left            =   1700
         TabIndex        =   14
         Top             =   870
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
         Container       =   "frmMantUsuarios.frx":62D5
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Pais 
         Height          =   315
         Left            =   2620
         TabIndex        =   15
         Top             =   870
         Width           =   5050
         _ExtentX        =   8916
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
         Container       =   "frmMantUsuarios.frx":62F1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Depa 
         Height          =   315
         Left            =   1700
         TabIndex        =   16
         Top             =   1245
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
         Container       =   "frmMantUsuarios.frx":630D
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Depa 
         Height          =   315
         Left            =   2620
         TabIndex        =   17
         Top             =   1245
         Width           =   5050
         _ExtentX        =   8916
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
         Container       =   "frmMantUsuarios.frx":6329
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Prov 
         Height          =   315
         Left            =   1700
         TabIndex        =   19
         Top             =   1620
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
         Container       =   "frmMantUsuarios.frx":6345
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Prov 
         Height          =   315
         Left            =   2620
         TabIndex        =   20
         Top             =   1620
         Width           =   5050
         _ExtentX        =   8916
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
         Container       =   "frmMantUsuarios.frx":6361
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Distrito 
         Height          =   315
         Left            =   1700
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1995
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
         Container       =   "frmMantUsuarios.frx":637D
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Distrito 
         Height          =   315
         Left            =   2620
         TabIndex        =   22
         Top             =   1995
         Width           =   5050
         _ExtentX        =   8916
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
         Container       =   "frmMantUsuarios.frx":6399
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Perfil 
         Height          =   315
         Left            =   1700
         TabIndex        =   4
         Tag             =   "TidPerfil"
         Top             =   3495
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
         Container       =   "frmMantUsuarios.frx":63B5
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Perfil 
         Height          =   315
         Left            =   2620
         TabIndex        =   23
         Top             =   3495
         Width           =   5050
         _ExtentX        =   8916
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
         Container       =   "frmMantUsuarios.frx":63D1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Usu 
         Height          =   315
         Left            =   1700
         TabIndex        =   2
         Tag             =   "TvarUsuario"
         Top             =   2760
         Width           =   2145
         _ExtentX        =   3784
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
         MaxLength       =   20
         Container       =   "frmMantUsuarios.frx":63ED
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Pass 
         Height          =   315
         Left            =   1700
         TabIndex        =   3
         Tag             =   "TvarPass"
         Top             =   3120
         Width           =   2145
         _ExtentX        =   3784
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
         MaxLength       =   20
         PasswordChar    =   "X"
         Container       =   "frmMantUsuarios.frx":6409
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox CATTextBox1 
         Height          =   315
         Left            =   1700
         TabIndex        =   6
         Tag             =   "Tserieetiquetera"
         Top             =   4275
         Width           =   2145
         _ExtentX        =   3784
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
         MaxLength       =   20
         Container       =   "frmMantUsuarios.frx":6425
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtImpresoraLetras 
         Height          =   315
         Left            =   1700
         TabIndex        =   5
         Tag             =   "TImpresoraLetras"
         Top             =   3900
         Width           =   4530
         _ExtentX        =   7990
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
         Container       =   "frmMantUsuarios.frx":6441
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Impresora Letras"
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
         Left            =   250
         TabIndex        =   39
         Top             =   3960
         Width           =   1230
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Serie Etiquetera"
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
         Left            =   250
         TabIndex        =   38
         Top             =   4365
         Width           =   1140
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
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
         Left            =   250
         TabIndex        =   36
         Top             =   3210
         Width           =   840
      End
      Begin VB.Label Label9 
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   250
         TabIndex        =   35
         Top             =   2850
         Width           =   555
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
         Left            =   250
         TabIndex        =   30
         Top             =   2505
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
         Left            =   250
         TabIndex        =   29
         Top             =   1680
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
         Left            =   250
         TabIndex        =   28
         Top             =   1305
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
         Left            =   250
         TabIndex        =   27
         Top             =   930
         Width           =   300
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
         Left            =   250
         TabIndex        =   26
         Top             =   2130
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Entidad"
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
         Left            =   250
         TabIndex        =   25
         Top             =   555
         Width           =   525
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Perfil"
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
         Left            =   250
         TabIndex        =   24
         Top             =   3555
         Width           =   360
      End
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   7440
      Left            =   90
      TabIndex        =   31
      Top             =   675
      Width           =   8310
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   6315
         Left            =   135
         OleObjectBlob   =   "frmMantUsuarios.frx":645D
         TabIndex        =   1
         Top             =   1005
         Width           =   8055
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   120
         TabIndex        =   32
         Top             =   165
         Width           =   8040
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   990
            TabIndex        =   0
            Top             =   255
            Width           =   6915
            _ExtentX        =   12197
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
            Container       =   "frmMantUsuarios.frx":8032
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
            TabIndex        =   33
            Top             =   300
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmMantUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim indBoton As Integer

Private Sub cmbAyudaPerfil_Click()

    mostrarAyuda "PERFIL", txtCod_Perfil, txtGls_Perfil, " And CodSistema = '" & StrcodSistema & "'  "

End Sub

Private Sub cmbAyudaPersona_Click()
    
    mostrarAyuda "PERSONAUSUARIO", txtCod_Persona, txtGls_Persona
    If txtCod_Persona.Text <> "" Then mostrarDatosPersona

End Sub

Private Sub mostrarDatosPersona()
Dim rst As New ADODB.Recordset

    csql = "SELECT idPersona,GlsPersona," & _
            "direccion, p.iddistrito, u.idDpto, u.idProv " & _
            "FROM personas p,ubigeo u " & _
            "WHERE p.iddistrito = u.iddistrito and idpersona = '" & Trim(txtCod_Persona.Text) & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        txtCod_Pais.Text = "02001"
        txtCod_Depa.Text = "" & rst.Fields("idDpto")
        txtCod_Prov.Text = "" & rst.Fields("idProv")
        txtCod_Distrito.Text = "" & rst.Fields("iddistrito")
        txtGls_Direccion.Text = "" & rst.Fields("direccion")
    Else
        txtCod_Pais.Text = ""
        txtCod_Depa.Text = ""
        txtCod_Prov.Text = ""
        txtCod_Distrito.Text = ""
        txtGls_Direccion.Text = ""
    End If
    rst.Close: Set rst = Nothing

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    Me.left = 0
    Me.top = 0
    ConfGrid gLista, False, False, False, False
    ConfGrid gSucursales, True, False, False, False
    
    listaUsuario StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 6
    
    nuevo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarUsuario gLista.Columns.ColumnByName("idUsuario").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    txtCod_Persona.Enabled = False
    cmbAyudaPersona.Enabled = False
    habilitaBotones 2
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gSucursales_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gSucursales.Columns.ColumnByFieldName("item").Value = gSucursales.Count
        gSucursales.Dataset.Post
    End If

End Sub

Private Sub gSucursales_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If gSucursales.Columns.ColumnByFieldName("idEmpresa").Value = "" Or gSucursales.Columns.ColumnByFieldName("idSucursal").Value = "" Then
            Allow = False
        Else
            gSucursales.Columns.FocusedIndex = gSucursales.Columns.ColumnByFieldName("idEmpresa").Index
        End If
    End If

End Sub

Private Sub gSucursales_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim strDes As String
    
    Select Case Column.Index
        Case gSucursales.Columns.ColumnByFieldName("GlsEmpresa").Index
            strCod = gSucursales.Columns.ColumnByFieldName("idEmpresa").Value
            strDes = gSucursales.Columns.ColumnByFieldName("GlsEmpresa").Value
            mostrarAyudaTexto "EMPRESA", strCod, strDes
            gSucursales.Dataset.Edit
            gSucursales.Columns.ColumnByFieldName("idEmpresa").Value = strCod
            gSucursales.Columns.ColumnByFieldName("GlsEmpresa").Value = strDes
            gSucursales.Dataset.Post
        
        Case gSucursales.Columns.ColumnByFieldName("GlsSucursal").Index
            strCod = gSucursales.Columns.ColumnByFieldName("idSucursal").Value
            strDes = gSucursales.Columns.ColumnByFieldName("GlsSucursal").Value
            mostrarAyudaTexto "SUCURSAL", strCod, strDes
            gSucursales.Dataset.Edit
            gSucursales.Columns.ColumnByFieldName("idSucursal").Value = strCod
            gSucursales.Columns.ColumnByFieldName("GlsSucursal").Value = strDes
            gSucursales.Dataset.Post
    End Select
    
End Sub

Private Sub gSucursales_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gSucursales.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gSucursales.Count = 1 Then
                    gSucursales.Dataset.Edit
                    gSucursales.Columns.ColumnByFieldName("Item").Value = 1
                    gSucursales.Columns.ColumnByFieldName("idEmpresa").Value = ""
                    gSucursales.Columns.ColumnByFieldName("GlsEmpresa").Value = ""
                    gSucursales.Columns.ColumnByFieldName("idSucursal").Value = ""
                    gSucursales.Columns.ColumnByFieldName("GlsSucursal").Value = ""
                    gSucursales.Dataset.Post
                
                Else
                    gSucursales.Dataset.Delete
                    gSucursales.Dataset.First
                    Do While Not gSucursales.Dataset.EOF
                        i = i + 1
                        gSucursales.Dataset.Edit
                        gSucursales.Columns.ColumnByFieldName("Item").Value = i
                        gSucursales.Dataset.Post
                        gSucursales.Dataset.Next
                    Loop
                    If gSucursales.Dataset.State = dsEdit Or gSucursales.Dataset.State = dsInsert Then
                        gSucursales.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gSucursales.Dataset.State = dsEdit Or gSucursales.Dataset.State = dsInsert Then
              gSucursales.Dataset.Post
        End If
    End If

End Sub

Private Sub gSucursales_OnKeyPress(Key As Integer)
Dim strCod As String
Dim strDes As String

    Select Case gSucursales.Columns.FocusedColumn.Index
        Case gSucursales.Columns.ColumnByFieldName("GlsEmpresa").Index
            strCod = gSucursales.Columns.ColumnByFieldName("idEmpresa").Value
            strDes = gSucursales.Columns.ColumnByFieldName("GlsEmpresa").Value
            
            mostrarAyudaKeyasciiTexto Key, "EMPRESA", strCod, strDes
            Key = 0
            gSucursales.Dataset.Edit
            gSucursales.Columns.ColumnByFieldName("idEmpresa").Value = strCod
            gSucursales.Columns.ColumnByFieldName("GlsEmpresa").Value = strDes
            gSucursales.Dataset.Post
            gSucursales.SetFocus
                
        Case gSucursales.Columns.ColumnByFieldName("GlsSucursal").Index
            strCod = gSucursales.Columns.ColumnByFieldName("idSucursal").Value
            strDes = gSucursales.Columns.ColumnByFieldName("GlsSucursal").Value
            
            mostrarAyudaKeyasciiTexto Key, "SUCURSAL", strCod, strDes
            Key = 0
            gSucursales.Dataset.Edit
            gSucursales.Columns.ColumnByFieldName("idSucursal").Value = strCod
            gSucursales.Columns.ColumnByFieldName("GlsSucursal").Value = strDes
            gSucursales.Dataset.Post
            gSucursales.SetFocus
    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            indBoton = 0
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
            chkIndJefe.Value = 0
            
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            indBoton = 1
            fraGeneral.Enabled = True
        Case 4, 6 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        Case 5 'Imprimir
            gLista.m.ExportToXLS App.Path & "\Temporales\Mantenimiento_Usuarios.xls"
            ShellEx App.Path & "\Temporales\Mantenimiento_Usuarios.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 7 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaUsuario StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Perfil_Change()
    
    txtGls_Perfil.Text = traerCampo("perfil", "GlsPerfil", "idPerfil", txtCod_Perfil.Text, True)

End Sub

Private Sub txtCod_Perfil_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PERFIL", txtCod_Perfil, txtGls_Perfil
        KeyAscii = 0
        If txtCod_Perfil.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Persona_Change()
    
    txtGls_Persona.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Persona.Text, False)

End Sub

Private Sub txtCod_Depa_Change()
    
    txtGls_Depa.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", txtCod_Depa.Text, False, " idProv = '00'")

End Sub

Private Sub txtCod_Distrito_Change()
    
    txtGls_Distrito.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", txtCod_Distrito.Text, False)

End Sub

Private Sub txtCod_Pais_Change()
    
    txtGls_Pais.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_Pais.Text, False)

End Sub

Private Sub txtCod_Persona_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PERSONAUSUARIO", txtCod_Persona, txtGls_Persona
        KeyAscii = 0
        If txtCod_Persona.Text <> "" Then
            mostrarDatosPersona
            SendKeys "{tab}"
        End If
    End If

End Sub

Private Sub txtCod_Prov_Change()
    
    txtGls_Prov.Text = traerCampo("Ubigeo", "GlsUbigeo", "idProv", txtCod_Prov.Text, False, " idDpto = '" & txtCod_Depa.Text & "' and idProv <> '00' and idDist = '00'")

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
            Toolbar1.Buttons(5).Visible = True
            Toolbar1.Buttons(6).Visible = False
    End Select

End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo As String
Dim strMsg      As String
Dim csql        As String
Dim rsdatos     As New ADODB.Recordset
Dim csqlconsul  As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    If indBoton = 0 Then 'graba
        EjecutaSQLFormUsuario Me, 0, True, "usuarios", StrMsgError, "", gSucursales, "sucursalesempresa", "idUsuario", txtCod_Persona.Text
        If StrMsgError <> "" Then GoTo Err
        
        csql = "insert into perfilesporusuario (idempresa,idusuario,idperfil,CodSistema) values " & _
               "('" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & txtCod_Perfil.Text & "','" & StrcodSistema & "')"
        Cn.Execute (csql)
        
        strMsg = "Grabo"
    
    Else 'modifica
        EjecutaSQLFormUsuario Me, 1, True, "usuarios", StrMsgError, "idUsuario", gSucursales, "sucursalesempresa", "idUsuario", txtCod_Persona.Text
        If StrMsgError <> "" Then GoTo Err
        
        csqlconsul = "select * from perfilesporusuario  where idusuario = '" & txtCod_Persona.Text & "' " & _
                    "  and CodSistema = '" & StrcodSistema & "' and idempresa = '" & glsEmpresa & "' "
        If rsdatos.State = 1 Then rsdatos.Close
        rsdatos.Open csqlconsul, Cn, adOpenStatic, adLockReadOnly
        
        If Not rsdatos.EOF Then
            csql = "update perfilesporusuario set idperfil = '" & txtCod_Perfil.Text & "' where idusuario = '" & txtCod_Persona.Text & "' " & _
                   "and idperfil = '" & rsdatos.Fields("idperfil") & "'  and CodSistema = '" & StrcodSistema & "' and idempresa = '" & glsEmpresa & "' "
            Cn.Execute (csql)
        Else
            csql = "insert into perfilesporusuario (idempresa,idusuario,idperfil,CodSistema) values " & _
                   "('" & glsEmpresa & "','" & txtCod_Persona.Text & "','" & txtCod_Perfil.Text & "','" & StrcodSistema & "')"
            Cn.Execute (csql)
        
        End If
        strMsg = "Modifico"
    End If
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    listaUsuario StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    limpiaForm Me
    txtCod_Persona.Enabled = True
    cmbAyudaPersona.Enabled = True
    
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "idEmpresa", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsEmpresa", adVarChar, 250, adFldIsNullable
    rst.Fields.Append "idSucursal", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsSucursal", adVarChar, 250, adFldIsNullable
    rst.Open
    
    rst.AddNew
    rst.Fields("Item") = 1
    rst.Fields("idEmpresa") = ""
    rst.Fields("GlsEmpresa") = ""
    rst.Fields("idSucursal") = ""
    rst.Fields("GlsSucursal") = ""
    
    mostrarDatosGridSQL gSucursales, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gSucursales.Columns.FocusedIndex = gSucursales.Columns.ColumnByFieldName("GlsEmpresa").Index
        
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub listaUsuario(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND GlsPersona LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT v.idUsuario ,p.GlsPersona ," & _
            "concat(p.direccion,', ',u.GlsUbigeo) as Direccion " & _
            "FROM usuarios v,personas p,ubigeo u " & _
            "WHERE v.idUsuario = p.idPersona AND p.iddistrito = u.iddistrito AND v.idEmpresa = '" & glsEmpresa & "'" & strCond & _
            "ORDER BY idusuario"
            
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
'        .KeyField = "idusuario"
'    End With
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarUsuario(strCodUsu As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim rsdatos  As New ADODB.Recordset
Dim sql As String

    csql = "SELECT v.indJefe,v.idusuario,v.idPerfil,varUsuario,varPass,v.serieetiquetera, v.ImpresoraLetras " & _
           "FROM usuarios v " & _
           "WHERE v.idEmpresa = '" & glsEmpresa & "' AND v.idusuario = '" & strCodUsu & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    mostrarDatosPersona
    
    csql = "SELECT se.item,se.idEmpresa,e.glsEmpresa,se.idSucursal,p.glsPersona GlsSucursal " & _
            "FROM sucursalesempresa se,empresas e,sucursales s,personas p " & _
            "WHERE se.idEmpresa = e.idEmpresa " & _
            "AND se.idSucursal = s.idSucursal " & _
            "AND s.idSucursal = p.idPersona " & _
            "AND se.idEmpresa = '" & glsEmpresa & "' " & _
            "AND s.idEmpresa = '" & glsEmpresa & "' " & _
            "AND se.idUsuario = '" & strCodUsu & "'"
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idEmpresa", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsEmpresa", adVarChar, 250, adFldIsNullable
    rsg.Fields.Append "idSucursal", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsSucursal", adVarChar, 250, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idEmpresa") = ""
        rsg.Fields("GlsEmpresa") = ""
        rsg.Fields("idSucursal") = ""
        rsg.Fields("GlsSucursal") = ""
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = rst.Fields("Item")
            rsg.Fields("idEmpresa") = rst.Fields("idEmpresa")
            rsg.Fields("GlsEmpresa") = rst.Fields("GlsEmpresa")
            rsg.Fields("idSucursal") = rst.Fields("idSucursal")
            rsg.Fields("GlsSucursal") = rst.Fields("GlsSucursal")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gSucursales, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
          
    sql = "Select P.IdPerfil,P.GlsPerfil " & _
          "From Perfil P " & _
          "Inner Join PerfilesPorUsuario Pxu " & _
            "On P.IdEmpresa = Pxu.IdEmpresa And P.IdPerfil = Pxu.IdPerfil And P.CodSistema = Pxu.CodSistema " & _
          "Where P.IdEmpresa = '" & glsEmpresa & "' And Pxu.IdUsuario = '" & txtCod_Persona.Text & "' And Pxu.CodSistema = '" & StrcodSistema & "'"
    
    If rsdatos.State = 1 Then rsdatos.Close
    rsdatos.Open sql, Cn, adOpenStatic, adLockBatchOptimistic
    
    If Not rsdatos.EOF Then
        txtCod_Perfil.Text = "" & Trim(rsdatos.Fields("idperfil"))
        txtGls_Perfil.Text = "" & Trim(rsdatos.Fields("glsperfil"))
    Else
        txtCod_Perfil.Text = ""
        txtGls_Perfil.Text = ""
    End If
    Me.Refresh

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub EjecutaSQLFormUsuario(F As Form, tipoOperacion As Integer, indEmpresa As Boolean, strTabla As String, ByRef StrMsgError As String, Optional strCampoCod As String, Optional g As dxDBGrid, Optional strTablaDet As String, Optional strCampoDet As String, Optional strDataCampo As String)
On Error GoTo Err
Dim C As Object
Dim csql As String
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strValCod As String
Dim strCampoEmpresa As String
Dim strValorEmpresa As String
Dim strCondEmpresa As String
Dim indTrans As Boolean

    If indEmpresa Then
        strCampoEmpresa = ",idEmpresa"
        strValorEmpresa = ",'" & glsEmpresa & "'"
        strCondEmpresa = " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    indTrans = False
    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & C.Value & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    Case 1
                        If UCase(strCampoCod) <> UCase(strCampo) Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = C.Value
                                Case "T"
                                    strValores = "'" & C.Value & "'"
                                Case "F"
                                    strValores = "'" & Format(C.Value, "yyyy-mm-dd") & "'"
                            End Select
                            strCampos = strCampos & strCampo & "=" & strValores & ","
                        Else
                            strValCod = C.Value
                        End If
                End Select
            End If
        End If
    Next
    
    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
    
    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            csql = "INSERT INTO " & strTabla & "(" & strCampos & strCampoEmpresa & ") VALUES(" & strValores & strValorEmpresa & ")"
        Case 1
            csql = "UPDATE " & strTabla & " SET " & strCampos & " WHERE " & strCampoCod & " = '" & strValCod & "'" & strCondEmpresa
    End Select
    indTrans = True
    Cn.BeginTrans
    
    'Graba controles
    If strCampos <> "" Then
        Cn.Execute csql
    End If
    
    'Grabando Grilla
    If TypeName(g) <> "Nothing" Then
        Cn.Execute "DELETE FROM " & strTablaDet & " WHERE " & strCampoDet & " = '" & strDataCampo & "'" & strCondEmpresa
        
        g.Dataset.First
        Do While Not g.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To g.Columns.Count - 1
                If UCase(left(g.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(g.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(g.Columns(i).ObjectName, 3)
                    
                    strCampos = strCampos & strCampo & ","
                    
                    Select Case strTipoDato
                        Case "N"
                            strValores = strValores & g.Columns(i).Value & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO " & strTablaDet & "(" & strCampos & "," & strCampoDet & ") VALUES(" & strValores & ",'" & strDataCampo & "')"
            Cn.Execute csql
            
            g.Dataset.Next
        Loop
    End If
    Cn.CommitTrans
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
End Sub
