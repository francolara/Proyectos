VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantUbigeo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Ubigeo"
   ClientHeight    =   8820
   ClientLeft      =   1665
   ClientTop       =   1530
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   495
      Top             =   7515
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
            Picture         =   "frmMantUbigeo.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantUbigeo.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   8175
      Left            =   45
      TabIndex        =   11
      Top             =   600
      Width           =   10200
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   165
         TabIndex        =   12
         Top             =   150
         Width           =   9855
         Begin VB.ComboBox cbx_TipoUbigeo 
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
            ItemData        =   "frmMantUbigeo.frx":3518
            Left            =   7515
            List            =   "frmMantUbigeo.frx":3528
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   315
            Width           =   2190
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1005
            TabIndex        =   0
            Top             =   300
            Width           =   5580
            _ExtentX        =   9843
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
            Container       =   "frmMantUbigeo.frx":3555
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
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
            Left            =   7065
            TabIndex        =   24
            Top             =   375
            Width           =   300
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
            Left            =   165
            TabIndex        =   13
            Top             =   345
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   6945
         Left            =   165
         OleObjectBlob   =   "frmMantUbigeo.frx":3571
         TabIndex        =   2
         Top             =   1050
         Width           =   9900
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8160
      Left            =   45
      TabIndex        =   8
      Top             =   600
      Width           =   10185
      Begin VB.Frame fraOtros 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   300
         TabIndex        =   26
         Top             =   1920
         Width           =   7905
         Begin VB.CommandButton cmbAyudaZona 
            Height          =   315
            Left            =   7425
            Picture         =   "frmMantUbigeo.frx":6C38
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   430
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Zona 
            Height          =   315
            Left            =   1350
            TabIndex        =   7
            Top             =   430
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
            Container       =   "frmMantUbigeo.frx":6FC2
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Zona 
            Height          =   315
            Left            =   2280
            TabIndex        =   28
            Top             =   430
            Width           =   5130
            _ExtentX        =   9049
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
            Container       =   "frmMantUbigeo.frx":6FDE
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtGlosa 
            Height          =   315
            Left            =   1350
            TabIndex        =   6
            Tag             =   "TglsCaja"
            Top             =   75
            Width           =   6060
            _ExtentX        =   10689
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
            Container       =   "frmMantUbigeo.frx":6FFA
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin VB.Label lblGlosa 
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
            Left            =   75
            TabIndex        =   30
            Top             =   150
            Width           =   540
         End
         Begin VB.Label lblZona 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Zona:"
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
            Left            =   75
            TabIndex        =   29
            Top             =   450
            Width           =   420
         End
      End
      Begin VB.Frame fraUbigeo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   300
         TabIndex        =   14
         Top             =   750
         Width           =   7905
         Begin VB.CommandButton cmbAyudaProvincia 
            Height          =   315
            Left            =   7425
            Picture         =   "frmMantUbigeo.frx":7016
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   825
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaDepartamento 
            Height          =   315
            Left            =   7425
            Picture         =   "frmMantUbigeo.frx":73A0
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   450
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaPais 
            Height          =   315
            Left            =   7425
            Picture         =   "frmMantUbigeo.frx":772A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   90
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Pais 
            Height          =   315
            Left            =   1350
            TabIndex        =   3
            Top             =   105
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
            Container       =   "frmMantUbigeo.frx":7AB4
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Pais 
            Height          =   315
            Left            =   2280
            TabIndex        =   16
            Top             =   105
            Width           =   5130
            _ExtentX        =   9049
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
            Container       =   "frmMantUbigeo.frx":7AD0
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Departamento 
            Height          =   315
            Left            =   1350
            TabIndex        =   4
            Top             =   465
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
            Container       =   "frmMantUbigeo.frx":7AEC
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Departamento 
            Height          =   315
            Left            =   2280
            TabIndex        =   19
            Top             =   465
            Width           =   5130
            _ExtentX        =   9049
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
            Container       =   "frmMantUbigeo.frx":7B08
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Provincia 
            Height          =   315
            Left            =   1350
            TabIndex        =   5
            Top             =   840
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
            Container       =   "frmMantUbigeo.frx":7B24
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Provincia 
            Height          =   315
            Left            =   2280
            TabIndex        =   22
            Top             =   840
            Width           =   5130
            _ExtentX        =   9049
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
            Container       =   "frmMantUbigeo.frx":7B40
            Vacio           =   -1  'True
         End
         Begin VB.Label Label5 
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
            Left            =   75
            TabIndex        =   23
            Top             =   885
            Width           =   705
         End
         Begin VB.Label Label4 
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
            Left            =   75
            TabIndex        =   20
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label2 
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
            Left            =   75
            TabIndex        =   17
            Top             =   150
            Width           =   345
         End
      End
      Begin CATControls.CATTextBox txtCodigo 
         Height          =   315
         Left            =   7230
         TabIndex        =   9
         Top             =   225
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
         Container       =   "frmMantUbigeo.frx":7B5C
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_TipoDato 
         Height          =   285
         Left            =   3090
         TabIndex        =   31
         Top             =   6735
         Visible         =   0   'False
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
         Container       =   "frmMantUbigeo.frx":7B78
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Distrito 
         Height          =   285
         Left            =   6315
         TabIndex        =   32
         Top             =   6735
         Visible         =   0   'False
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
         Container       =   "frmMantUbigeo.frx":7B94
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
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
         Left            =   6630
         TabIndex        =   10
         Top             =   255
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1164
      ButtonWidth     =   2381
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "        Nuevo        "
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
Attribute VB_Name = "frmMantUbigeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbx_TipoUbigeo_Click()
On Error GoTo Err
Dim StrMsgError As String
    Me.left = 0
    Me.top = 0
    listaUbigeo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaDepartamento_Click()
    
    mostrarAyuda "DEPARTAMENTO", txtCod_Departamento, txtGls_Departamento, "AND idPais = '" & txtCod_Pais.Text & "'"
    If txtCod_Departamento.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaPais_Click()
    
    mostrarAyuda "PAIS", txtCod_Pais, txtGls_Pais
    If txtCod_Pais.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaProvincia_Click()
    
    mostrarAyuda "PROVINCIA", txtCod_Provincia, txtGls_Provincia, "AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Departamento.Text + "'"
    If txtCod_Provincia.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaZona_Click()
    
    mostrarAyuda "ZONAS", txtCod_Zona, txtGls_Zona
    If txtCod_Zona.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLista, False, False, False, False
    cbx_TipoUbigeo.ListIndex = 0
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
Dim RsC As New ADODB.Recordset
Dim strCodigo As String
Dim strMsg      As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Select Case cbx_TipoUbigeo.ListIndex
        Case 1 'Departamento
            txtCod_Pais.Tag = "TidPais"
            txtCod_Departamento.Tag = "TidDpto"
            txtCod_Provincia.Text = "00"
            txtCod_Provincia.Tag = "TidProv"
            txtCod_Distrito.Text = "00"
            txtCod_Distrito.Tag = "TidDist"
        
        Case 2 'Provincia
            txtCod_Pais.Tag = "TidPais"
            txtCod_Departamento.Tag = "TidDpto"
            txtCod_Provincia.Tag = "TidProv"
            txtCod_Distrito.Text = "00"
            txtCod_Distrito.Tag = "TidDist"
        
        Case 3 'Distrito
            txtCod_Pais.Tag = "TidPais"
            txtCod_Departamento.Tag = "TidDpto"
            txtCod_Provincia.Tag = "TidProv"
            txtCod_Distrito.Tag = "TidDist"
            txtCod_Zona.Tag = "TidZona"
    End Select
    
    If txtCodigo.Text = "" Then 'graba
        If cbx_TipoUbigeo.ListIndex = 0 Then
            txtCodigo.Text = generaCorrelativo("datos", "idDato", 3, "02", False)
            EjecutaSQLForm Me, 0, False, "datos", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Else
            Select Case cbx_TipoUbigeo.ListIndex
                Case 1 '--- Departamento
                    csql = "SELECT MAX(idDpto) as Codigo FROM ubigeo WHERE idPais = '" & txtCod_Pais.Text & "' AND idProv = '00' AND idDist = '00'"
                    RsC.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                    If Not RsC.EOF Then
                        txtCod_Departamento.Text = Format(CStr(Val("" & RsC.Fields("Codigo")) + 1), "00")
                    Else
                        txtCod_Departamento.Text = "01"
                    End If
                    txtCod_Provincia.Text = "00"
                    txtCod_Distrito.Text = "00"
                    txtCodigo.Text = txtCod_Departamento.Text & txtCod_Provincia.Text & txtCod_Distrito.Text
                    
                Case 2 '--- Provincia
                    csql = "SELECT MAX(idProv) as Codigo FROM ubigeo WHERE idPais = '" & txtCod_Pais.Text & "' AND idDpto <> '00' AND idDist = '00'"
                    RsC.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                    If Not RsC.EOF Then
                        txtCod_Provincia.Text = Format(CStr(Val("" & RsC.Fields("Codigo")) + 1), "00")
                    Else
                        txtCod_Provincia.Text = "01"
                    End If
                    txtCod_Distrito.Text = "00"
                    txtCodigo.Text = txtCod_Departamento.Text & txtCod_Provincia.Text & txtCod_Distrito.Text
                
                Case 3 '--- Distrito
                    csql = "SELECT MAX(idDist) as Codigo FROM ubigeo WHERE idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & Trim(txtCod_Departamento.Text) & "'  AND idProv = '" & Trim(txtCod_Provincia.Text) & "'"
                    RsC.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                    If Not RsC.EOF Then
                        txtCod_Distrito.Text = Format(CStr(Val("" & RsC.Fields("Codigo")) + 1), "00")
                    Else
                        txtCod_Distrito.Text = "01"
                    End If
                    txtCodigo.Text = txtCod_Departamento.Text & txtCod_Provincia.Text & txtCod_Distrito.Text
                    txtCod_Distrito.Text = right(txtCodigo.Text, 2)
            End Select
            
            EjecutaSQLForm Me, 0, False, "ubigeo", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
        strMsg = "Grabo"
    
    Else '--- Modifica
        If cbx_TipoUbigeo.ListIndex = 0 Then
            EjecutaSQLForm Me, 1, False, "datos", StrMsgError, "idDato"
            If StrMsgError <> "" Then GoTo Err
        Else
            Select Case cbx_TipoUbigeo.ListIndex
                Case 1, 2 '--- Departamento / Provincia
                    csql = "UPDATE ubigeo SET GlsUbigeo = '" & txtGlosa.Text & "' " & _
                            "WHERE idPais = '" & txtCod_Pais.Text & "' AND idDistrito = '" & txtCodigo.Text & "'"
                Case 3 '--- Distrito
                    csql = "UPDATE ubigeo SET GlsUbigeo = '" & txtGlosa.Text & "',idZona = '" & txtCod_Zona.Text & "' " & _
                            "WHERE idPais = '" & txtCod_Pais.Text & "' AND idDistrito = '" & txtCodigo.Text & "'"
            End Select
            Cn.Execute csql
        End If
        strMsg = "Modifico"
    End If
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    listaUbigeo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()
Dim C As Object

    For Each C In Me.Controls
        If TypeOf C Is TextBox Or TypeOf C Is CATTextBox Then
            If C.Name <> txt_TextoBuscar.Name Then
                C.Text = ""
                C.Tag = ""
            End If
        End If
        
        If TypeOf C Is DTPicker Then
            C.Value = getFechaSistema
            C.Tag = ""
        End If
    Next
    
    fraUbigeo.Enabled = True
    lblZona.Visible = False
    txtCod_Zona.Visible = False
    txtGls_Zona.Visible = False
    cmbAyudaZona.Visible = False
    
    fraUbigeo.Height = 385 * (cbx_TipoUbigeo.ListIndex + 0)
    fraOtros.top = fraUbigeo.top + fraUbigeo.Height '+ 45
    
    Select Case cbx_TipoUbigeo.ListIndex
        Case 0 '--- Pais
            txtCod_TipoDato.Tag = "TidTipoDatos"
            txtCod_TipoDato.Text = "02"
            
            txtCodigo.Tag = "TidDato"
            lblGlosa.Caption = "Pais"
            txtGlosa.Tag = "TGlsDato"
        Case 1 '--- Departamento
            txtCodigo.Tag = "TidDistrito"
            lblGlosa.Caption = "Departamento"
            txtGlosa.Tag = "TGlsUbigeo"
            txtCod_Pais.Vacio = False
        
        Case 2 '--- Provincia
            txtCodigo.Tag = "TidDistrito"
            lblGlosa.Caption = "Provincia"
            txtGlosa.Tag = "TGlsUbigeo"
            txtCod_Pais.Vacio = False
            txtCod_Departamento.Vacio = False
        
        Case 3 '--- Distrito
            txtCodigo.Tag = "TidDistrito"
            lblGlosa.Caption = "Distrito"
            txtGlosa.Tag = "TGlsUbigeo"
            txtCod_Distrito.Tag = "TidDist"
            
            lblZona.Visible = True
            txtCod_Zona.Visible = True
            txtGls_Zona.Visible = True
            cmbAyudaZona.Visible = True
            
            txtCod_Pais.Vacio = False
            txtCod_Departamento.Vacio = False
            txtCod_Provincia.Vacio = False
            txtCod_Zona.Vacio = True
    End Select
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarUbigeo gLista.Columns.ColumnByName("C0").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnReloadGroupList()
    
    gLista.m.FullExpand

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
        Case 4, 7  'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
            
            listaUbigeo StrMsgError
            If StrMsgError <> "" Then GoTo Err
    
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 6 'Imprimir
            gLista.m.ExportToXLS App.Path & "\Temporales\Mantenimiento_Ubigeo.xls"
            ShellEx App.Path & "\Temporales\Mantenimiento_Ubigeo.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
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
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
    End Select

End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaUbigeo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaUbigeo(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCampoMostrar As String
Dim rsdatos                     As New ADODB.Recordset

    gLista.Columns.ColumnByName("C0").FieldName = "Codigo"
    gLista.Columns.ColumnByName("D1").GroupIndex = -1
    gLista.Columns.ColumnByName("D2").GroupIndex = -1
    gLista.Columns.ColumnByName("D3").GroupIndex = -1
    
    For i = 0 To gLista.Columns.Count - 1
        gLista.Columns(i).Visible = False
    Next
    
    Select Case cbx_TipoUbigeo.ListIndex
        Case 0 '--- Pais
            strCampoMostrar = "p.GlsDato"
            gLista.Columns.ColumnByName("C1").FieldName = "idPais"
            gLista.Columns.ColumnByName("D1").FieldName = "GlsPais"
            gLista.Columns.ColumnByName("C1").Visible = True
            gLista.Columns.ColumnByName("D1").Visible = True
            
            csql = "SELECT p.idDato as Codigo,p.idDato AS idPais,p.GlsDato AS GlsPais " & _
                   "FROM datos p " & _
                   "WHERE p.idtipoDatos = '02'"

        Case 1 '--- Departamento
            strCampoMostrar = "d.GlsUbigeo"
            gLista.Columns.ColumnByName("C1").FieldName = "idPais"
            gLista.Columns.ColumnByName("D1").FieldName = "GlsPais"
            gLista.Columns.ColumnByName("C2").FieldName = "idDpto"
            gLista.Columns.ColumnByName("D2").FieldName = "GlsDpto"
            gLista.Columns.ColumnByName("D1").Visible = True
            gLista.Columns.ColumnByName("C2").Visible = True
            gLista.Columns.ColumnByName("D2").Visible = True
            
            csql = "SELECT d.idDistrito as Codigo,p.idDato AS idPais, p.GlsDato AS GlsPais,d.idDpto, d.GlsUbigeo AS GlsDpto " & _
                    "FROM datos p,ubigeo d " & _
                    "WHERE P.idDato = D.idPais " & _
                    "AND p.idtipoDatos = '02' " & _
                    "AND d.idProv = '00' AND d.idDist = '00'"
  
      Case 2 '--- Provincia
          strCampoMostrar = "r.GlsUbigeo"
          gLista.Columns.ColumnByName("C1").FieldName = "idPais"
          gLista.Columns.ColumnByName("D1").FieldName = "GlsPais"
          gLista.Columns.ColumnByName("C2").FieldName = "idDpto"
          gLista.Columns.ColumnByName("D2").FieldName = "GlsDpto"
          gLista.Columns.ColumnByName("C3").FieldName = "idProv"
          gLista.Columns.ColumnByName("D3").FieldName = "GlsProv"
          
          gLista.Columns.ColumnByName("D1").Visible = True
          gLista.Columns.ColumnByName("D2").Visible = True
          gLista.Columns.ColumnByName("C3").Visible = True
          gLista.Columns.ColumnByName("D3").Visible = True
          
          csql = "SELECT r.idDistrito as Codigo,p.idDato AS idPais, p.GlsDato AS GlsPais,d.idDpto, d.GlsUbigeo AS GlsDpto, r.idProv, r.GlsUbigeo AS GlsProv " & _
                    "FROM datos p,ubigeo d,ubigeo r " & _
                    "WHERE P.idDato = D.idPais " & _
                    "AND d.idPais = r.idPais " & _
                    "AND d.idDpto = r.idDpto " & _
                    "AND p.idtipoDatos = '02' " & _
                    "AND d.idProv = '00' AND d.idDist = '00' " & _
                    "AND r.idProv <> '00' AND r.idDist = '00'"
    
      Case 3 '--- Distrito
          strCampoMostrar = "i.GlsUbigeo"
          gLista.Columns.ColumnByName("C1").FieldName = "idPais"
          gLista.Columns.ColumnByName("D1").FieldName = "GlsPais"
          gLista.Columns.ColumnByName("C2").FieldName = "idDpto"
          gLista.Columns.ColumnByName("D2").FieldName = "GlsDpto"
          gLista.Columns.ColumnByName("C3").FieldName = "idProv"
          gLista.Columns.ColumnByName("D3").FieldName = "GlsProv"
          gLista.Columns.ColumnByName("C4").FieldName = "idDist"
          gLista.Columns.ColumnByName("D4").FieldName = "GlsDist"
          
          gLista.Columns.ColumnByName("D1").Visible = True
          gLista.Columns.ColumnByName("D2").Visible = True
          gLista.Columns.ColumnByName("D3").Visible = True
          gLista.Columns.ColumnByName("C4").Visible = True
          gLista.Columns.ColumnByName("D4").Visible = True
      
          csql = "SELECT i.idDistrito as Codigo,p.idDato AS idPais,p.GlsDato AS GlsPais,d.idDpto, d.GlsUbigeo AS GlsDpto, r.idProv, r.GlsUbigeo AS GlsProv, i.idDist, i.GlsUbigeo AS GlsDist " & _
                    "FROM datos p,ubigeo d,ubigeo r,ubigeo i " & _
                    "WHERE P.idDato = D.idPais " & _
                    "AND d.idPais = r.idPais " & _
                    "AND d.idDpto = r.idDpto " & _
                    "AND r.idPais = i.idPais " & _
                    "AND r.idDpto = i.idDpto " & _
                    "AND r.idProv = i.idProv " & _
                    "AND p.idtipoDatos = '02' " & _
                    "AND d.idProv = '00' AND d.idDist = '00' " & _
                    "AND r.idProv <> '00' AND r.idDist = '00' " & _
                    "AND i.idDist <> '00'"
    End Select
    
    If Trim(txt_TextoBuscar.Text) <> "" Then
        csql = csql & " AND " & strCampoMostrar & " LIKE '%" & Trim(txt_TextoBuscar.Text) & "%'"
    End If
    csql = csql & " ORDER BY " & strCampoMostrar

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
'        .KeyField = "Codigo"
'    End With
    
    
    Select Case cbx_TipoUbigeo.ListIndex
        Case 1 '--- Departamento
            gLista.Columns.ColumnByName("D1").GroupIndex = 0
        Case 2 '--- Provincia
            gLista.Columns.ColumnByName("D1").GroupIndex = 0
            gLista.Columns.ColumnByName("D2").GroupIndex = 1
        Case 3 '--- Distrito
            gLista.Columns.ColumnByName("D1").GroupIndex = 0
            gLista.Columns.ColumnByName("D2").GroupIndex = 1
            gLista.Columns.ColumnByName("D3").GroupIndex = 2
        End Select
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarUbigeo(strCod As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    nuevo
    fraUbigeo.Enabled = False
    
    If cbx_TipoUbigeo.ListIndex = 0 Then
        csql = "SELECT a.GlsDato " & _
               "FROM datos a " & _
               "WHERE a.idDato = '" & strCod & "'"
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        
        txtCodigo.Text = strCod
        txtGlosa.Text = "" & rst.Fields("GlsDato")
    
    Else
        csql = "SELECT a.GlsUbigeo,a.idZona " & _
               "FROM ubigeo a " & _
               "WHERE a.idDistrito = '" & strCod & "' AND a.idPais = '" & gLista.Columns.ColumnByFieldName("idPais").Value & "'"
        rst.Open csql, Cn, adOpenStatic, adLockReadOnly
        
        txtCodigo.Text = strCod
        txtCod_Pais.Text = gLista.Columns.ColumnByFieldName("idPais").Value
        txtCod_Departamento.Text = left(strCod, 2)
        txtCod_Provincia.Text = Mid(strCod, 3, 2)
        txtCod_Distrito.Text = right(strCod, 2)
        txtGlosa.Text = "" & rst.Fields("GlsUbigeo")
        txtCod_Zona.Text = "" & rst.Fields("idZona")
    End If
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    strCodigo = Trim(txtCodigo.Text)
    
    csql = "Select idDistrito From Personas Where idDistrito = '" & strCodigo & "' And idPais = '" & txtCod_Pais.Text & "' "
    If rsValida.State = 1 Then rsValida.Close
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Entidades)"
        GoTo Err
    End If
    
    csql = "Select idDistrito From tiendascliente WHERE idDistrito = '" & strCodigo & "' And idEmpresa = '" & glsEmpresa & "'"
    If rsValida.State = 1 Then rsValida.Close
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Tiendas Cliente)"
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando el registro
    csql = "Delete  From Ubigeo Where idDistrito = '" & strCodigo & "' And idPais = '" & Trim(txtCod_Pais.Text) & "' And idDpto='" & Trim(txtCod_Departamento.Text) & "'  And idDist ='" & Trim(txtCod_Distrito.Text) & "'"
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

Private Sub txtCod_Departamento_Change()

    txtGls_Departamento.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", txtCod_Departamento.Text, False, " idPais = '" & txtCod_Pais.Text & "' AND idProv = '00'")
    txtCod_Provincia.Text = ""
    txtGls_Provincia.Text = ""

End Sub

Private Sub txtCod_Departamento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DEPARTAMENTO", txtCod_Departamento, txtGls_Departamento, "AND idPais = '" & txtCod_Pais.Text & "'"
        KeyAscii = 0
        If txtCod_Departamento.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub


Private Sub txtCod_Pais_Change()
    
    txtGls_Pais.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_Pais.Text, False)
    txtCod_Departamento.Text = ""
    txtGls_Departamento.Text = ""
    txtCod_Provincia.Text = ""
    txtGls_Provincia.Text = ""
    
End Sub

Private Sub txtCod_Pais_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PAIS", txtCod_Pais, txtGls_Pais
        KeyAscii = 0
        If txtCod_Pais.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Provincia_Change()
    
    txtGls_Provincia.Text = traerCampo("ubigeo", "GlsUbigeo", "idProv", txtCod_Provincia.Text, False, " idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Departamento.Text & "' and idProv <> '00' and idDist = '00'")

End Sub

Private Sub txtCod_Provincia_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PROVINCIA", txtCod_Provincia, txtGls_Provincia, "AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Departamento.Text + "'"
        KeyAscii = 0
        If txtCod_Provincia.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Zona_Change()
    
    txtGls_Zona.Text = traerCampo("zonas", "GlsZona", "idZona", txtCod_Zona.Text, False)

End Sub

Private Sub txtCod_Zona_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "ZONAS", txtCod_Zona, txtGls_Zona
        KeyAscii = 0
        If txtCod_Zona.Text <> "" Then SendKeys "{tab}"
    End If

End Sub
