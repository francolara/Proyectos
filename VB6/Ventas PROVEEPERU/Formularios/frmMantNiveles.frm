VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantNiveles 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Niveles"
   ClientHeight    =   8790
   ClientLeft      =   4425
   ClientTop       =   2760
   ClientWidth     =   10950
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
   ScaleHeight     =   8790
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   9945
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":317E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":3518
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantNiveles.frx":3622
            Key             =   ""
         EndProperty
      EndProperty
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
      Height          =   8070
      Left            =   45
      TabIndex        =   15
      Top             =   645
      Width           =   10860
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
         Left            =   120
         TabIndex        =   16
         Top             =   150
         Width           =   10605
         Begin VB.ComboBox cbx_TipoNivel 
            Height          =   330
            ItemData        =   "frmMantNiveles.frx":3A74
            Left            =   8790
            List            =   "frmMantNiveles.frx":3A76
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   270
            Width           =   1665
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   945
            TabIndex        =   0
            Top             =   270
            Width           =   6720
            _ExtentX        =   11853
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
            Container       =   "frmMantNiveles.frx":3A78
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Tipo Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7980
            TabIndex        =   43
            Top             =   330
            Width           =   690
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Búsqueda"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   17
            Top             =   330
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   6870
         Left            =   120
         OleObjectBlob   =   "frmMantNiveles.frx":3A94
         TabIndex        =   2
         Top             =   1050
         Width           =   10650
      End
   End
   Begin VB.Frame fraGeneral 
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
      Height          =   8085
      Left            =   45
      TabIndex        =   11
      Top             =   630
      Width           =   10845
      Begin VB.Frame FraCtaContable 
         Appearance      =   0  'Flat
         Caption         =   " Cuentas Contables "
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   675
         TabIndex        =   51
         Top             =   4545
         Visible         =   0   'False
         Width           =   7890
         Begin VB.CommandButton CmdAyudaCtaContableVR 
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
            Left            =   7290
            Picture         =   "frmMantNiveles.frx":7736
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   1260
            Width           =   345
         End
         Begin VB.CommandButton CmdAyudaCtaContableV 
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
            Left            =   7290
            Picture         =   "frmMantNiveles.frx":7AC0
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   810
            Width           =   345
         End
         Begin VB.CommandButton CmdAyudaCtaContableC 
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
            Left            =   7290
            Picture         =   "frmMantNiveles.frx":7E4A
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   360
            Width           =   345
         End
         Begin CATControls.CATTextBox TxtCodCtaContableC 
            Height          =   315
            Left            =   1440
            TabIndex        =   53
            Tag             =   "TIdCtaContableC"
            Top             =   360
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
            Container       =   "frmMantNiveles.frx":81D4
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsCtaContableC 
            Height          =   315
            Left            =   2385
            TabIndex        =   54
            Top             =   360
            Width           =   4875
            _ExtentX        =   8599
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
            Container       =   "frmMantNiveles.frx":81F0
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtCodCtaContableV 
            Height          =   315
            Left            =   1440
            TabIndex        =   57
            Tag             =   "TIdCtaContableV"
            Top             =   810
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
            Container       =   "frmMantNiveles.frx":820C
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsCtaContableV 
            Height          =   315
            Left            =   2385
            TabIndex        =   58
            Top             =   810
            Width           =   4875
            _ExtentX        =   8599
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
            Container       =   "frmMantNiveles.frx":8228
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtCodCtaContableVR 
            Height          =   315
            Left            =   1440
            TabIndex        =   61
            Tag             =   "TIdCtaContableVR"
            Top             =   1260
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
            Container       =   "frmMantNiveles.frx":8244
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsCtaContableVR 
            Height          =   315
            Left            =   2385
            TabIndex        =   62
            Top             =   1260
            Width           =   4875
            _ExtentX        =   8599
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
            Container       =   "frmMantNiveles.frx":8260
            Vacio           =   -1  'True
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "V. Relacionada"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   270
            TabIndex        =   63
            Top             =   1305
            Width           =   1095
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Ventas"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   270
            TabIndex        =   59
            Top             =   855
            Width           =   525
         End
         Begin VB.Label LblCtaContable 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Compras"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   270
            TabIndex        =   55
            Top             =   405
            Width           =   645
         End
      End
      Begin VB.CommandButton cmbAyudaTiposNivel 
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
         Left            =   8160
         Picture         =   "frmMantNiveles.frx":827C
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2755
         Visible         =   0   'False
         Width           =   390
      End
      Begin CATControls.CATTextBox txtGls_Nivel 
         Height          =   315
         Left            =   1935
         TabIndex        =   5
         Tag             =   "TglsNivel"
         Top             =   1980
         Width           =   6615
         _ExtentX        =   11668
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
         Container       =   "frmMantNiveles.frx":8606
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Frame fraNivel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Height          =   450
         Left            =   450
         TabIndex        =   23
         Top             =   1485
         Width           =   8145
         Begin VB.CommandButton cmbAyudaNivelPred 
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
            Index           =   0
            Left            =   7710
            Picture         =   "frmMantNiveles.frx":8622
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   90
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivelPred 
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
            Index           =   1
            Left            =   7710
            Picture         =   "frmMantNiveles.frx":89AC
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   405
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivelPred 
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
            Index           =   2
            Left            =   7710
            Picture         =   "frmMantNiveles.frx":8D36
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   720
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivelPred 
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
            Index           =   3
            Left            =   7710
            Picture         =   "frmMantNiveles.frx":90C0
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1080
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivelPred 
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
            Index           =   4
            Left            =   7710
            Picture         =   "frmMantNiveles.frx":944A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1440
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_NivelPred 
            Height          =   315
            Index           =   0
            Left            =   1485
            TabIndex        =   4
            Top             =   85
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
            Container       =   "frmMantNiveles.frx":97D4
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_NivelPred 
            Height          =   315
            Index           =   0
            Left            =   2430
            TabIndex        =   29
            Top             =   85
            Width           =   5250
            _ExtentX        =   9260
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
            Container       =   "frmMantNiveles.frx":97F0
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_NivelPred 
            Height          =   315
            Index           =   1
            Left            =   2430
            TabIndex        =   30
            Top             =   435
            Width           =   5250
            _ExtentX        =   9260
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
            Container       =   "frmMantNiveles.frx":980C
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_NivelPred 
            Height          =   315
            Index           =   2
            Left            =   1485
            TabIndex        =   31
            Top             =   750
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
            Container       =   "frmMantNiveles.frx":9828
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_NivelPred 
            Height          =   315
            Index           =   2
            Left            =   2430
            TabIndex        =   32
            Top             =   750
            Width           =   5250
            _ExtentX        =   9260
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
            Container       =   "frmMantNiveles.frx":9844
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_NivelPred 
            Height          =   315
            Index           =   3
            Left            =   1485
            TabIndex        =   33
            Top             =   1110
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
            Container       =   "frmMantNiveles.frx":9860
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_NivelPred 
            Height          =   315
            Index           =   3
            Left            =   2430
            TabIndex        =   34
            Top             =   1110
            Width           =   5250
            _ExtentX        =   9260
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
            Container       =   "frmMantNiveles.frx":987C
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_NivelPred 
            Height          =   315
            Index           =   4
            Left            =   1485
            TabIndex        =   35
            Top             =   1470
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
            Container       =   "frmMantNiveles.frx":9898
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_NivelPred 
            Height          =   315
            Index           =   4
            Left            =   2430
            TabIndex        =   36
            Top             =   1470
            Width           =   5250
            _ExtentX        =   9260
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
            Container       =   "frmMantNiveles.frx":98B4
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_NivelPred 
            Height          =   315
            Index           =   1
            Left            =   1485
            TabIndex        =   42
            Top             =   435
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
            Container       =   "frmMantNiveles.frx":98D0
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label lblNivelPred 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   270
            TabIndex        =   41
            Top             =   135
            Width           =   345
         End
         Begin VB.Label lblNivelPred 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   270
            TabIndex        =   40
            Top             =   405
            Width           =   390
         End
         Begin VB.Label lblNivelPred 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   270
            TabIndex        =   39
            Top             =   765
            Width           =   390
         End
         Begin VB.Label lblNivelPred 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   270
            TabIndex        =   38
            Top             =   1125
            Width           =   390
         End
         Begin VB.Label lblNivelPred 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   270
            TabIndex        =   37
            Top             =   1485
            Width           =   390
         End
      End
      Begin VB.Frame fraNivelPred 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Height          =   435
         Left            =   1890
         TabIndex        =   22
         Top             =   1485
         Width           =   6480
      End
      Begin VB.CommandButton cmbAyudaTipoNivel 
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
         Left            =   8160
         Picture         =   "frmMantNiveles.frx":98EC
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1110
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Nivel 
         Height          =   315
         Left            =   7635
         TabIndex        =   12
         Tag             =   "TidNivel"
         Top             =   360
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
         Container       =   "frmMantNiveles.frx":9C76
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_TipoNivel 
         Height          =   315
         Left            =   1935
         TabIndex        =   3
         Tag             =   "TidTipoNivel"
         Top             =   1080
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
         Container       =   "frmMantNiveles.frx":9C92
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TipoNivel 
         Height          =   315
         Left            =   2880
         TabIndex        =   20
         Top             =   1110
         Width           =   5250
         _ExtentX        =   9260
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
         Container       =   "frmMantNiveles.frx":9CAE
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Orden 
         Height          =   315
         Left            =   1935
         TabIndex        =   6
         Tag             =   "NOrden"
         Top             =   2385
         Visible         =   0   'False
         Width           =   930
         _ExtentX        =   1640
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
         Container       =   "frmMantNiveles.frx":9CCA
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_TiposNivel 
         Height          =   315
         Left            =   1935
         TabIndex        =   7
         Tag             =   "TTipo"
         Top             =   2760
         Visible         =   0   'False
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
         Container       =   "frmMantNiveles.frx":9CE6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TiposNivel 
         Height          =   315
         Left            =   2880
         TabIndex        =   47
         Top             =   2760
         Visible         =   0   'False
         Width           =   5250
         _ExtentX        =   9260
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
         Container       =   "frmMantNiveles.frx":9D02
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox Txt_PorcComision 
         Height          =   315
         Left            =   1935
         TabIndex        =   8
         Tag             =   "NPorcComision"
         Top             =   3150
         Visible         =   0   'False
         Width           =   930
         _ExtentX        =   1640
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
         Container       =   "frmMantNiveles.frx":9D1E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtDctoListaPrec 
         Height          =   315
         Left            =   1935
         TabIndex        =   9
         Tag             =   "NMaxDcto"
         Top             =   3555
         Visible         =   0   'False
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "frmMantNiveles.frx":9D3A
         Estilo          =   3
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGls_Abreviatura 
         Height          =   315
         Left            =   1935
         TabIndex        =   10
         Tag             =   "TGlsAbreviatura"
         Top             =   4005
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
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
         Container       =   "frmMantNiveles.frx":9D56
         Vacio           =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Abreviatura"
         Height          =   240
         Left            =   720
         TabIndex        =   50
         Top             =   4095
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Max % Dscto."
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   720
         TabIndex        =   49
         Top             =   3600
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Lbl_PorcComision 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "% Comisión"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   720
         TabIndex        =   48
         Top             =   3135
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblTiposNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   720
         TabIndex        =   45
         Top             =   2775
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label LblOrden 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Orden"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   720
         TabIndex        =   44
         Top             =   2415
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo Nivel"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   720
         TabIndex        =   21
         Top             =   1155
         Width           =   690
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6990
         TabIndex        =   14
         Top             =   390
         Width           =   495
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   720
         TabIndex        =   13
         Top             =   2010
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
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
Attribute VB_Name = "frmMantNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim indCargando As Boolean

Private Sub cbx_TipoNivel_Click()
On Error GoTo Err
Dim StrMsgError As String

    If indCargando Then Exit Sub
    listaNivel StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaNivelPred_Click(Index As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
    
    peso = Index + 1
    strCodTipoNivel = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
    strCondPred = ""
    If peso > 1 Then
        strCondPred = " AND idNivelPred = '" & txtCod_NivelPred(Index - 1).Text & "'"
    End If
    mostrarAyuda "NIVEL", txtCod_NivelPred(Index), txtGls_NivelPred(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred

End Sub

Private Sub cmbAyudaTiposNivel_Click()
    
    mostrarAyuda "TIPOSNIVEL", txtCod_TiposNivel, txtGls_TiposNivel

End Sub

Private Sub CmdAyudaCtaContableC_Click()
On Error GoTo Err
Dim StrMsgError                             As String
    
    Ayuda_PlanContable StrMsgError, TxtCodCtaContableC
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableV_Click()
On Error GoTo Err
Dim StrMsgError                             As String
    
    Ayuda_PlanContable StrMsgError, TxtCodCtaContableV
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVR_Click()
On Error GoTo Err
Dim StrMsgError                             As String
    
    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVR
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    Me.top = 0
    Me.left = 0
    indCargando = True
    
    ConfGrid gLista, False, False, False, False
    llenaCombo cbx_TipoNivel, "tiposniveles", "GlsTipoNIvel", False, "idTipoNivel", " idEmpresa = '" & glsEmpresa & "'"
    If cbx_TipoNivel.ListCount > 0 Then cbx_TipoNivel.ListIndex = cbx_TipoNivel.ListCount - 1
    
    listaNivel StrMsgError
    If StrMsgError <> "" Then GoTo Err
    indCargando = False
    
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
Dim strCodigo                           As String
Dim strMsg                              As String
Dim CSqlC                               As String
Dim NPesoNivel                          As Integer
Dim cWhereNiveles                       As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    'validaHomonimia "niveles", "GlsNivel", "idNivel", txtGls_Nivel.Text, txtCod_Nivel.Text, True, StrMsgError, " idTipoNivel = '" & txtCod_TipoNivel.Text & "'"
    'If StrMsgError <> "" Then GoTo Err
    
    If txtCod_Nivel.Text = "" Then
        txtCod_Nivel.Text = GeneraCorrelativoAnoMes("niveles", "idNivel")
        EjecutaSQLForm Me, 0, True, "niveles", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabó"
        
    Else
        EjecutaSQLForm Me, 1, True, "niveles", StrMsgError, "idNivel"
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modificó"
    End If
    
    If traerCampo("Parametros", "ValParametro", "GlsParametro", "ACTUALIZA_DESCUENTO_NIVEL", True) = "1" Then
        cWhereNiveles = ""
        If Len(Trim(txtCod_NivelPred(0).Text)) > 0 Then
            cWhereNiveles = cWhereNiveles & "And N.IdNivel" & Format(glsNumNiveles, "00") & " = '" & txtCod_NivelPred(0).Text & "' "
            If Len(Trim(txtCod_NivelPred(1).Text)) > 0 Then
                cWhereNiveles = cWhereNiveles & "And N.IdNivel" & Format(glsNumNiveles - 1, "00") & " = '" & txtCod_NivelPred(1).Text & "' "
            End If
        End If
        NPesoNivel = Val("" & traerCampo("TiposNiveles", "Peso", "IdTipoNivel", txtCod_TipoNivel.Text, True))
        cWhereNiveles = cWhereNiveles & "And N.IdNivel" & Format(IIf(NPesoNivel = 1, "3", IIf(NPesoNivel = 3, "1", "2")), "00") & " = '" & txtCod_Nivel.Text & "' "
        CSqlC = "Update Vw_Niveles N " & _
                "Inner Join Productos A " & _
                "On N.IdEmpresa = A.IdEmpresa And N.IdNivel01 = A.IdNivel " & _
                "Inner Join PreciosVenta B " & _
                "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                "Set B.MaxDcto = " & Val("" & TxtDctoListaPrec.Text) & " " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' " & cWhereNiveles
        Cn.Execute (CSqlC)
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    listaNivel StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()
Dim strTipoNivelSel As String
Dim strPesoTipoNivelSel As String

    limpiaForm Me
    txtCod_TipoNivel.Text = right(cbx_TipoNivel.Text, 8)
    strTipoNivelSel = right(cbx_TipoNivel.Text, 8)
    strPesoTipoNivelSel = traerCampo("tiposniveles", "peso", "idTipoNivel", strTipoNivelSel, True)
    txt_Orden.Text = 0
    Txt_PorcComision.Text = 0
    TxtDctoListaPrec.Text = 0
    
    If strPesoTipoNivelSel = "2" Then
        LblOrden.Visible = True
        txt_Orden.Visible = True
        lblTiposNivel.Visible = False
        txtCod_TiposNivel.Visible = False
        txtGls_TiposNivel.Visible = False
        cmbAyudaTiposNivel.Visible = False
    Else
        LblOrden.Visible = False
        txt_Orden.Visible = False
        lblTiposNivel.Visible = False
        txtCod_TiposNivel.Visible = False
        txtGls_TiposNivel.Visible = False
        cmbAyudaTiposNivel.Visible = False
    End If

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String
Dim strTipoNivelSel As String
Dim strPesoTipoNivelSel As String

    mostrarNivel gLista.Columns.ColumnByName("idNivel").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    strTipoNivelSel = right(cbx_TipoNivel.Text, 8)
    strPesoTipoNivelSel = traerCampo("tiposniveles", "peso", "idTipoNivel", strTipoNivelSel, True)
    If strPesoTipoNivelSel = "2" Then
        LblOrden.Visible = True
        txt_Orden.Visible = True
        lblTiposNivel.Visible = False
        txtCod_TiposNivel.Visible = False
        txtGls_TiposNivel.Visible = False
        cmbAyudaTiposNivel.Visible = False
    Else
        LblOrden.Visible = False
        txt_Orden.Visible = False
        lblTiposNivel.Visible = False
        txtCod_TiposNivel.Visible = False
        txtGls_TiposNivel.Visible = False
        cmbAyudaTiposNivel.Visible = False
    End If
    If strPesoTipoNivelSel = glsNumNiveles Then
        Lbl_PorcComision.Visible = False
        Txt_PorcComision.Visible = False
    Else
        Lbl_PorcComision.Visible = False
        Txt_PorcComision.Visible = False
    End If
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
            
            listaNivel StrMsgError
            If StrMsgError <> "" Then GoTo Err
    
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 6 'Imprimir
            gLista.m.ExportToXLS App.Path & "\Temporales\Mantenimiento_Niveles.xls"
            ShellEx App.Path & "\Temporales\Mantenimiento_Niveles.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
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
            Toolbar1.Buttons(6).Visible = Not indHabilitar 'Imprimir
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

    listaNivel StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaNivel(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim strCond As String
Dim intNumNiveles As Integer
Dim strTabla As String
Dim strWhere As String
Dim strCampos As String
Dim strTablas As String
Dim strTablaAnt As String
Dim strTipoNivelSel As String
Dim strPesoTipoNivelSel As String
Dim i As Integer
Dim rsdatos                     As New ADODB.Recordset
 
    strTipoNivelSel = right(cbx_TipoNivel.Text, 8)
    strPesoTipoNivelSel = traerCampo("tiposniveles", "peso", "idTipoNivel", strTipoNivelSel, True)

    rst.Open "SELECT idTipoNivel,GlsTipoNivel,peso FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' AND peso <= " & strPesoTipoNivelSel & " ORDER BY peso DESC", Cn, adOpenKeyset, adLockOptimistic
    intNumNiveles = rst.RecordCount - 1
    rst.MoveNext
    
    Do While Not rst.EOF
        i = i + 1
        strTabla = "niveles" & Format(i, "00")
        If i = 1 Then
            strWhere = "n.idNivelPred = " & strTabla & ".idNivel AND " & strTabla & ".idEmpresa = '" & glsEmpresa & "' AND "
        Else
            strWhere = strWhere & strTablaAnt & ".idNivelPred = " & strTabla & ".idNivel AND " & strTabla & ".idEmpresa = '" & glsEmpresa & "' AND "
        End If
        strCampos = strCampos & strTabla & ".idNivel as idNivel" & Format(i, "00") & "," & strTabla & ".GlsNivel as GlsNivel" & Format(i, "00") & ","
        
        strTablas = strTablas & ",niveles " & strTabla
        strTablaAnt = strTabla
        
        rst.MoveNext
    Loop

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND n.GlsNivel LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT " & strCampos & " n.idNivel ,n.GlsNivel " & _
           "FROM niveles n " & strTablas & " WHERE " & strWhere & " n.idEmpresa = '" & glsEmpresa & "' AND n.idTipoNivel = '" & strTipoNivelSel & "'"
    If strCond <> "" Then csql = csql + strCond
    csql = csql & " ORDER BY n.idNivel"
    
    
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
'        .KeyField = "idNivel"
'    End With


    '--- Realizando las agrupaciones dinamicamente
    gLista.m.ClearGroupColumns
    
    For i = 1 To 4
         gLista.Columns.ColumnByName("GlsNivel" & Format(i, "00")).Visible = False
    Next
    
    If gLista.Ex.GroupColumnCount = 0 Then
        For i = 1 To intNumNiveles
            gLista.Columns.ColumnByName("GlsNivel" & Format(i, "00")).Caption = "Nivel:"
            gLista.Columns.ColumnByName("GlsNivel" & Format(i, "00")).Visible = True
            gLista.Columns.ColumnByName("GlsNivel" & Format(i, "00")).GroupIndex = intNumNiveles - i
        Next
    End If
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarNivel(strCodNivel As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst                     As New ADODB.Recordset
Dim peso                    As Integer
Dim i                       As Long
    
    csql = "Select N.IdNivel,N.GlsNivel,N.IdTipoNivel,N.IdNivelPred,isnull(N.Orden,0) Orden,N.Tipo,N.PorcComision,N.MaxDcto,N.GlsAbreviatura,N.IdCtaContableC," & _
           "N.IdCtaContableV,N.IdCtaContableVR " & _
           "From Niveles N " & _
           "Where N.IdEmpresa = '" & glsEmpresa & "' And N.IdNivel = '" & strCodNivel & "'"
    
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
    
        txtCod_Nivel.Text = "" & rst.Fields("idNivel")
        txtCod_TipoNivel.Text = "" & rst.Fields("idTipoNivel")
        txtGls_Nivel.Text = "" & rst.Fields("GlsNivel")
        txt_Orden.Text = "" & rst.Fields("Orden")
        txtCod_TiposNivel.Text = "" & rst.Fields("Tipo")
        Txt_PorcComision.Text = "" & rst.Fields("PorcComision")
        TxtDctoListaPrec.Text = "" & rst.Fields("MaxDcto")
        peso = pesoTipoNivel(txtCod_TipoNivel.Text)
        TxtGls_Abreviatura.Text = Trim("" & rst.Fields("GlsAbreviatura"))
        
        If peso > 1 Then
            txtCod_NivelPred(peso - 2).Text = "" & rst.Fields("idNivelPred")
            For i = 0 To peso - 3
                txtCod_NivelPred(i).Text = traerCampo("niveles", "IdNivelPred", "IdNivel", txtCod_NivelPred(i + 1).Text, True)
            Next
        End If
    
        TxtCodCtaContableC.Text = "" & rst.Fields("IdCtaContableC")
        TxtCodCtaContableV.Text = "" & rst.Fields("IdCtaContableV")
        TxtCodCtaContableVR.Text = "" & rst.Fields("IdCtaContableVR")
        
    End If
    
    rst.Close: Set rst = Nothing
    
    Me.Refresh
    
    Exit Sub
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_TipoNivel_Change()
On Error GoTo Err
Dim StrMsgError                         As String
Dim peso                                As Integer
Dim strCodTipoNivel                     As String
    
    txtGls_TipoNivel.Text = traerCampo("tiposniveles", "GlsTipoNivel", "idTipoNivel", txtCod_TipoNivel.Text, True)
    
    mostrarNivelesPred
    
    If Val(traerCampo("tiposniveles", "peso", "idTipoNivel", txtCod_TipoNivel.Text, True)) = glsNumNiveles Then
        Lbl_PorcComision.Visible = False
        Txt_PorcComision.Visible = False
    Else
        Lbl_PorcComision.Visible = False
        Txt_PorcComision.Visible = False
    End If
    
    If leeParametro("NIVEL_CUENTA_CONTABLE") = txtCod_TipoNivel.Text Then
    
        FraCtaContable.Visible = True
    
    Else
        
        FraCtaContable.Visible = False
        
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_TipoNivel_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "TIPONIVEL", txtCod_TipoNivel, txtGls_TipoNivel
        KeyAscii = 0
        If txtCod_TipoNivel.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub cmbAyudaTipoNivel_Click()
    
    mostrarAyuda "TIPONIVEL", txtCod_TipoNivel, txtGls_TipoNivel

End Sub

Private Function pesoTipoNivel(strCodTipoNivel As String) As Integer
    
    pesoTipoNivel = Val("" & traerCampo("tiposniveles", "Peso", "idTipoNivel", strCodTipoNivel, True))

End Function

Private Sub mostrarNivelesPred()
Dim rsj As New ADODB.Recordset
Dim numPesos As Integer
Dim peso As Integer
Dim i As Integer

    For i = 0 To 4
        txtCod_NivelPred(i).Tag = ""
    Next
    
    peso = pesoTipoNivel(txtCod_TipoNivel.Text)
    
    '--- Tipo niveles
    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE peso < " & CStr(peso) & " AND idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    numPesos = Val("" & rsj.RecordCount)
    fraNivel.Height = 375 * numPesos
    i = 0
    
    Do While Not rsj.EOF
        If (i + 1) = numPesos Then
            txtCod_NivelPred(i).Tag = "TidNivelPred"
        End If
        lblNivelPred(i).Caption = "" & rsj.Fields("GlsTipoNivel")
        rsj.MoveNext
        i = i + 1
    Loop
    
    If peso = 1 Then
        txtGls_Nivel.top = txtGls_TipoNivel.top + txtGls_TipoNivel.Height + 160
        lblNivel.top = txtGls_TipoNivel.top + txtGls_TipoNivel.Height + 160
        txtCod_NivelPred(4).Tag = "TidNivelPred"
    Else
        txtGls_Nivel.top = fraNivel.top + fraNivel.Height + 70
        lblNivel.top = fraNivel.top + fraNivel.Height + 70
    End If

End Sub

Private Sub txtCod_NivelPred_Change(Index As Integer)
    
    txtGls_NivelPred(Index).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_NivelPred(Index).Text, True)

End Sub

Private Sub txtCod_NivelPred_KeyPress(Index As Integer, KeyAscii As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
    
    If KeyAscii <> 13 Then
        peso = Index + 1
        strCodTipoNivel = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
        strCondPred = ""
        
        If peso > 1 Then
            strCondPred = " AND idNivelPred = '" & txtCod_NivelPred(Index - 1).Text & "'"
        End If
                    
        mostrarAyudaKeyascii KeyAscii, "NIVEL", txtCod_NivelPred(Index), txtGls_NivelPred(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred
    End If

End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("Está seguro(a) de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    strCodigo = Trim(txtCod_Nivel.Text)
    
    csql = "SELECT idNivelPred FROM niveles WHERE idNivelPred = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Nivel Predecesor)."
        GoTo Err
    End If
    
    csql = "SELECT idNivel FROM productos WHERE idNivel = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    If rsValida.State = 1 Then rsValida.Close
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Productos)."
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando el registro
    csql = "DELETE FROM niveles WHERE idNivel = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
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

Private Sub txtCod_TiposNivel_Change()
    
    If txtCod_TiposNivel.Text = "" Then
        txtGls_TiposNivel.Text = ""
    Else
        txtGls_TiposNivel.Text = traerCampo("datos", "glsDato", "idDato", txtCod_TiposNivel.Text, False, "idTipoDatos = '20' ")
    End If
    
End Sub

Private Sub TxtCodCtaContableC_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableC.Text <> "" Then
        TxtGlsCtaContableC.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableC.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableC.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableC_GotFocus()
On Error GoTo Err
Dim StrMsgError                             As String
    
    Foco StrMsgError, TxtCodCtaContableC
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableC_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableC
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableC_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyAscii = 13 Then
    
        TxtCodCtaContableV.SetFocus
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableV_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableV.Text <> "" Then
        TxtGlsCtaContableV.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableV.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableV.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableV_GotFocus()
On Error GoTo Err
Dim StrMsgError                             As String
    
    Foco StrMsgError, TxtCodCtaContableV
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableV_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableV
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableV_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyAscii = 13 Then
    
        TxtCodCtaContableVR.SetFocus
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVR_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVR.Text <> "" Then
        TxtGlsCtaContableVR.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVR.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVR.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVR_GotFocus()
On Error GoTo Err
Dim StrMsgError                             As String
    
    Foco StrMsgError, TxtCodCtaContableVR
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVR_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableVR
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVR_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyAscii = 13 Then
    
        'TxtCodCtaContableVR.SetFocus
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Foco(StrMsgError As String, PTxt As Object)
On Error GoTo Err
    
    PTxt.SelStart = 0: PTxt.SelLength = Len(PTxt.Text)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Ayuda_PlanContable(StrMsgError As String, txtCod As Object)
On Error GoTo Err
Dim Cod                                     As String
    
    mostrarAyudaTextoPlanCuentas strcnConta, "PLANCUENTAS", Cod, "", "", "2011"
    If StrMsgError <> "" Then GoTo Err
    
    If Len(Trim("" & Cod)) > 0 Then txtCod.Text = Cod
    
    Select Case txtCod.Name
        Case "TxtCodCtaContableC"
            TxtCodCtaContableV.SetFocus
        
        Case "TxtCodCtaContableV"
            TxtCodCtaContableVR.SetFocus
    
    End Select
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
