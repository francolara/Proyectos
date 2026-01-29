VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmMantDocumentos 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos"
   ClientHeight    =   7095
   ClientLeft      =   1110
   ClientTop       =   720
   ClientWidth     =   15195
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
   ScaleHeight     =   7095
   ScaleWidth      =   15195
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   8190
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
            Picture         =   "FrmMantDocumentos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantDocumentos.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1164
      ButtonWidth     =   2461
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
   Begin VB.Frame FraListado 
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
      Height          =   6435
      Left            =   45
      TabIndex        =   3
      Top             =   630
      Width           =   15075
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
         Height          =   660
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   14895
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   990
            TabIndex        =   0
            Top             =   210
            Width           =   13800
            _ExtentX        =   24342
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
            Container       =   "FrmMantDocumentos.frx":3518
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
            Left            =   135
            TabIndex        =   5
            Top             =   270
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   5460
         Left            =   90
         OleObjectBlob   =   "FrmMantDocumentos.frx":3534
         TabIndex        =   1
         Top             =   915
         Width           =   14895
      End
   End
   Begin VB.Frame FraGeneral 
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
      Height          =   6435
      Left            =   45
      TabIndex        =   6
      Top             =   630
      Width           =   15075
      Begin VB.ComboBox CmbEstado 
         Height          =   330
         ItemData        =   "FrmMantDocumentos.frx":5131
         Left            =   13095
         List            =   "FrmMantDocumentos.frx":5133
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1125
         Width           =   1770
      End
      Begin CATControls.CATTextBox TxtCodDocumento 
         Height          =   315
         Left            =   13815
         TabIndex        =   8
         Tag             =   "TidCodClasActivos"
         Top             =   225
         Width           =   1005
         _ExtentX        =   1773
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
         MaxLength       =   2
         Container       =   "FrmMantDocumentos.frx":5135
      End
      Begin CATControls.CATTextBox TxtGlsDocumento 
         Height          =   315
         Left            =   1710
         TabIndex        =   2
         Tag             =   "TglsClaActivos"
         Top             =   675
         Width           =   6540
         _ExtentX        =   11536
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
         Container       =   "FrmMantDocumentos.frx":5151
         Estilo          =   1
      End
      Begin CATControls.CATTextBox TxtAbreviatura 
         Height          =   315
         Left            =   1710
         TabIndex        =   13
         Tag             =   "TAbreviatura"
         Top             =   1095
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
         MaxLength       =   3
         Container       =   "FrmMantDocumentos.frx":516D
         EnterTab        =   -1  'True
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4605
         Left            =   225
         TabIndex        =   16
         Top             =   1620
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   8123
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Cuentas Contables Ventas"
         TabPicture(0)   =   "FrmMantDocumentos.frx":5189
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame5"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame10"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame12"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Cuentas Contables Compras"
         TabPicture(1)   =   "FrmMantDocumentos.frx":51A5
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame13"
         Tab(1).Control(1)=   "Frame11"
         Tab(1).Control(2)=   "Frame9"
         Tab(1).Control(3)=   "Frame8"
         Tab(1).Control(4)=   "Frame7"
         Tab(1).Control(5)=   "Frame6"
         Tab(1).ControlCount=   6
         Begin VB.Frame Frame13 
            Appearance      =   0  'Flat
            Caption         =   " Socios "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   -67620
            TabIndex        =   116
            Top             =   3105
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableCSocS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":51C1
               Style           =   1  'Graphical
               TabIndex        =   118
               Top             =   315
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableCSocD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":554B
               Style           =   1  'Graphical
               TabIndex        =   117
               Top             =   720
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCSocS 
               Height          =   315
               Left            =   1170
               TabIndex        =   119
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":58D5
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCSocS 
               Height          =   315
               Left            =   2205
               TabIndex        =   120
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":58F1
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCSocD 
               Height          =   315
               Left            =   1170
               TabIndex        =   121
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":590D
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCSocD 
               Height          =   315
               Left            =   2205
               TabIndex        =   122
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":5929
               Vacio           =   -1  'True
            End
            Begin VB.Label Label30 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   124
               Top             =   360
               Width           =   405
            End
            Begin VB.Label Label29 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   123
               Top             =   765
               Width           =   555
            End
         End
         Begin VB.Frame Frame12 
            Appearance      =   0  'Flat
            Caption         =   " Socios "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   7380
            TabIndex        =   107
            Top             =   3105
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableVSocD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":5945
               Style           =   1  'Graphical
               TabIndex        =   109
               Top             =   720
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableVSocS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":5CCF
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   315
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVSocS 
               Height          =   315
               Left            =   1170
               TabIndex        =   110
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":6059
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVSocS 
               Height          =   315
               Left            =   2205
               TabIndex        =   111
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":6075
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVSocD 
               Height          =   315
               Left            =   1170
               TabIndex        =   112
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":6091
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVSocD 
               Height          =   315
               Left            =   2205
               TabIndex        =   113
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":60AD
               Vacio           =   -1  'True
            End
            Begin VB.Label Label28 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   115
               Top             =   765
               Width           =   555
            End
            Begin VB.Label Label27 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   114
               Top             =   360
               Width           =   405
            End
         End
         Begin VB.Frame Frame11 
            Appearance      =   0  'Flat
            Caption         =   " Personal "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   -74820
            TabIndex        =   98
            Top             =   3105
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableCPerD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":60C9
               Style           =   1  'Graphical
               TabIndex        =   100
               Top             =   720
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableCPerS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":6453
               Style           =   1  'Graphical
               TabIndex        =   99
               Top             =   315
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCPerS 
               Height          =   315
               Left            =   1170
               TabIndex        =   101
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":67DD
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCPerS 
               Height          =   315
               Left            =   2205
               TabIndex        =   102
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":67F9
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCPerD 
               Height          =   315
               Left            =   1170
               TabIndex        =   103
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":6815
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCPerD 
               Height          =   315
               Left            =   2205
               TabIndex        =   104
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":6831
               Vacio           =   -1  'True
            End
            Begin VB.Label Label26 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   106
               Top             =   765
               Width           =   555
            End
            Begin VB.Label Label25 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   105
               Top             =   360
               Width           =   405
            End
         End
         Begin VB.Frame Frame10 
            Appearance      =   0  'Flat
            Caption         =   " Personal "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   180
            TabIndex        =   89
            Top             =   3105
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableVPerS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":684D
               Style           =   1  'Graphical
               TabIndex        =   91
               Top             =   315
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableVPerD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":6BD7
               Style           =   1  'Graphical
               TabIndex        =   90
               Top             =   720
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVPerS 
               Height          =   315
               Left            =   1170
               TabIndex        =   92
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":6F61
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVPerS 
               Height          =   315
               Left            =   2205
               TabIndex        =   93
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":6F7D
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVPerD 
               Height          =   315
               Left            =   1170
               TabIndex        =   94
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":6F99
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVPerD 
               Height          =   315
               Left            =   2205
               TabIndex        =   95
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":6FB5
               Vacio           =   -1  'True
            End
            Begin VB.Label Label24 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   97
               Top             =   360
               Width           =   405
            End
            Begin VB.Label Label23 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   96
               Top             =   765
               Width           =   555
            End
         End
         Begin VB.Frame Frame9 
            Appearance      =   0  'Flat
            Caption         =   " Relacionada - Matriz "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   -67620
            TabIndex        =   80
            Top             =   405
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableCRelMD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":6FD1
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   720
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableCRelMS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":735B
               Style           =   1  'Graphical
               TabIndex        =   81
               Top             =   315
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCRelMS 
               Height          =   315
               Left            =   1170
               TabIndex        =   83
               Tag             =   "TidCtaContableD"
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":76E5
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCRelMS 
               Height          =   315
               Left            =   2205
               TabIndex        =   84
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":7701
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCRelMD 
               Height          =   315
               Left            =   1170
               TabIndex        =   85
               Tag             =   "TidCtaContableH"
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":771D
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCRelMD 
               Height          =   315
               Left            =   2205
               TabIndex        =   86
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":7739
               Vacio           =   -1  'True
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   88
               Top             =   765
               Width           =   555
            End
            Begin VB.Label Label21 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   87
               Top             =   360
               Width           =   405
            End
         End
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            Caption         =   " Terceros "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   -74820
            TabIndex        =   71
            Top             =   405
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableCTerD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":7755
               Style           =   1  'Graphical
               TabIndex        =   73
               Top             =   720
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableCTerS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":7ADF
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   315
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCTerS 
               Height          =   315
               Left            =   1170
               TabIndex        =   74
               Tag             =   "TidCtaContableD"
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":7E69
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCTerS 
               Height          =   315
               Left            =   2205
               TabIndex        =   75
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":7E85
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCTerD 
               Height          =   315
               Left            =   1170
               TabIndex        =   76
               Tag             =   "TidCtaContableH"
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":7EA1
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCTerD 
               Height          =   315
               Left            =   2205
               TabIndex        =   77
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":7EBD
               Vacio           =   -1  'True
            End
            Begin VB.Label Label20 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   79
               Top             =   765
               Width           =   555
            End
            Begin VB.Label Label19 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   78
               Top             =   360
               Width           =   405
            End
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            Caption         =   " Relacionada - SubSidiaria "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   -74820
            TabIndex        =   62
            Top             =   1755
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableCRelSS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":7ED9
               Style           =   1  'Graphical
               TabIndex        =   64
               Top             =   315
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableCRelSD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":8263
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   720
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCRelSS 
               Height          =   315
               Left            =   1170
               TabIndex        =   65
               Tag             =   "TidCtaContableD"
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":85ED
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCRelSS 
               Height          =   315
               Left            =   2205
               TabIndex        =   66
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":8609
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCRelSD 
               Height          =   315
               Left            =   1170
               TabIndex        =   67
               Tag             =   "TidCtaContableH"
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":8625
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCRelSD 
               Height          =   315
               Left            =   2205
               TabIndex        =   68
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":8641
               Vacio           =   -1  'True
            End
            Begin VB.Label Label18 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   70
               Top             =   360
               Width           =   405
            End
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   69
               Top             =   765
               Width           =   555
            End
         End
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            Caption         =   " Relacionada - Asociada "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   -67620
            TabIndex        =   53
            Top             =   1755
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableCRelAS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":865D
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   315
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableCRelAD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":89E7
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   720
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCRelAS 
               Height          =   315
               Left            =   1170
               TabIndex        =   56
               Tag             =   "TidCtaContableD"
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":8D71
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCRelAS 
               Height          =   315
               Left            =   2205
               TabIndex        =   57
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":8D8D
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableCRelAD 
               Height          =   315
               Left            =   1170
               TabIndex        =   58
               Tag             =   "TidCtaContableH"
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":8DA9
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableCRelAD 
               Height          =   315
               Left            =   2205
               TabIndex        =   59
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":8DC5
               Vacio           =   -1  'True
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   61
               Top             =   360
               Width           =   405
            End
            Begin VB.Label Label15 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   60
               Top             =   765
               Width           =   555
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            Caption         =   " Relacionada - Asociada "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   7380
            TabIndex        =   44
            Top             =   1755
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableVRelAD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":8DE1
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   720
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableVRelAS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":916B
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   315
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVRelAS 
               Height          =   315
               Left            =   1170
               TabIndex        =   47
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":94F5
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVRelAS 
               Height          =   315
               Left            =   2205
               TabIndex        =   48
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":9511
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVRelAD 
               Height          =   315
               Left            =   1170
               TabIndex        =   49
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":952D
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVRelAD 
               Height          =   315
               Left            =   2205
               TabIndex        =   50
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":9549
               Vacio           =   -1  'True
            End
            Begin VB.Label Label14 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   52
               Top             =   765
               Width           =   555
            End
            Begin VB.Label Label13 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   51
               Top             =   360
               Width           =   405
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            Caption         =   " Relacionada - SubSidiaria "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   180
            TabIndex        =   35
            Top             =   1755
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableVRelSD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":9565
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   720
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableVRelSS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":98EF
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   315
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVRelSS 
               Height          =   315
               Left            =   1170
               TabIndex        =   38
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":9C79
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVRelSS 
               Height          =   315
               Left            =   2205
               TabIndex        =   39
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":9C95
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVRelSD 
               Height          =   315
               Left            =   1170
               TabIndex        =   40
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":9CB1
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVRelSD 
               Height          =   315
               Left            =   2205
               TabIndex        =   41
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":9CCD
               Vacio           =   -1  'True
            End
            Begin VB.Label Label12 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   43
               Top             =   765
               Width           =   555
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   42
               Top             =   360
               Width           =   405
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            Caption         =   " Terceros "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   180
            TabIndex        =   18
            Top             =   405
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableVTerS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":9CE9
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   315
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableVTerD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":A073
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   720
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVTerS 
               Height          =   315
               Left            =   1170
               TabIndex        =   21
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":A3FD
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVTerS 
               Height          =   315
               Left            =   2205
               TabIndex        =   22
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":A419
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVTerD 
               Height          =   315
               Left            =   1170
               TabIndex        =   23
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":A435
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVTerD 
               Height          =   315
               Left            =   2205
               TabIndex        =   24
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":A451
               Vacio           =   -1  'True
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   26
               Top             =   360
               Width           =   405
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   25
               Top             =   765
               Width           =   555
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            Caption         =   " Relacionada - Matriz "
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   7380
            TabIndex        =   17
            Top             =   405
            Width           =   7035
            Begin VB.CommandButton CmdAyudaCtaContableVRelMS 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":A46D
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   315
               Width           =   345
            End
            Begin VB.CommandButton CmdAyudaCtaContableVRelMD 
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
               Left            =   6480
               Picture         =   "FrmMantDocumentos.frx":A7F7
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   720
               Width           =   345
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVRelMS 
               Height          =   315
               Left            =   1170
               TabIndex        =   29
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":AB81
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVRelMS 
               Height          =   315
               Left            =   2205
               TabIndex        =   30
               Top             =   315
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":AB9D
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtCodCtaContableVRelMD 
               Height          =   315
               Left            =   1170
               TabIndex        =   31
               Top             =   720
               Width           =   1005
               _ExtentX        =   1773
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
               Container       =   "FrmMantDocumentos.frx":ABB9
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox TxtGlsCtaContableVRelMD 
               Height          =   315
               Left            =   2205
               TabIndex        =   32
               Top             =   720
               Width           =   4245
               _ExtentX        =   7488
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
               Container       =   "FrmMantDocumentos.frx":ABD5
               Vacio           =   -1  'True
            End
            Begin VB.Label Label10 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Soles"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   34
               Top             =   360
               Width           =   405
            End
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Dólares"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   270
               TabIndex        =   33
               Top             =   765
               Width           =   555
            End
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   12465
         TabIndex        =   15
         Top             =   1215
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Abreviatura"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   1140
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   210
         Left            =   13185
         TabIndex        =   10
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   210
         Left            =   270
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Factor"
      Height          =   210
      Left            =   800
      TabIndex        =   11
      Top             =   1155
      Width           =   465
   End
End
Attribute VB_Name = "FrmMantDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim indBoton                                As Integer
Dim CIdDocumentoAux                         As String

Private Sub Ayuda_PlanContable(StrMsgError As String, txtCod As Object)
On Error GoTo Err
Dim Cod                                     As String
    
    mostrarAyudaTextoPlanCuentas strcnConta, "PLANCUENTAS", Cod, "", "", "2011"
    If StrMsgError <> "" Then GoTo Err
    
    If Len(Trim("" & Cod)) > 0 Then txtCod.Text = Cod
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub CmdAyudaCtaContableCPerD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCPerD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCPerS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCPerS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCRelAD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelAD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCRelAS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelAS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCRelMD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelMD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCRelMS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelMS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCRelSD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelSD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCRelSS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelSS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCSocD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCSocD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCSocS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCSocS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCTerD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCTerD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableCTerS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableCTerS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVPerD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVPerD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVPerS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVPerS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVRelAD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelAD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVRelAS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelAS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVRelMD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelMD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVRelMS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelMS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVRelSD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelSD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVRelSS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelSS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVSocD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVSocD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVSocS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVSocS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVTerD_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVTerD
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCtaContableVTerS_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    Ayuda_PlanContable StrMsgError, TxtCodCtaContableVTerS
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError                                 As String
    
    Me.top = 0
    Me.left = 0
    
    ConfGrid gLista, False, False, False, False
 
    ListaDocumentos StrMsgError
    If StrMsgError <> "" Then GoTo Err

    fraListado.Visible = True
    fraGeneral.Visible = False
    
    CmbEstado.AddItem "Activado" & Space(150) & "A"
    CmbEstado.AddItem "Desactivado" & Space(150) & "D"
    CmbEstado.ListIndex = 0
    
    habilitaBotones StrMsgError, 7
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ListaDocumentos(StrMsgError As String)
On Error GoTo Err
Dim strCond                                 As String
Dim CSqlC                                   As String
Dim rsdatos                                 As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = "Where IdDocumento Like'%" & strCond & "%' Or GlsDocumento Like'%" & strCond & "%' Or AbreDocumento Like'%" & strCond & "%' "
    End If
    
    CSqlC = "Select IdDocumento,GlsDocumento,AbreDocumento " & _
            "From Documentos " & _
            strCond & _
            "Order By IdDocumento"
            
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open CSqlC, Cn, adOpenStatic, adLockOptimistic
        
    Set gLista.DataSource = rsdatos

'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = CSqlC
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "IdDocumento"
'    End With
    
    Me.Refresh
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError                                     As String

    MostrarDocumento gLista.Columns.ColumnByName("IdDocumento").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    
    habilitaBotones StrMsgError, 2
    If StrMsgError <> "" Then GoTo Err
    
    SSTab1.Tab = 0
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub MostrarDocumento(PIdDocumento As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim RsC                                             As New ADODB.Recordset
Dim CSqlC                                           As String

    CSqlC = "Select A.IdDocumento,A.GlsDocumento,A.AbreDocumento,A.EstDocumentos,B.IdCtaVtaSoles,B.IdCtaVtaDolares,B.IdCtaVtaTerSoles,B.IdCtaVtaTerDolares," & _
            "B.IdCtaCompSoles,B.IdCtaCompDolares,B.IdCtaCompTerSoles,B.IdCtaCompTerDolares,B.IdCtaCompTerMatrizS,B.IdCtaCompTerMatrizD,B.IdCtaCompTerSubSiS," & _
            "B.IdCtaCompTerSubSiD,B.IdCtaVtaTerMatrizS,B.IdCtaVtaTerMatrizD,B.IdCtaVtaTerSubSiS,B.IdCtaVtaTerSubSiD,B.IdCtaVtaPerS,B.IdCtaVtaPerD," & _
            "B.IdCtaCompPerS,B.IdCtaCompPerD,B.IdCtaVtaSocS,B.IdCtaVtaSocD,B.IdCtaCompSocS,B.IdCtaCompSocD " & _
            "From Documentos A " & _
            "Left Join DocumentosCuentas B " & _
                "On '" & glsEmpresa & "' = B.IdEmpresa And A.IdDocumento = B.IdDocumento " & _
            "Where A.IdDocumento = '" & PIdDocumento & "'"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not RsC.EOF Then
        
        TxtCodDocumento.Text = Trim("" & RsC.Fields("IdDocumento"))
        TxtGlsDocumento.Text = Trim("" & RsC.Fields("GlsDocumento"))
        TxtAbreviatura.Text = Trim("" & RsC.Fields("AbreDocumento"))
        CmbEstado.ListIndex = IIf(Trim("" & RsC.Fields("EstDocumentos")) = "ACT", 0, 1)
        TxtCodCtaContableVTerS.Text = Trim("" & RsC.Fields("IdCtaVtaSoles"))
        TxtCodCtaContableVTerD.Text = Trim("" & RsC.Fields("IdCtaVtaDolares"))
        TxtCodCtaContableVRelMS.Text = Trim("" & RsC.Fields("IdCtaVtaTerMatrizS"))
        TxtCodCtaContableVRelMD.Text = Trim("" & RsC.Fields("IdCtaVtaTerMatrizD"))
        TxtCodCtaContableVRelSS.Text = Trim("" & RsC.Fields("IdCtaVtaTerSubSiS"))
        TxtCodCtaContableVRelSD.Text = Trim("" & RsC.Fields("IdCtaVtaTerSubSiD"))
        TxtCodCtaContableVRelAS.Text = Trim("" & RsC.Fields("IdCtaVtaTerSoles"))
        TxtCodCtaContableVRelAD.Text = Trim("" & RsC.Fields("IdCtaVtaTerDolares"))
        TxtCodCtaContableCTerS.Text = Trim("" & RsC.Fields("IdCtaCompSoles"))
        TxtCodCtaContableCTerD.Text = Trim("" & RsC.Fields("IdCtaCompDolares"))
        TxtCodCtaContableCRelMS.Text = Trim("" & RsC.Fields("IdCtaCompTerMatrizS"))
        TxtCodCtaContableCRelMD.Text = Trim("" & RsC.Fields("IdCtaCompTerMatrizD"))
        TxtCodCtaContableCRelSS.Text = Trim("" & RsC.Fields("IdCtaCompTerSubSiS"))
        TxtCodCtaContableCRelSD.Text = Trim("" & RsC.Fields("IdCtaCompTerSubSiD"))
        TxtCodCtaContableCRelAS.Text = Trim("" & RsC.Fields("IdCtaCompTerSoles"))
        TxtCodCtaContableCRelAD.Text = Trim("" & RsC.Fields("IdCtaCompTerDolares"))
        TxtCodCtaContableCPerS.Text = Trim("" & RsC.Fields("IdCtaCompPerS"))
        TxtCodCtaContableCPerD.Text = Trim("" & RsC.Fields("IdCtaCompPerD"))
        TxtCodCtaContableVPerS.Text = Trim("" & RsC.Fields("IdCtaVtaPerS"))
        TxtCodCtaContableVPerD.Text = Trim("" & RsC.Fields("IdCtaVtaPerD"))
        TxtCodCtaContableCSocS.Text = Trim("" & RsC.Fields("IdCtaCompSocS"))
        TxtCodCtaContableCSocD.Text = Trim("" & RsC.Fields("IdCtaCompSocD"))
        TxtCodCtaContableVSocS.Text = Trim("" & RsC.Fields("IdCtaVtaSocS"))
        TxtCodCtaContableVSocD.Text = Trim("" & RsC.Fields("IdCtaVtaSocD"))
        
        CIdDocumentoAux = PIdDocumento
        
    End If
    
    RsC.Close: Set RsC = Nothing
    
Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gLista_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error GoTo Err
Dim StrMsgError                                     As String

    If KeyCode = 116 Then
        
        ListaDocumentos StrMsgError
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub nuevo(StrMsgError As String)
On Error GoTo Err

    indBoton = 0
    limpiaForm Me
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = True
    
    TxtCodDocumento.Vacio = True
    CIdDocumentoAux = ""
    
    SSTab1.Tab = 0
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError                                 As String

    Select Case Button.Index
        Case 1 'Nuevo
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        Case 2 'Grabar
            Grabar StrMsgError, indBoton
            If StrMsgError <> "" Then GoTo Err
            
        Case 3 'Modificar
            indBoton = 1
            fraGeneral.Enabled = True
            
            TxtCodDocumento.Vacio = False
            
        Case 4, 7 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
            
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        Case 6 'Imprimir
            gLista.m.ExportToXLS App.Path & "\Temporales\Mantenimiento_Documentos.xls"
            ShellEx App.Path & "\Temporales\Mantenimiento_Documentos.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
            
        Case 7 'Lista
            fraListado.Visible = True
            fraGeneral.Visible = False
            
        Case 8 'Salir
            Unload Me
            
    End Select
       
    habilitaBotones StrMsgError, Button.Index
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub habilitaBotones(StrMsgError As String, indexBoton As Integer)
On Error GoTo Err
    
    Select Case indexBoton
        Case 1, 5 'Nuevo, Eliminar
            Toolbar1.Buttons(1).Visible = False 'Nuevo
            Toolbar1.Buttons(2).Visible = True 'Grabar
            Toolbar1.Buttons(3).Visible = False 'Modificar
            Toolbar1.Buttons(4).Visible = False 'Cancelar
            Toolbar1.Buttons(5).Visible = False 'Eliminar
            Toolbar1.Buttons(6).Visible = False 'Imprimir
            Toolbar1.Buttons(7).Visible = True 'Lista
        Case 2 'Grabar, Cancelar
            Toolbar1.Buttons(1).Visible = True 'Nuevo
            Toolbar1.Buttons(2).Visible = False 'Grabar
            Toolbar1.Buttons(3).Visible = True 'Modificar
            Toolbar1.Buttons(4).Visible = False 'Cancelar
            Toolbar1.Buttons(5).Visible = True 'Eliminar
            Toolbar1.Buttons(6).Visible = False 'Imprimir
            Toolbar1.Buttons(7).Visible = True 'Lista
        Case 3 'Modificar
            Toolbar1.Buttons(1).Visible = False 'Nuevo
            Toolbar1.Buttons(2).Visible = True 'Grabar
            Toolbar1.Buttons(3).Visible = False 'Modificar
            Toolbar1.Buttons(4).Visible = True 'Cancelar
            Toolbar1.Buttons(5).Visible = False 'Eliminar
            Toolbar1.Buttons(6).Visible = False 'Imprimir
            Toolbar1.Buttons(7).Visible = False 'Lista
        Case 7, 4 'Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
    End Select

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Grabar(ByRef StrMsgError As String, IndGraba As Integer)
On Error GoTo Err
Dim strCodigo                               As String
Dim strMsg                                  As String
Dim CSqlC                                   As String
Dim cEstado                                 As String
Dim indTrans                                As Boolean
    
    indTrans = False
    
    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    cEstado = IIf(CmbEstado.ListIndex = 0, "ACT", "INA")
    
    If indBoton = 0 Then
        
        If Len(Trim("" & TxtCodDocumento.Text)) = 0 Then
            TxtCodDocumento.Text = generaCorrelativo("Documentos", "IdDocumento", 2, , False)
        Else
            
            If Len(Trim("" & traerCampo("Documentos", "IdDocumento", "IdDocumento", TxtCodDocumento.Text, False))) > 0 Then
            
                StrMsgError = "El Código ya existe. Verifique.": GoTo Err
            
            End If
            
        End If
        
        If Len(Trim("" & traerCampo("Documentos", "AbreDocumento", "AbreDocumento", TxtAbreviatura.Text, False))) > 0 Then
            
            StrMsgError = "La Abreviatura ya existe. Verifique.": GoTo Err
        
        End If
            
        CSqlC = "Insert Into Documentos(IdDocumento,GlsDocumento,AbreDocumento,EstDocumentos)Values(" & _
                "'" & TxtCodDocumento.Text & "','" & TxtGlsDocumento.Text & "','" & TxtAbreviatura.Text & "','" & cEstado & "')"
        
        Cn.BeginTrans
        indTrans = True
        Cn.Execute CSqlC
        
        CSqlC = "Insert Into DocumentosCuentas(IdEmpresa,IdDocumento,IdCtaVtaSoles,IdCtaVtaDolares,IdCtaVtaTerSoles,IdCtaVtaTerDolares,IdCtaCompSoles," & _
                "IdCtaCompDolares,IdCtaCompTerSoles,IdCtaCompTerDolares,IdCtaCompTerMatrizS,IdCtaCompTerMatrizD,IdCtaCompTerSubSiS,IdCtaCompTerSubSiD," & _
                "IdCtaVtaTerMatrizS,IdCtaVtaTerMatrizD,IdCtaVtaTerSubSiS,IdCtaVtaTerSubSiD,IdCtaVtaPerS,IdCtaVtaPerD,IdCtaCompPerS,IdCtaCompPerD," & _
                "IdCtaVtaSocS,IdCtaVtaSocD,IdCtaCompSocS,IdCtaCompSocD)Values(" & _
                "'" & glsEmpresa & "','" & TxtCodDocumento.Text & "'," & _
                "'" & TxtCodCtaContableVTerS.Text & "','" & TxtCodCtaContableVTerD.Text & "','" & TxtCodCtaContableVRelAS.Text & "'," & _
                "'" & TxtCodCtaContableVRelAD.Text & "','" & TxtCodCtaContableCTerS.Text & "','" & TxtCodCtaContableCTerD.Text & "'," & _
                "'" & TxtCodCtaContableCRelAS.Text & "','" & TxtCodCtaContableCRelAD.Text & "','" & TxtCodCtaContableCRelMS.Text & "'," & _
                "'" & TxtCodCtaContableCRelMD.Text & "','" & TxtCodCtaContableCRelSS.Text & "','" & TxtCodCtaContableCRelSD.Text & "'," & _
                "'" & TxtCodCtaContableVRelMS.Text & "','" & TxtCodCtaContableVRelMD.Text & "','" & TxtCodCtaContableVRelSS.Text & "'," & _
                "'" & TxtCodCtaContableVRelSD.Text & "','" & TxtCodCtaContableVPerS.Text & "','" & TxtCodCtaContableVPerD.Text & "'," & _
                "'" & TxtCodCtaContableCPerS.Text & "','" & TxtCodCtaContableCPerD.Text & "','" & TxtCodCtaContableVSocS.Text & "'," & _
                "'" & TxtCodCtaContableVSocD.Text & "','" & TxtCodCtaContableCSocS.Text & "','" & TxtCodCtaContableCSocD.Text & "')"
        
        Cn.Execute CSqlC
        
        strMsg = "Grabo"
        
    Else
        
        If Len(Trim("" & traerCampo("Documentos", "IdDocumento", "IdDocumento", TxtCodDocumento.Text, False, "IdDocumento <> '" & CIdDocumentoAux & "'"))) > 0 Then
            
            StrMsgError = "El Nuevo Código ya existe. Verifique.": GoTo Err
        
        End If
        
        If Len(Trim("" & traerCampo("Documentos", "AbreDocumento", "AbreDocumento", TxtAbreviatura.Text, False, "IdDocumento <> '" & CIdDocumentoAux & "'"))) > 0 Then
            
            StrMsgError = "La Nueva Abreviatura ya existe. Verifique.": GoTo Err
        
        End If
            
        CSqlC = "Update Documentos " & _
                "Set IdDocumento = '" & TxtCodDocumento.Text & "',GlsDocumento = '" & TxtGlsDocumento.Text & "',AbreDocumento = '" & TxtAbreviatura.Text & "'," & _
                "EstDocumentos = '" & cEstado & "' " & _
                "Where IdDocumento = '" & CIdDocumentoAux & "'"
        
        Cn.BeginTrans
        indTrans = True
        Cn.Execute CSqlC
        
        CSqlC = "Update DocumentosCuentas " & _
                "Set IdCtaVtaSoles = '" & TxtCodCtaContableVTerS.Text & "'," & _
                "IdCtaVtaDolares = '" & TxtCodCtaContableVTerD.Text & "',IdCtaVtaTerSoles = '" & TxtCodCtaContableVRelAS.Text & "'," & _
                "IdCtaVtaTerDolares = '" & TxtCodCtaContableVRelAD.Text & "',IdCtaCompSoles = '" & TxtCodCtaContableCTerS.Text & "'," & _
                "IdCtaCompDolares = '" & TxtCodCtaContableCTerD.Text & "',IdCtaCompTerSoles = '" & TxtCodCtaContableCRelAS.Text & "'," & _
                "IdCtaCompTerDolares = '" & TxtCodCtaContableCRelAD.Text & "',IdCtaCompTerMatrizS = '" & TxtCodCtaContableCRelMS.Text & "'," & _
                "IdCtaCompTerMatrizD = '" & TxtCodCtaContableCRelMD.Text & "',IdCtaCompTerSubSiS = '" & TxtCodCtaContableCRelSS.Text & "'," & _
                "IdCtaCompTerSubSiD = '" & TxtCodCtaContableCRelSD.Text & "',IdCtaVtaTerMatrizS = '" & TxtCodCtaContableVRelMS.Text & "'," & _
                "IdCtaVtaTerMatrizD = '" & TxtCodCtaContableVRelMD.Text & "',IdCtaVtaTerSubSiS = '" & TxtCodCtaContableVRelSS.Text & "'," & _
                "IdCtaVtaTerSubSiD = '" & TxtCodCtaContableVRelSD.Text & "',IdCtaVtaPerS = '" & TxtCodCtaContableVPerS.Text & "'," & _
                "IdCtaVtaPerD = '" & TxtCodCtaContableVPerD.Text & "',IdCtaCompPerS = '" & TxtCodCtaContableCPerS.Text & "'," & _
                "IdCtaCompPerD = '" & TxtCodCtaContableCPerD.Text & "',IdCtaVtaSocS = '" & TxtCodCtaContableVSocS.Text & "',IdCtaVtaSocD = '" & TxtCodCtaContableVSocD.Text & "',IdCtaCompSocS = '" & TxtCodCtaContableCSocS.Text & "',IdCtaCompSocD = '" & TxtCodCtaContableCSocD.Text & "' " & _
                "Where IdEmpresa = '" & glsEmpresa & "' And IdDocumento = '" & CIdDocumentoAux & "'"
        
        Cn.Execute CSqlC
        
        strMsg = "Modifico"
        
    End If
    
    Cn.CommitTrans
    indTrans = False
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    indBoton = 1
    fraGeneral.Enabled = False
    CIdDocumentoAux = TxtCodDocumento.Text
    
    ListaDocumentos StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If indTrans Then Cn.RollbackTrans: indTrans = False
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub eliminar(StrMsgError As String)
On Error GoTo Err
Dim indTrans                                As Boolean
Dim strCodigo                               As String
Dim RsC                                     As New ADODB.Recordset
Dim CSqlC                                   As String

    If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    strCodigo = Trim(TxtCodDocumento.Text)
    
    CSqlC = "Select IdDocumento From DocVentas Where IdDocumento Not In('94','OS','87') And IdDocumento = '" & strCodigo & "' Limit 1"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not RsC.EOF Then StrMsgError = "El Documento está siendo usado en Ventas": GoTo Err
    RsC.Close: Set RsC = Nothing
    
    CSqlC = "Select IdDocumento From DocVentas Where IdDocumento In('94','OS','87') And IdDocumento = '" & strCodigo & "' Limit 1"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not RsC.EOF Then StrMsgError = "El Documento está siendo usado en Compras": GoTo Err
    RsC.Close: Set RsC = Nothing
    
    CSqlC = "Select IdDocumento From DocVentasPres Where IdDocumento = '" & strCodigo & "' Limit 1"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not RsC.EOF Then StrMsgError = "El Documento está siendo usado en Producción": GoTo Err
    RsC.Close: Set RsC = Nothing
    
    CSqlC = "Select IdDocumento From DocVentasPedidos Where IdDocumento = '" & strCodigo & "' Limit 1"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not RsC.EOF Then StrMsgError = "El Documento está siendo usado en Producción, Pedido de Materiales": GoTo Err
    RsC.Close: Set RsC = Nothing
    
    CSqlC = "Select TipoDcto From RegisDoc Where TipoDcto = '" & strCodigo & "' Limit 1"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not RsC.EOF Then StrMsgError = "El Documento está siendo usado en el Registro de Compras": GoTo Err
    RsC.Close: Set RsC = Nothing
    
    Cn.BeginTrans
    indTrans = True
 
    CSqlC = "Delete From Documentos Where IdDocumento = '" & strCodigo & "'"
    Cn.Execute CSqlC
     
    CSqlC = "Delete From DocumentosCuentas Where IdEmpresa = '" & glsEmpresa & "' And IdDocumento = '" & strCodigo & "'"
    Cn.Execute CSqlC
    
    Cn.CommitTrans
    indTrans = False
    
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    
    ListaDocumentos StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If indTrans Then Cn.RollbackTrans: indTrans = False
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    ListaDocumentos StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCPerD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCPerD.Text <> "" Then
        TxtGlsCtaContableCPerD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCPerD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCPerD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCPerS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCPerS.Text <> "" Then
        TxtGlsCtaContableCPerS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCPerS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCPerS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelAD_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelAD
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelAS_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelAS
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelMD_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelMD
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelMS_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelMS
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelSD_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelSD
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelSS_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableCRelSS
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCSocD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCSocD.Text <> "" Then
        TxtGlsCtaContableCSocD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCSocD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCSocD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCSocS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCSocS.Text <> "" Then
        TxtGlsCtaContableCSocS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCSocS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCSocS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCTerD_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableCTerD
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCTerS_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableCTerS
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVPerS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVPerS.Text <> "" Then
        TxtGlsCtaContableVPerS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVPerS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVPerS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVPerD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVPerD.Text <> "" Then
        TxtGlsCtaContableVPerD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVPerD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVPerD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelAD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVRelAD.Text <> "" Then
        TxtGlsCtaContableVRelAD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVRelAD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVRelAD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelAD_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelAD
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelAS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVRelAS.Text <> "" Then
        TxtGlsCtaContableVRelAS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVRelAS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVRelAS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelAS_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelAS
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelMD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVRelMD.Text <> "" Then
        TxtGlsCtaContableVRelMD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVRelMD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVRelMD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelMD_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelMD
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelMS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVRelMS.Text <> "" Then
        TxtGlsCtaContableVRelMS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVRelMS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVRelMS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelMS_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelMS
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelSD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVRelSD.Text <> "" Then
        TxtGlsCtaContableVRelSD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVRelSD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVRelSD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelSD_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelSD
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelSS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVRelSS.Text <> "" Then
        TxtGlsCtaContableVRelSS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVRelSS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVRelSS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVRelSS_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableVRelSS
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVSocD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVSocD.Text <> "" Then
        TxtGlsCtaContableVSocD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVSocD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVSocD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVSocS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVSocS.Text <> "" Then
        TxtGlsCtaContableVSocS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVSocS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVSocS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVTerD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVTerD.Text <> "" Then
        TxtGlsCtaContableVTerD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVTerD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVTerD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVTerD_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableVTerD
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVTerS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableVTerS.Text <> "" Then
        TxtGlsCtaContableVTerS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableVTerS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableVTerS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                                     As String

    If KeyCode = 116 Then
        
        ListaDocumentos StrMsgError
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelAD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCRelAD.Text <> "" Then
        TxtGlsCtaContableCRelAD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCRelAD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCRelAD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelAS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCRelAS.Text <> "" Then
        TxtGlsCtaContableCRelAS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCRelAS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCRelAS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelMD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCRelMD.Text <> "" Then
        TxtGlsCtaContableCRelMD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCRelMD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCRelMD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelMS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCRelMS.Text <> "" Then
        TxtGlsCtaContableCRelMS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCRelMS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCRelMS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelSD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCRelSD.Text <> "" Then
        TxtGlsCtaContableCRelSD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCRelSD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCRelSD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCRelSS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCRelSS.Text <> "" Then
        TxtGlsCtaContableCRelSS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCRelSS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCRelSS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCTerD_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCTerD.Text <> "" Then
        TxtGlsCtaContableCTerD.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCTerD.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCTerD.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableCTerS_Change()
On Error GoTo Err
Dim StrMsgError                             As String

    If TxtCodCtaContableCTerS.Text <> "" Then
        TxtGlsCtaContableCTerS.Text = Trim("" & traerCampoConta("PlanCuentas", "GlsNombreCuenta", "IdCtaContable", TxtCodCtaContableCTerS.Text, True, "IdAnno = '2011'"))
    Else
        TxtGlsCtaContableCTerS.Text = ""
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodCtaContableVTerS_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                             As String
    
    If KeyCode = 114 Then
    
        Ayuda_PlanContable StrMsgError, TxtCodCtaContableVTerS
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
