VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantListaPrecios 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Listas de Precios"
   ClientHeight    =   9390
   ClientLeft      =   1740
   ClientTop       =   1245
   ClientWidth     =   13065
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
   ScaleHeight     =   9390
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   7575
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
            Picture         =   "frmMantListaPrecios.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantListaPrecios.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
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
      Height          =   8550
      Left            =   45
      TabIndex        =   0
      Top             =   705
      Width           =   12930
      Begin VB.Frame fraDetalle 
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
         Height          =   6735
         Left            =   165
         TabIndex        =   1
         Top             =   1635
         Width           =   12600
         Begin VB.CommandButton cmbPrecio 
            Caption         =   "Nuevo Precio"
            Height          =   465
            Left            =   11160
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   270
            Width           =   1185
         End
         Begin VB.Frame fraContenido 
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
            Height          =   5850
            Left            =   135
            TabIndex        =   39
            Top             =   630
            Width           =   12330
            Begin DXDBGRIDLibCtl.dxDBGrid gListaPrecio 
               Height          =   5400
               Left            =   45
               OleObjectBlob   =   "frmMantListaPrecios.frx":3518
               TabIndex        =   40
               Top             =   450
               Width           =   12225
            End
            Begin CATControls.CATTextBox txtBus_Producto 
               Height          =   315
               Left            =   1305
               TabIndex        =   11
               Top             =   60
               Width           =   7080
               _ExtentX        =   12488
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
               Container       =   "frmMantListaPrecios.frx":C205
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Búsqueda"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   120
               TabIndex        =   42
               Top             =   120
               Width           =   735
            End
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
            Height          =   405
            Left            =   120
            TabIndex        =   23
            Top             =   300
            Width           =   8985
            Begin VB.CommandButton cmbAyudaNivel 
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
               Left            =   8400
               Picture         =   "frmMantListaPrecios.frx":C221
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   1440
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaNivel 
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
               Left            =   8400
               Picture         =   "frmMantListaPrecios.frx":C5AB
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   1080
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaNivel 
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
               Left            =   8400
               Picture         =   "frmMantListaPrecios.frx":C935
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   720
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaNivel 
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
               Left            =   8400
               Picture         =   "frmMantListaPrecios.frx":CCBF
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   360
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaNivel 
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
               Left            =   8400
               Picture         =   "frmMantListaPrecios.frx":D049
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   0
               Width           =   390
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   0
               Left            =   1320
               TabIndex        =   6
               Tag             =   "TidNivelPred"
               Top             =   20
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
               Container       =   "frmMantListaPrecios.frx":D3D3
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   0
               Left            =   2280
               TabIndex        =   29
               Top             =   20
               Width           =   6090
               _ExtentX        =   10742
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
               Container       =   "frmMantListaPrecios.frx":D3EF
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   1
               Left            =   1305
               TabIndex        =   7
               Tag             =   "TidNivelPred"
               Top             =   390
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
               Container       =   "frmMantListaPrecios.frx":D40B
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   1
               Left            =   2280
               TabIndex        =   30
               Top             =   390
               Width           =   6090
               _ExtentX        =   10742
               _ExtentY        =   556
               BackColor       =   12648447
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
               Container       =   "frmMantListaPrecios.frx":D427
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   2
               Left            =   1305
               TabIndex        =   8
               Tag             =   "TidNivelPred"
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
               Container       =   "frmMantListaPrecios.frx":D443
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   2
               Left            =   2280
               TabIndex        =   31
               Top             =   750
               Width           =   6090
               _ExtentX        =   10742
               _ExtentY        =   556
               BackColor       =   12648447
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
               Container       =   "frmMantListaPrecios.frx":D45F
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   3
               Left            =   1305
               TabIndex        =   9
               Tag             =   "TidNivelPred"
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
               Container       =   "frmMantListaPrecios.frx":D47B
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   3
               Left            =   2280
               TabIndex        =   32
               Top             =   1110
               Width           =   6090
               _ExtentX        =   10742
               _ExtentY        =   556
               BackColor       =   12648447
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
               Container       =   "frmMantListaPrecios.frx":D497
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   4
               Left            =   1305
               TabIndex        =   10
               Tag             =   "TidNivelPred"
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
               Container       =   "frmMantListaPrecios.frx":D4B3
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   4
               Left            =   2280
               TabIndex        =   33
               Top             =   1470
               Width           =   6090
               _ExtentX        =   10742
               _ExtentY        =   556
               BackColor       =   12648447
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
               Container       =   "frmMantListaPrecios.frx":D4CF
               Vacio           =   -1  'True
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   135
               TabIndex        =   38
               Top             =   1485
               Width           =   345
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   135
               TabIndex        =   37
               Top             =   1125
               Width           =   345
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   135
               TabIndex        =   36
               Top             =   765
               Width           =   345
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   135
               TabIndex        =   35
               Top             =   405
               Width           =   345
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   135
               TabIndex        =   34
               Top             =   45
               Width           =   345
            End
         End
      End
      Begin VB.Frame fraCabecera 
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
         Height          =   1395
         Left            =   165
         TabIndex        =   13
         Top             =   135
         Width           =   12600
         Begin VB.CommandButton cmbAyudaMoneda 
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
            Left            =   6225
            Picture         =   "frmMantListaPrecios.frx":D4EB
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   935
            Width           =   390
         End
         Begin VB.CheckBox chkEstado 
            Appearance      =   0  'Flat
            Caption         =   "Activo"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   11550
            TabIndex        =   5
            Tag             =   "NestLista"
            Top             =   930
            Value           =   1  'Checked
            Width           =   840
         End
         Begin CATControls.CATTextBox txtCod_Lista 
            Height          =   285
            Left            =   11475
            TabIndex        =   14
            Tag             =   "TidLista"
            Top             =   195
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
            Container       =   "frmMantListaPrecios.frx":D875
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Lista 
            Height          =   315
            Left            =   1290
            TabIndex        =   2
            Tag             =   "TglsLista"
            Top             =   555
            Width           =   11100
            _ExtentX        =   19579
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
            Container       =   "frmMantListaPrecios.frx":D891
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp_Emision 
            Height          =   315
            Left            =   8415
            TabIndex        =   4
            Tag             =   "FfecVcto"
            Top             =   930
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
            Format          =   132120577
            CurrentDate     =   38955
         End
         Begin CATControls.CATTextBox txtCod_Moneda 
            Height          =   315
            Left            =   1290
            TabIndex        =   3
            Tag             =   "TidMoneda"
            Top             =   930
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
            Container       =   "frmMantListaPrecios.frx":D8AD
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Moneda 
            Height          =   315
            Left            =   2220
            TabIndex        =   44
            Top             =   930
            Width           =   3990
            _ExtentX        =   7038
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
            Container       =   "frmMantListaPrecios.frx":D8C9
            Vacio           =   -1  'True
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   45
            Top             =   975
            Width           =   570
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Código"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   10770
            TabIndex        =   17
            Top             =   210
            Width           =   495
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vcto."
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7215
            TabIndex        =   15
            Top             =   975
            Width           =   885
         End
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
      ForeColor       =   &H00C00000&
      Height          =   8550
      Left            =   60
      TabIndex        =   18
      Top             =   720
      Width           =   12900
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
         TabIndex        =   19
         Top             =   180
         Width           =   12630
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   990
            TabIndex        =   20
            Top             =   270
            Width           =   11505
            _ExtentX        =   20294
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
            Container       =   "frmMantListaPrecios.frx":D8E5
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
            Left            =   120
            TabIndex        =   21
            Top             =   300
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   7335
         Left            =   120
         OleObjectBlob   =   "frmMantListaPrecios.frx":D901
         TabIndex        =   22
         Top             =   1065
         Width           =   12660
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   1164
      ButtonWidth     =   3043
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "           Nuevo        "
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
            Caption         =   "Copiar Lista"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actualiza Max. Dcto."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMantListaPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private indInserta  As Boolean
Private NumNiveles As Integer
Private indMovNivel As Boolean
Dim cParametro  As String

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda

End Sub

Private Sub cmbAyudaNivel_Click(Index As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
    
    peso = Index + 1
    strCodTipoNivel = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
    strCondPred = ""
    If peso > 1 Then
        strCondPred = " AND idNivelPred = '" & txtCod_Nivel(Index - 1).Text & "'"
    End If
    mostrarAyuda "NIVEL", txtCod_Nivel(Index), txtGls_Nivel(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred
    
End Sub

Private Sub cmbPrecio_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim strCodNivel As String

    strCodNivel = txtCod_Nivel(NumNiveles - 1).Text
    frmMantPrecios.mostrarFormNuevo txtCod_Lista.Text, strCodNivel, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    ListaDetallePrecios StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

    Me.left = 0
    Me.top = 0
    
    indMovNivel = False
    indInserta = False
    
    ConfGrid gLista, False, False, False, False
    ConfGrid gListaPrecio, False, False, False, False
    
    gListaPrecio.Columns.ColumnByFieldName("VVUnit").DecimalPlaces = glsDecimalesPrecios
    gListaPrecio.Columns.ColumnByFieldName("IGVUnit").DecimalPlaces = glsDecimalesPrecios
    gListaPrecio.Columns.ColumnByFieldName("PVUnit").DecimalPlaces = glsDecimalesPrecios
    
    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    cParametro = traerCampo("Parametros", "Valparametro", "GlsParametro", "VISUALIZA_LISTA_PRECIO_ALMACEN", True)
    listaListaPrecio StrMsgError
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
    
    validaHomonimia "listaprecios", "GlsLista", "idLista", txtGls_Lista.Text, txtCod_Lista.Text, True, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_Lista.Text = "" Then 'graba
        txtCod_Lista.Text = GeneraCorrelativoAnoMes("listaprecios", "idLista")
        EjecutaSQLForm Me, 0, True, "listaprecios", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabo"
    
    Else 'modifica
        EjecutaSQLForm Me, 1, True, "listaprecios", StrMsgError, "idLista"
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modifico"
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    fraCabecera.Enabled = False
    listaListaPrecio StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()
    
    limpiaForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    gListaPrecio.Dataset.Active = False
    Set gListaPrecio.DataSource = Nothing
    
    gLista.Dataset.Active = False
    Set gLista.DataSource = Nothing

End Sub

Private Sub gListaPrecio_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String
Dim strCodProd As String
Dim strCodUM As String
Dim dblFactor As Double
Dim dblVV As Double
Dim dblIGV As Double
Dim dblPV As Double
Dim dblCostoUnit As Double
Dim dblFactorUnit As Double
Dim dblFactor2Unit As Double
Dim lngFila As Long

    lngFila = gListaPrecio.Dataset.RecNo
    lngFila = gListaPrecio.Dataset.RecNo
    lngFila = gListaPrecio.Dataset.RecNo
    
    strCodProd = gListaPrecio.Columns.ColumnByFieldName("idProducto").Value
    If strCodProd = "" Then Exit Sub
    
    strCodUM = gListaPrecio.Columns.ColumnByFieldName("idUM").Value
    dblFactor = Val("" & gListaPrecio.Columns.ColumnByFieldName("factor").Value)
    
    dblVV = gListaPrecio.Columns.ColumnByFieldName("VVUnit").Value
    dblIGV = gListaPrecio.Columns.ColumnByFieldName("IGVUnit").Value
    dblPV = gListaPrecio.Columns.ColumnByFieldName("PVUnit").Value
    dblCostoUnit = gListaPrecio.Columns.ColumnByFieldName("CostoUnit").Value
    dblFactorUnit = gListaPrecio.Columns.ColumnByFieldName("FactorUnit").Value
    dblFactor2Unit = gListaPrecio.Columns.ColumnByFieldName("Factor2Unit").Value
    
    frmMantPrecios.MostrarForm txtCod_Lista.Text, strCodProd, strCodUM, dblFactor, dblVV, dblIGV, dblPV, dblCostoUnit, dblFactorUnit, dblFactor2Unit, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    ListaDetallePrecios StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gListaPrecio.Dataset.RecNo = lngFila
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gListaPrecio_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error GoTo Err
Dim strCodProd As String
Dim StrMsgError As String

    If KeyCode = 46 Then
        strCodProd = Trim("" & gListaPrecio.Columns.ColumnByFieldName("idProducto").Value)
        If strCodProd = "" Then Exit Sub
        If MsgBox("¿Seguro de eliminar el registro?", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
        csql = "DELETE FROM preciosventa WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & txtCod_Lista.Text & "' AND idProducto = '" & strCodProd & "'"
        Cn.Execute csql
        
        ListaDetallePrecios StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gListaPrecio_OnReloadGroupList()
    
    gListaPrecio.m.FullExpand

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarLista gLista.Columns.ColumnByName("idLista").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraCabecera.Enabled = False
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
            fraCabecera.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            fraCabecera.Enabled = False
        Case 3 'Modificar
            fraCabecera.Enabled = True
        Case 4, 7 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraCabecera.Enabled = False
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 6 'Imprimir
            gListaPrecio.m.ExportToXLS App.Path & "\Temporales\ListadoPrecios.xls"
            ShellEx App.Path & "\Temporales\ListadoPrecios.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 8 'Copiar Lista
            frmCopiarListaPrecios.Show 1
            listaListaPrecio StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 9 'Actualiza Max Dcto
            FrmActualizaMaxDcto.Show 1
        Case 10 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
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
            
            If indexBoton = 2 Then
                Toolbar1.Buttons(8).Visible = Not indHabilitar 'Copiar Lista
            Else
                Toolbar1.Buttons(8).Visible = indHabilitar 'Copiar Lista
            End If
            Toolbar1.Buttons(9).Visible = False 'Max Dcto
            
        Case 4, 7 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = False 'Max Dcto
            
            If indexBoton = "7" And traerCampo("Parametros", "ValParametro", "GlsParametro", "ACTUALIZA_DESCUENTO_NIVEL", True) = "1" Then
                Toolbar1.Buttons(9).Visible = True
            End If
    End Select
    
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaListaPrecio StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaListaPrecio(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND GlsLista LIKE '%" & strCond & "%'"
    End If
    csql = "SELECT a.idLista ,a.GlsLista,a.FecVcto,a.estLista " & _
           "FROM ListaPrecios a WHERE a.idEmpresa = '" & glsEmpresa & "'"
    If strCond <> "" Then csql = csql & strCond
    csql = csql & " ORDER BY a.idLista"
    
'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idLista"
'    End With
    
If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
Set gLista.DataSource = rsdatos
Me.Refresh
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarLista(strCodLista As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT p.idLista,p.GlsLista,p.fecVcto,p.estLista,p.idMoneda " & _
           "FROM ListaPrecios p " & _
           "WHERE p.idLista = '" & strCodLista & "' AND p.idEmpresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
   
    If cParametro = "N" Then
        ListaDetallePrecios StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    Me.Refresh
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtBus_Producto_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyAscii = 13 Then
        ListaDetallePrecios StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MONEDA", txtCod_Moneda, txtGls_Moneda
        KeyAscii = 0
        'If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtCod_Nivel_Change(Index As Integer)
On Error GoTo Err
Dim StrMsgError As String
Dim i As Integer

    If indMovNivel Then Exit Sub
    txtGls_Nivel(Index).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_Nivel(Index).Text, True)
    indMovNivel = True
    For i = Index + 1 To txtCod_Nivel.Count - 1
        txtCod_Nivel(i).Text = ""
        txtGls_Nivel(i).Text = ""
    Next
    indMovNivel = False
    
    If NumNiveles = Index + 1 Then
        ListaDetallePrecios StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Nivel_KeyPress(Index As Integer, KeyAscii As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
Dim strCodJerarquia As String
    
    If KeyAscii <> 13 Then
        peso = Index + 1
        strCodJerarquia = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
        strCondPred = ""
        If peso > 1 Then
            strCondPred = " AND idNivelPred = '" & txtCod_Nivel(Index - 1).Text & "'"
        End If
        mostrarAyudaKeyascii KeyAscii, "NIVEL", txtCod_Nivel(Index), txtGls_Nivel(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred
    End If

End Sub

Private Sub mostrarNiveles(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsj As New ADODB.Recordset
Dim i As Integer

    For i = 0 To 4
        txtCod_Nivel(i).Tag = ""
    Next
    
    '--- Jalamos tipos nivel
    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    NumNiveles = Val("" & rsj.RecordCount)
    fraNivel.Height = 355 * NumNiveles
    i = 0
    Do While Not rsj.EOF
        lblNivel(i).Caption = "" & rsj.Fields("GlsTipoNivel")
        rsj.MoveNext
        i = i + 1
    Loop
    
    fraContenido.Height = fraContenido.Height - (355 * NumNiveles) - 300
    gListaPrecio.Height = fraContenido.Height - txtBus_Producto.top - txtBus_Producto.Height - 100
    fraContenido.top = fraNivel.top + fraNivel.Height '- 70
    fraDetalle.Height = fraNivel.top + fraNivel.Height + fraContenido.Height + 100
    fraGeneral.Height = fraCabecera.top + fraCabecera.Height + fraDetalle.Height + 200
    fraListado.Height = fraGeneral.Height
    gLista.Height = fraListado.Height - gLista.top - 70
    Me.Height = Toolbar1.Height + fraListado.Height + 500
    
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    
    Exit Sub
    
Err:
    If rsj.State = 1 Then rsj.Close:  Set rsj = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub ListaDetallePrecios(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond     As String
Dim strCodNivel As String
Dim strTabla    As String
Dim strWhere    As String
Dim strCampos   As String
Dim strTablas   As String
Dim strTablaAnt As String
Dim ccadagrupa  As String
Dim i           As Integer
Dim cCadStkAlmacenes As String
Dim cCadColAlmacenes As String
Dim rsdatos                     As New ADODB.Recordset

    For i = 1 To NumNiveles
        strTabla = "niveles" & Format(i, "00")
        If i = 1 Then
            strWhere = "p.idNivel = " & strTabla & ".idNivel AND " & strTabla & ".idEmpresa = '" & glsEmpresa & "' "
        Else
            strWhere = strTablaAnt & ".idNivelPred = " & strTabla & ".idNivel AND " & strTabla & ".idEmpresa = '" & glsEmpresa & "' "
        End If
        strCampos = strCampos & strTabla & ".idNivel as idNivel" & Format(i, "00") & "," & strTabla & ".GlsNivel as GlsNivel" & Format(i, "00") & ","
        
        strTablas = strTablas & " INNER JOIN niveles " & strTabla & " ON " & strWhere
        strTablaAnt = strTabla
    Next
    
    If Trim(txtBus_Producto.Text) <> "" Then
        strCond = Trim(txtBus_Producto.Text)
        strCond = " AND p.GlsProducto LIKE '%" & strCond & "%'"
    End If
    
    strCodNivel = Trim(txtCod_Nivel(NumNiveles - 1).Text)
    If strCodNivel <> "" Then
        strCond = strCond + " AND p.idNivel = '" & strCodNivel & "'"
    End If
    
    If cParametro = "S" Then
        cCadStkAlmacenes = "Left Join (Select idAlmacen, p.idProducto, item, idUM, vc.idEmpresa, vc.idSucursal, " & _
                        "CAST(sum(if(vd.tipovale = 'I',Cantidad,Cantidad * -1)) AS NUMERIC(12,2)) as  StkPrincipal, vvUnit " & _
                        "From  ValesDet vd " & _
                        "Inner Join ValesCab vc " & _
                        "On vd.idValesCab = vc.idValesCab   And vd.idEmpresa = vc.idEmpresa   And vd.idSucursal = vc.idSucursal   And vd.tipoVale = vc.tipoVale " & _
                        "Inner Join PeriodosINV pi on vc.idEmpresa = pi.idEmpresa And vc.idPeriodoINV = pi.idPeriodoINV And vc.idSucursal = pi.idSucursal " & _
                        "Inner Join productos p ON vd.idEmpresa = p.idEmpresa AND vd.idProducto = p.idProducto   " & _
                        "Where vc.idEmpresa = '" & glsEmpresa & "' And vc.idSucursal ='08090001' And vc.idAlmacen = '11010001' And vc.estValeCab <> 'ANU'   And p.GlsProducto Like '%" & txtBus_Producto.Text & "%'  And pi.estPeriodoInv = 'ACT' " & _
                        "Group by vd.idProducto,vd.idEmpresa,vd.idSucursal) P01 " & _
                        "ON p.idEmpresa = P01.idEmpresa   And p.idProducto = P01.idProducto   And p.idUMCompra = P01.idUM " & _
                        "Left Join (Select idAlmacen, p.idProducto, item, idUM, vc.idEmpresa, vc.idSucursal, " & _
                        "Format(sum(if(vd.tipovale = 'I',Cantidad,Cantidad * -1)),2) as  StkDansey, vvUnit " & _
                        "From  ValesDet vd " & _
                        "Inner Join ValesCab vc " & _
                        "On vd.idValesCab = vc.idValesCab   And vd.idEmpresa = vc.idEmpresa   And vd.idSucursal = vc.idSucursal   And vd.tipoVale = vc.tipoVale " & _
                        "Inner Join PeriodosINV pi on vc.idEmpresa = pi.idEmpresa And vc.idPeriodoINV = pi.idPeriodoINV And vc.idSucursal = pi.idSucursal " & _
                        "Inner Join productos p ON vd.idEmpresa = p.idEmpresa AND vd.idProducto = p.idProducto  " & _
                        "Where vc.idEmpresa = '01' And vc.idSucursal ='08090004' And vc.idAlmacen = '11030001' And vc.estValeCab <> 'ANU' And pi.estPeriodoInv = 'ACT' And p.GlsProducto Like '%" & txtBus_Producto.Text & "%'   And pi.estPeriodoInv = 'ACT' " & _
                        "Group by vd.idProducto,vd.idEmpresa,vd.idSucursal) P03 " & _
                        "On p.idEmpresa = P03.idEmpresa   And p.idProducto = P03.idProducto   And p.idUMCompra = P03.idUM "
            
        gListaPrecio.Columns.ColumnByFieldName("StkPrincipal").Visible = True
        gListaPrecio.Columns.ColumnByFieldName("StkDansey").Visible = True
        cCadColAlmacenes = ",P01.StkPrincipal,P03.StkDansey  "
           
    Else
        cCadStkAlmacenes = ""
        cCadColAlmacenes = ""
        gListaPrecio.Columns.ColumnByFieldName("StkPrincipal").Visible = False
        gListaPrecio.Columns.ColumnByFieldName("StkDansey").Visible = False
    End If
    
    csql = "Select " & strCampos & " (L.IdProducto + L.IdUM) As Item,L.IdProducto,P.GlsProducto,P.IdFabricante,uv.abreUM as UMCompra,p.AfectoIGV as Afecto," & _
            "l.idUM , ul.abreUM as GlsUM, r.factor, l.VVUnit as VVUnit, l.IGVUnit AS IGVUnit, l.PVUnit AS PVUnit, l.CostoUnit,l.FactorUnit,l.Factor2Unit " & cCadColAlmacenes & _
            "FROM preciosventa l " & _
            "INNER JOIN productos p ON l.idEmpresa = p.idEmpresa AND l.idProducto = p.idProducto " & _
            "INNER JOIN unidadMedida uv ON p.idUMCompra = uv.idUM " & _
            "INNER JOIN unidadMedida ul ON l.idUM       = ul.idUM " & _
            "INNER JOIN presentaciones r ON l.idProducto  = r.idProducto AND l.idUM = r.idUM AND r.idEmpresa = '" & glsEmpresa & "' " & _
            strTablas & " " & cCadStkAlmacenes & _
            "WHERE l.idLista = '" & txtCod_Lista.Text & "' " & _
            "AND l.idEmpresa = '" & glsEmpresa & "' "
            
    If strCond <> "" Then csql = csql & strCond
    csql = csql & ccadagrupa
    csql = csql & " ORDER BY l.idProducto"
    
If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
    
Set gListaPrecio.DataSource = rsdatos

    With gListaPrecio
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "item"
   
        .m.ClearGroupColumns
        If .Ex.GroupColumnCount = 0 Then
            For i = 1 To NumNiveles
             .Columns.ColumnByName("GlsNivel" & Format(i, "00")).Caption = "Nivel:"
             .Columns.ColumnByName("GlsNivel" & Format(i, "00")).Visible = True
             .Columns.ColumnByName("GlsNivel" & Format(i, "00")).GroupIndex = 0
            Next
        End If
    End With
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
    
    strCodigo = Trim(txtCod_Lista.Text)
    
    '--- Validando si existe en ventas
    csql = "SELECT idLista FROM docventas WHERE idLista = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Ventas)"
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando preciosventa
    csql = "DELETE FROM preciosventa WHERE idLista = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    '--- Eliminando el registro
    csql = "DELETE FROM listaprecios WHERE idLista = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    '--- Nuevo
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    
    listaListaPrecio StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub
    
Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
