VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmAyudaProductos 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Productos"
   ClientHeight    =   9330
   ClientLeft      =   2175
   ClientTop       =   2775
   ClientWidth     =   14730
   Icon            =   "frmAyudaProductos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   14730
   Begin VB.CommandButton cmbProdOtrasSucursales 
      Caption         =   "Consultar en otras sucursales"
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8595
      Width           =   2475
   End
   Begin VB.Frame fraContenido 
      Appearance      =   0  'Flat
      Caption         =   " Filtros "
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
      Height          =   1815
      Left            =   90
      TabIndex        =   9
      Top             =   720
      Width           =   11445
      Begin VB.Frame fraNivel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   8625
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   0
            Left            =   8100
            Picture         =   "frmAyudaProductos.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   1
            Left            =   8100
            Picture         =   "frmAyudaProductos.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   360
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   2
            Left            =   8100
            Picture         =   "frmAyudaProductos.frx":0720
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   720
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   3
            Left            =   8100
            Picture         =   "frmAyudaProductos.frx":0AAA
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1080
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   4
            Left            =   8100
            Picture         =   "frmAyudaProductos.frx":0E34
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1440
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   0
            Left            =   1305
            TabIndex        =   18
            Tag             =   "TidNivelPred"
            Top             =   30
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
            Container       =   "frmAyudaProductos.frx":11BE
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   19
            Top             =   30
            Width           =   5790
            _ExtentX        =   10213
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
            Container       =   "frmAyudaProductos.frx":11DA
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   1
            Left            =   1305
            TabIndex        =   20
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
            Container       =   "frmAyudaProductos.frx":11F6
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   21
            Top             =   390
            Width           =   5790
            _ExtentX        =   10213
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
            Container       =   "frmAyudaProductos.frx":1212
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   2
            Left            =   1305
            TabIndex        =   22
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
            Container       =   "frmAyudaProductos.frx":122E
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   2
            Left            =   2280
            TabIndex        =   23
            Top             =   750
            Width           =   5790
            _ExtentX        =   10213
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
            Container       =   "frmAyudaProductos.frx":124A
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   3
            Left            =   1305
            TabIndex        =   24
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
            Container       =   "frmAyudaProductos.frx":1266
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   3
            Left            =   2280
            TabIndex        =   25
            Top             =   1110
            Width           =   5790
            _ExtentX        =   10213
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
            Container       =   "frmAyudaProductos.frx":1282
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   4
            Left            =   1305
            TabIndex        =   26
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
            Container       =   "frmAyudaProductos.frx":129E
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   4
            Left            =   2280
            TabIndex        =   27
            Top             =   1470
            Width           =   5790
            _ExtentX        =   10213
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
            Container       =   "frmAyudaProductos.frx":12BA
            Vacio           =   -1  'True
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
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
            Index           =   0
            Left            =   135
            TabIndex        =   32
            Top             =   45
            Width           =   345
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
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
            Index           =   1
            Left            =   135
            TabIndex        =   31
            Top             =   405
            Width           =   345
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
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
            Index           =   2
            Left            =   135
            TabIndex        =   30
            Top             =   765
            Width           =   345
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
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
            Index           =   3
            Left            =   135
            TabIndex        =   29
            Top             =   1125
            Width           =   345
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
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
            Index           =   4
            Left            =   135
            TabIndex        =   28
            Top             =   1485
            Width           =   345
         End
      End
      Begin CATControls.CATTextBox TxtBusq 
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   1140
         Width           =   6750
         _ExtentX        =   11906
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
         Container       =   "frmAyudaProductos.frx":12D6
         Vacio           =   -1  'True
      End
      Begin VB.Label lblBusq 
         Appearance      =   0  'Flat
         Caption         =   "Producto"
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
         Height          =   240
         Left            =   195
         TabIndex        =   10
         Top             =   1200
         Width           =   795
      End
   End
   Begin VB.Frame fraPresentaciones 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   60
      TabIndex        =   8
      Top             =   6930
      Width           =   14565
      Begin DXDBGRIDLibCtl.dxDBGrid gPresentaciones 
         Height          =   1305
         Left            =   60
         OleObjectBlob   =   "frmAyudaProductos.frx":12F2
         TabIndex        =   33
         Top             =   180
         Width           =   14475
      End
   End
   Begin VB.Frame fraTipoProd 
      Appearance      =   0  'Flat
      Caption         =   " Tipo "
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
      Height          =   1815
      Left            =   11610
      TabIndex        =   4
      Top             =   720
      Width           =   3045
      Begin VB.OptionButton opt_MateriaPrima 
         Caption         =   "Materia Prima"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   870
         TabIndex        =   7
         Top             =   1140
         Width           =   1290
      End
      Begin VB.OptionButton opt_Servicios 
         Caption         =   "Servicios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   870
         TabIndex        =   6
         Top             =   795
         Width           =   1065
      End
      Begin VB.OptionButton opt_Producto 
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   870
         TabIndex        =   5
         Top             =   450
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton OptFormulas 
         Caption         =   "Fórmulas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         TabIndex        =   38
         Top             =   1305
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Frame fraGrilla 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   90
      TabIndex        =   3
      Top             =   2520
      Width           =   14565
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   3780
         Left            =   90
         OleObjectBlob   =   "frmAyudaProductos.frx":3C62
         TabIndex        =   37
         Top             =   225
         Width           =   14385
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   750
      Top             =   4500
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
            Picture         =   "frmAyudaProductos.frx":8C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":900E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":9460
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":97FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":9B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":9F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":A2C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":A662
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":A9FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":AD96
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":B130
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProductos.frx":BDF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   1164
      ButtonWidth     =   2858
      ButtonHeight    =   1005
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "          Aceptar          "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Otros Productos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "F6 para navegar entre controles"
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
      Height          =   285
      Left            =   9945
      TabIndex        =   34
      Top             =   8955
      Width           =   4680
   End
   Begin VB.Label lblPresentaciones 
      Appearance      =   0  'Flat
      Caption         =   "Otras Presentaciones:"
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
      Height          =   240
      Left            =   45
      TabIndex        =   11
      Top             =   6660
      Width           =   3435
   End
   Begin VB.Label LblReg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "(0) Registros"
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
      Height          =   285
      Left            =   12690
      TabIndex        =   2
      Top             =   6615
      Width           =   1905
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Presionar Enter en el registro para obtener el resultado "
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
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   8955
      Width           =   4680
   End
End
Attribute VB_Name = "FrmAyudaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SqlAdic         As String
Private sqlBus          As String
Private sqlCond         As String
Private SRptBus(3)      As String
Private EsNuevo         As Boolean
Private indAlmacen      As Boolean
Private indValidaStock  As Boolean
Private indPedido       As Boolean
Private strCodAlmacen   As String
Private indUMVenta      As Boolean
Private indMostrarPresentaciones As Boolean
Private strCodLista     As String
Private indMovNivel     As Boolean
Private intFoco         As Integer '0 = Texto,1 = Grilla productos, 2 = Grilla presentaciones
Private rsg             As New ADODB.Recordset
Private CodDcto         As String
Private IndAgrega       As Boolean
Private CIdCliente      As String
Private CCodMotivo      As String

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

Private Sub cmbProdOtrasSucursales_Click()
On Error GoTo Err
Dim strCodProd As String, StrMsgError As String

    strCodProd = G.Columns.ColumnByFieldName("idProducto").Value
    frmProdOtrasSucursales.MostrarForm strCodProd, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Deactivate()

    SqlAdic = ""
    If EsNuevo = False Then
        TxtBusq.Text = ""
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 27
            Unload Me
        Case 117
            Select Case intFoco
            Case 0
                G.SetFocus
            Case 1
                gPresentaciones.SetFocus
            Case 2
                TxtBusq.SetFocus
            End Select
    End Select

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
Dim intFor As Integer
      
    IndAgrega = True
    indEvaluaVacio = False
    
    Me.Caption = "Ayuda de productos"
    Me.top = 0
    Me.left = 0
    
    ConfGrid G, True, False, False, False
    ConfGrid gPresentaciones, False, False, False, False
    EsNuevo = True
 
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    IndAgrega = False
    SqlAdic = ""
    TxtBusq.Text = ""
    
End Sub

Private Sub g_GotFocus()
    
    intFoco = 1

End Sub

Private Sub G_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrMsgError As String
    
    listaOtrasPresentaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub g_OnDblClick()

    g_OnKeyDown 13, 1

End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    Select Case KeyCode
        Case 13:
            SRptBus(0) = G.Columns.ColumnByFieldName("idProducto").Value
            SRptBus(1) = G.Columns.ColumnByFieldName("GlsProducto").Value '+ Space(150) + Trim(CStr(g.Columns.ColumnByFieldName("cod").Value))
            SRptBus(2) = G.Columns.ColumnByFieldName("idUMVenta").Value
        
            G.Dataset.Edit
            G.Columns.ColumnByFieldName("CHK").Value = 1
            G.Dataset.Post
            
            G.Dataset.Close
            G.Dataset.Active = False
            
            Me.Hide
       
        Case 27
            Unload Me
        Case 117
            Select Case intFoco
                Case 0
                    G.SetFocus
                Case 1
                    gPresentaciones.SetFocus
                Case 2
                    TxtBusq.SetFocus
            End Select
    End Select

End Sub

Private Sub gPresentaciones_GotFocus()
    
    intFoco = 2

End Sub

Private Sub gPresentaciones_OnDblClick()

    gPresentaciones_OnKeyDown 13, 1
    
End Sub

Private Sub gPresentaciones_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    Select Case KeyCode
        Case 13:
            SRptBus(0) = G.Columns.ColumnByFieldName("idProducto").Value
            SRptBus(1) = G.Columns.ColumnByFieldName("GlsProducto").Value '+ Space(150) + Trim(CStr(g.Columns.ColumnByFieldName("cod").Value))
            SRptBus(2) = gPresentaciones.Columns.ColumnByFieldName("idUM").Value
            
            G.Dataset.Close
            G.Dataset.Active = False
            Me.Hide
        Case 27
            Unload Me
        Case 117
            Select Case intFoco
                Case 0
                    G.SetFocus
                Case 1
                    gPresentaciones.SetFocus
                Case 2
                    TxtBusq.SetFocus
            End Select
    End Select

End Sub

Private Sub opt_MateriaPrima_Click()
Dim StrMsgError As String
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub opt_Producto_Click()
Dim StrMsgError As String
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub opt_Servicios_Click()
Dim StrMsgError As String
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub OptFormulas_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
Dim strCod1 As String
Dim strDes1 As String
On Error GoTo Err

    Select Case Button.Index
        Case 1  'Aceptar
            If G.Count > 0 Then
                G.Dataset.First
                G.Dataset.Filtered = False
                G.Dataset.Refresh
                Me.Hide
            End If
        Case 2  'Otros Productos

        Case 3  'Salir
            Me.Hide
    End Select
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtBusq_Change()
Dim StrMsgError As String

    If EsNuevo = False Then
        If glsEnterAyudaProductos = False Then
            
            fill StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        End If
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtBusq_GotFocus()

    intFoco = 0

End Sub

Private Sub TxtBusq_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyDown Then G.SetFocus
    If EsNuevo = True Then TxtBusq.SelStart = Len(TxtBusq.Text) + 1

End Sub

Private Sub TxtBusq_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyAscii = 13 Then
        fill StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    EsNuevo = False

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub fill(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsdatos         As New ADODB.Recordset
Dim Orden           As String
Dim Valida          As String
Dim intFor          As Integer
Dim StrtipProduct   As String

    Valida = "N"
    'sqlBus = setSqlAlm(strCodAlmacen)
    
'    If Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "ORDERNAR_AYUDA_PRODUCTOS", True)) = "S" Then
'        Orden = " order by 5 "
'    Else
        Orden = " order by 2"
'    End If
            
    sqlCond = sqlBus + " like '%" & Trim(TxtBusq.Text) & "%' OR P.IdProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') "
    If glsNumNiveles > 0 Then
        If txtCod_Nivel(glsNumNiveles - 1).Text <> "" Then
            sqlCond = sqlCond & " AND idNivel = '" & txtCod_Nivel(glsNumNiveles - 1).Text & "'"
        End If
    End If
    If opt_Producto.Value Then
        'sqlCond = sqlCond & " AND idTipoProducto = '06001'"
        StrtipProduct = "06001"
    ElseIf opt_MateriaPrima.Value Then
        'sqlCond = sqlCond & " AND idTipoProducto = '06003'"
        StrtipProduct = "06003"
    ElseIf opt_Servicios.Value Then
        StrtipProduct = "06002"
    ElseIf OptFormulas.Value Then
        'sqlCond = sqlCond & " AND idTipoProducto = '06004'"
        StrtipProduct = "06004"
    End If
    sqlCond = sqlCond & " AND estProducto = 'A' "
    sqlCond = sqlCond & SqlAdic
    
'    If Trim("" & traerCampo("Empresas", "RUC", "idEmpresa", glsEmpresa, False)) = "20257354041" Then
'        sqlCond = sqlCond & " Group by p.idproducto " & Orden
'    Else
'        sqlCond = sqlCond & Orden
'    End If
    sqlCond = "EXEC spu_Docventa_Lista_Productos '" & glsEmpresa & "','" & strCodLista & "','" & strCodAlmacen & "','" & StrtipProduct & "','" & Trim("" & txtCod_Nivel(glsNumNiveles - 1).Text) & "','%" & Trim(TxtBusq.Text) & "%'"
    If rsdatos.State = 1 Then rsdatos.Close
    rsdatos.Open sqlCond, Cn, adOpenStatic, adLockReadOnly
    If rsg.State = 1 Then rsg.Close
    
    rsg.Fields.Append "CHK", adInteger, adFldIsNullable
    rsg.Fields.Append "idProducto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 350, adFldIsNullable
    rsg.Fields.Append "GlsMarca", adVarChar, 80, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 80, adFldIsNullable
    rsg.Fields.Append "GlsMoneda", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsTallaPeso", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "Afecto", adVarChar, 5, adFldIsNullable
    rsg.Fields.Append "idUMVenta", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "Stock", adDouble, 11, adFldIsNullable
    rsg.Fields.Append "idFabricante", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GLSDOCREFERENCIA", adVarChar, 200, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 200, adFldIsNullable
    
    If STR_VALIDA_SEPARACION = "S" Then
        If CodDcto = "40" Or CodDcto = "92" Then
            Valida = "S"
            rsg.Fields.Append "Separacion", adDouble, 11, adFldIsNullable
            rsg.Fields.Append "Disponible", adDouble, 11, adFldIsNullable
        End If
    End If
    rsg.Open
    
    If rsdatos.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("CHK") = 0
        rsg.Fields("idProducto") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("GlsMarca") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("GlsMoneda") = ""
        rsg.Fields("GlsTallaPeso") = 0#
        rsg.Fields("Afecto") = 0#
        rsg.Fields("idUMVenta") = ""
        rsg.Fields("Stock") = Val(Format(0, "0.00"))
        rsg.Fields("idFabricante") = ""
        rsg.Fields("GLSDOCREFERENCIA") = ""
        rsg.Fields("CodigoRapido") = ""
        
    Else
        Do While Not rsdatos.EOF
            item = (rsg.RecordCount) + 1
            rsg.AddNew
            rsg.Fields("CHK") = 0
            rsg.Fields("idProducto") = "" & rsdatos.Fields("idProducto")
            rsg.Fields("GlsProducto") = "" & rsdatos.Fields("GlsProducto")
            rsg.Fields("GlsMarca") = "" & rsdatos.Fields("GlsMarca")
            rsg.Fields("GlsUM") = "" & rsdatos.Fields("GlsUM")
            rsg.Fields("GlsMoneda") = "" & rsdatos.Fields("GlsMoneda")
            rsg.Fields("GlsTallaPeso") = "" & rsdatos.Fields("GlsTallaPeso")
            rsg.Fields("Afecto") = rsdatos.Fields("Afecto")
            rsg.Fields("Stock") = Val(Format(rsdatos.Fields("Stock"), "0.00"))
            rsg.Fields("CodigoRapido") = rsdatos.Fields("CodigoRapido")
            rsg.Fields("IdFabricante") = rsdatos.Fields("IdFabricante")
            If Valida = "S" Then
                If CodDcto = "40" Or CodDcto = "92" Then
                    rsg.Fields("Separacion") = Val(Format(rsdatos.Fields("Separacion"), "0.00"))
                    rsg.Fields("Disponible") = Val(Format(rsdatos.Fields("Disponible"), "0.00"))
                End If
            End If
            rsdatos.MoveNext
       Loop
    End If
    Set G.DataSource = Nothing
    
    With G
        .DefaultFields = False
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        Set .DataSource = rsg
        .Dataset.Active = True
        .KeyField = "idProducto"
        .Dataset.Edit
        .Dataset.Post
    End With
    
    For intFor = 0 To G.Columns.Count - 1
        G.m.ApplyBestFit G.Columns(intFor)
    Next intFor
    
    Set rsdatos = Nothing
    If Not G.Dataset.EOF Then
        G.Dataset.First
    End If
    
    LblReg.Caption = "(" + Format(G.Count, "0") + ")Registros"
    listaOtrasPresentaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub CargarProductosCliente(ByRef StrMsgError As String)
On Error GoTo Err
Dim CSqlC                           As String
Dim rsdatos                         As New ADODB.Recordset
Dim CIdTipoProducto                 As String
    
    If opt_Producto.Value Then
        CIdTipoProducto = "06001"
    Else
        CIdTipoProducto = "06003"
    End If
    
    CSqlC = "Select A.IdEmpresa,A.IdProducto,A.CodigoRapido,A.GlsProducto,M.GlsMarca,A.idUMCompra IdUMVenta,U.GlsUM,O.IdMoneda GlsMoneda," & _
            "If(A.AfectoIGV = 1,'S','N') Afecto,0 Separacion,0 Disponible,T.GlsTallaPeso,A.IdFabricante,A.IdTipoProducto," & _
            "A.EstProducto,0 Stock " & _
            "From Productos A " & _
            "Inner Join(" & _
                "Select B.IdEmpresa,B.IdProducto " & _
                "From DocVentas A " & _
                "Inner Join DocVentasDet B " & _
                    "On A.IdEmpresa = B.IdEmpresa And A.IdSucursal = B.IdSucursal And A.IdDocumento = B.IdDocumento And A.IdSerie = B.IdSerie " & _
                    "And A.IdDocVentas = B.IdDocVentas " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdPerCliente = '" & CIdCliente & "' " & _
                "Group By B.IdEmpresa,B.IdProducto" & _
            ") C " & _
                "On A.IdEmpresa = C.Idempresa And A.IdProducto = C.IdProducto " & _
            "Inner Join Marcas M " & _
                "On A.IdEmpresa = M.IdEmpresa And A.IdMarca = M.IdMarca " & _
            "Inner Join UnidadMedida U " & _
                "On A.IdUMCompra = U.IdUM " & _
            "Inner Join Monedas O " & _
                "On A.IdMoneda = O.IdMoneda " & _
            "Left Join TallaPeso T " & _
                "On A.IdEmpresa = T.IdEmpresa And A.IdTallaPeso = T.IdTallaPeso " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdTipoProducto = '" & CIdTipoProducto & "' And A.EstProducto = 'A' " & _
            "And (A.GlsProducto Like '%" & TxtBusq.Text & "%' Or A.IdProducto Like '%" & TxtBusq.Text & "%' Or A.CodigoRapido Like '%" & TxtBusq.Text & "%' Or A.IdFabricante Like '%" & TxtBusq.Text & "%') " & _
            "Order By 2"
    
    rsdatos.Open CSqlC, strcn, adOpenStatic, adLockReadOnly
    If rsg.State = 1 Then rsg.Close
    
    rsg.Fields.Append "CHK", adInteger, adFldIsNullable
    rsg.Fields.Append "idProducto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 350, adFldIsNullable
    rsg.Fields.Append "GlsMarca", adVarChar, 80, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 80, adFldIsNullable
    rsg.Fields.Append "GlsMoneda", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsTallaPeso", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "Afecto", adVarChar, 5, adFldIsNullable
    rsg.Fields.Append "idUMVenta", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "Stock", adDouble, 11, adFldIsNullable
    rsg.Fields.Append "idFabricante", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GLSDOCREFERENCIA", adVarChar, 200, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 200, adFldIsNullable
    
    If Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "VALIDA_SEPARACION", True) & "") = "S" Then
        If CodDcto = "40" Or CodDcto = "92" Then
            Valida = "S"
            rsg.Fields.Append "Separacion", adDouble, 11, adFldIsNullable
            rsg.Fields.Append "Disponible", adDouble, 11, adFldIsNullable
        End If
    End If
    rsg.Open
    
    If rsdatos.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("CHK") = 0
        rsg.Fields("idProducto") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("GlsMarca") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("GlsMoneda") = ""
        rsg.Fields("GlsTallaPeso") = 0#
        rsg.Fields("Afecto") = 0#
        rsg.Fields("idUMVenta") = ""
        rsg.Fields("Stock") = ""
        rsg.Fields("idFabricante") = ""
        rsg.Fields("GLSDOCREFERENCIA") = ""
        rsg.Fields("CodigoRapido") = ""
        
    Else
        Do While Not rsdatos.EOF
            item = (rsg.RecordCount) + 1
            rsg.AddNew
            rsg.Fields("CHK") = 0
            rsg.Fields("idProducto") = "" & rsdatos.Fields("idProducto")
            rsg.Fields("GlsProducto") = "" & rsdatos.Fields("GlsProducto")
            rsg.Fields("GlsMarca") = "" & rsdatos.Fields("GlsMarca")
            rsg.Fields("GlsUM") = "" & rsdatos.Fields("GlsUM")
            rsg.Fields("GlsMoneda") = "" & rsdatos.Fields("GlsMoneda")
            rsg.Fields("GlsTallaPeso") = "" & rsdatos.Fields("GlsTallaPeso")
            rsg.Fields("Afecto") = rsdatos.Fields("Afecto")
            rsg.Fields("Stock") = Val(Format(rsdatos.Fields("Stock"), "0.00"))
            rsg.Fields("CodigoRapido") = rsdatos.Fields("CodigoRapido")
            rsg.Fields("IdFabricante") = rsdatos.Fields("IdFabricante")
            If Valida = "S" Then
                If CodDcto = "40" Or CodDcto = "92" Then
                    rsg.Fields("Separacion") = Val(Format(rsdatos.Fields("Separacion"), "0.00"))
                    rsg.Fields("Disponible") = Val(Format(rsdatos.Fields("Disponible"), "0.00"))
                End If
            End If
            rsdatos.MoveNext
       Loop
    End If
    Set G.DataSource = Nothing
    
    With G
        .DefaultFields = False
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        Set .DataSource = rsg
        .Dataset.Active = True
        .KeyField = "idProducto"
        .Dataset.Edit
        .Dataset.Post
    End With
    
    For intFor = 0 To G.Columns.Count - 1
        G.m.ApplyBestFit G.Columns(intFor)
    Next intFor
    
    Set rsdatos = Nothing
    If Not G.Dataset.EOF Then
        G.Dataset.First
    End If
    
    LblReg.Caption = "(" + Format(G.Count, "0") + ")Registros"
    listaOtrasPresentaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Public Sub Execute(ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
'    Dim strMsgError As String
'    Dim intI As Integer
'    pblnAceptar = False
'    MousePointer = 0
'
'    SRptBus(0) = ""
'
'    SqlAdic = strParAdic
'
'    sqlBus = setSql(strParAyuda)
'
'    fill
'
'    Me.Show vbModal
'    If SRptBus(0) <> "" Then
'        TextBox1.Text = SRptBus(0)
'        TextBox2.Text = SRptBus(1)
'    End If
End Sub

Private Function setSqlAlm(strAlm As String) As String
Dim strCampoUM          As String
Dim strStockUM          As String
Dim strCantidad         As String
Dim StrSeparacion       As String
Dim Strdisponible       As String
Dim strTablaPresentaciones As String
Dim cadarmasaldo        As String
Dim ccadstk             As String
Dim strTipoProducto     As String
Dim nPC                 As String
Dim NIdStock            As String
Dim StrWhereMotivo      As String

    cadarmasaldo = ""
    ccadstk = ""
    
    If glsVisualizaCodFab = "N" Then
        G.Columns.ColumnByFieldName("IdFabricante").Visible = False
    End If
    
'    If Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "VIZUALIZA_CODIGO_RAPIDO", True)) = "S" Then
'        G.Columns.ColumnByFieldName("CodigoRapido").Visible = True
'    End If
    
    strCampoUM = "idUMVenta"
    strStockUM = "CantidadStockUV"
    strCantidad = "(a.CantidadStock / f.Factor )" 'Es la cantidad de venta
    
    If STR_VALIDA_SEPARACION = "S" Then
        If CodDcto = "40" Or CodDcto = "92" Then
            G.Columns.ColumnByFieldName("Separacion").Visible = True
            G.Columns.ColumnByFieldName("Disponible").Visible = True
            
            StrSeparacion = "(a.Separacion)"  'Es la Separacion
            Strdisponible = "((a.CantidadStock) -  if(left(a.Separacion,1)='-',0,a.Separacion))" 'Stock Disponible
        End If
    End If
    
    strTablaPresentaciones = " INNER JOIN presentaciones f ON p.idEmpresa = f.idEmpresa AND p.idProducto = f.idProducto AND p." & strCampoUM & " = f.idUM "
        
    If indUMVenta = False Then
        strCampoUM = "idUMCompra"
        strStockUM = "CantidadStockUC"
        strCantidad = "a.CantidadStock" 'Es la cantidad de compra
        strTablaPresentaciones = ""
    End If
    
    G.Columns.ColumnByFieldName("Stock").Visible = False
    
    If opt_Servicios.Value = False Then
        If OptFormulas.Value Then
            setSqlAlm = "Select P.IdProducto,P.GlsProducto,'' As GlsMarca,P.IdUMVenta,U.GlsUM,O.IdMoneda As GlsMoneda,P.IdFabricante," & _
                        "CASE WHEN P.AfectoIGV = 1 THEN 'S' ELSE 'N' END Afecto,CAST(0 AS DECIMAL(12,2)) As Stock,'' As GlsTallaPeso,'' As CodigoRapido " & _
                        "From Productos P " & _
                        "Inner Join ComboCab C " & _
                            "On P.IdEmpresa = C.IdEmpresa And P.IdProducto = C.IdComboCab " & _
                        "Inner Join Monedas O " & _
                            "On P.IdMoneda = O.IdMoneda " & _
                        "Inner Join UnidadMedida U " & _
                            "On P.IdUMVenta = U.IdUM " & _
                        "Where P.IdEmpresa = '" & glsEmpresa & "' And P.IdTipoProducto = '06004' And (P.GlsProducto "
            
        Else
            If indPedido = False Then
                If STR_VALIDA_SEPARACION = "S" Then
'                    If CodDcto = "40" Or CodDcto = "92" Then
'                        If Trim("" & traerCampo("Empresas", "RUC", "idEmpresa", glsEmpresa, False)) = "20430471750" Then
'                            If opt_Producto.Value Then
'                                strTipoProducto = "06001"
'                            ElseIf opt_MateriaPrima.Value Then
'                                strTipoProducto = "06003"
'                            ElseIf OptFormulas.Value Then
'                                strTipoProducto = "06004"
'                            End If
'
'                            setSqlAlm = "SELECT p.idempresa,a.idsucursal,p.idProducto,p.CodigoRapido,p.GlsProducto,m.GlsMarca,p.idUMCompra AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda,CASE WHEN p.afectoIGV = 1 THEN 'S' ELSE 'N' END Afecto,CAST(a.Separacion AS DECIAML(12,2)) as Separacion,CAST((s.Stock -  CASE WHEN left(a.Separacion,1) = '-' THEN 0 ELSE a.Separacion END) AS NUMERIC(12,2)) as Disponible,t.GlsTallaPeso , p.idfabricante, p.idTipoProducto, p.estProducto,CAST(s.Stock AS NUMERIC(12,2)) as Stock " & _
'                                       "FROM productos p " & _
'                                       "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
'                                       "INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM " & _
'                                       "INNER JOIN monedas o ON p.idMoneda = o.idMoneda " & _
'                                       "INNER JOIN productosalmacen a ON p.idEmpresa = a.idEmpresa " & _
'                                       "AND p.idProducto = a.idProducto AND p.idUMCompra = a.idUMCompra " & _
'                                       "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
'                                       "Inner Join " & _
'                                       "( " & _
'                                       "Select P.idEmpresa,IsNull(vd.idSucursal,'') Idsucursal,P.idProducto, " & _
'                                       "sum(If(vd.idempresa is null,0,If(vc.idSucursal = '" & glsSucursal & "' And vc.estValeCab <> 'ANU',if(vd.tipovale = 'I',Cantidad,Cantidad * -1),0))) as Stock " & _
'                                       "From Productos P " & _
'                                       "Left Join ValesDet vd On P.IdEmpresa = vd.IdEmpresa And P.IdProducto = vd.IdProducto " & _
'                                       "left Join Valescab vc On vd.idEmpresa = vc.idEmpresa And vd.idSucursal = vc.idSucursal And vd.tipoVale = vc.tipoVale And vd.idValesCab = vc.idValesCab " & _
'                                       "left Join PeriodosINV pi on vc.idEmpresa = pi.idEmpresa And vc.idPeriodoINV = pi.idPeriodoINV And vc.idSucursal = pi.idSucursal " & _
'                                       "Where P.idEmpresa = '" & glsEmpresa & "' And (vc.IdAlmacen = '" & strAlm & "' Or '' = '" & strAlm & "') AND idTipoProducto = '" & strTipoProducto & "'  AND estProducto = 'A' AND (p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%')And (pi.estPeriodoInv = 'ACT' Or pi.estPeriodoInv Is Null) " & _
'                                       "Group bY P.idEmpresa,P.idProducto order by P.idEmpresa,P.idProducto " & _
'                                       ") S " & _
'                                       "On P.idEmpresa = S.idEmpresa And P.idProducto = S.idProducto " & _
'                                       "Where p.idEmpresa = '" & glsEmpresa & "' And a.idSucursal = '" & glsSucursal & "'  " & _
'                                       "AND (p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') " & _
'                                       "AND idTipoProducto =  '" & strTipoProducto & "' AND estProducto = 'A' " & _
'                                       "AND p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa " & _
'                                       "WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "')  "
'
'                            If indValidaStock Then
'                                setSqlAlm = setSqlAlm & "AND " & strCantidad & "  > 0 "
'                            End If
'                            g.Columns.ColumnByFieldName("Stock").Visible = True
'                            setSqlAlm = setSqlAlm & "AND (p.GlsProducto "
'
'                        Else
'                            setSqlAlm = "SELECT p.CodigoRapido,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto, Format(" & strCantidad & ",2) as Stock,Format(" & StrSeparacion & ",2) as Separacion,Format(" & Strdisponible & ",2) as Disponible, t.GlsTallaPeso, p.idfabricante " & _
'                                        "FROM productos p " & _
'                                             "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
'                                             "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
'                                             "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
'                                              cadarmasaldo & _
'                                             "INNER JOIN productosalmacen a ON p.idEmpresa = a.idEmpresa AND a.idSucursal = '" & glsSucursal & "' " & _
'                                                                           "AND p.idProducto = a.idProducto " & _
'                                                                           "AND p." & strCampoUM & " = a.idUMCompra " & strTablaPresentaciones & _
'                                             "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
'                                        "WHERE p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "')" & _
'                                        "AND p.idEmpresa = '" & glsEmpresa & "' "
'
'                            If indValidaStock Then
'                                 setSqlAlm = setSqlAlm & "AND " & strCantidad & "  > 0 "
'                            End If
'                            g.Columns.ColumnByFieldName("Stock").Visible = True
'                            setSqlAlm = setSqlAlm & "AND (p.GlsProducto "
'                        End If
'
'                    Else
'                        If Trim("" & traerCampo("Empresas", "RUC", "idEmpresa", glsEmpresa, False)) = "20430471750" Then
'                            setSqlAlm = "SELECT p.CodigoRapido,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto, st.Stock,t.GlsTallaPeso, p.idfabricante " & _
'                                          "FROM productos p " & _
'                                          "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
'                                          "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
'                                          "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
'                                          "Left Join (Select format(sum(if(vd.tipovale = 'I',Cantidad,Cantidad * -1)),2) as Stock,vd.idEmpresa,vd.idSucursal,vd.idProducto " & _
'                                          "From ValesDet vd " & _
'                                          "Inner Join Valescab vc   On vd.idValesCab = vc.idValesCab   And vd.idEmpresa = vc.idEmpresa   And vd.idSucursal = vc.idSucursal   And vd.tipoVale = vc.tipoVale " & _
'                                          "Inner Join PeriodosINV pi   on vc.idEmpresa = pi.idEmpresa  And vc.idPeriodoINV = pi.idPeriodoINV   And vc.idSucursal = pi.idSucursal " & _
'                                          "Where vc.idEmpresa = '" & glsEmpresa & "' And (vc.IdAlmacen = '" & strAlm & "' Or '' = '" & strAlm & "') And vc.idSucursal = '" & glsSucursal & "' And pi.estPeriodoInv = 'ACT' And vc.estValeCab <> 'ANU' " & _
'                                          "Group bY vd.idProducto,vd.idEmpresa,vd.idSucursal) st " & _
'                                          "On p.idProducto = st.idProducto And p.idEmpresa = st.idEmpresa " & _
'                                          "INNER JOIN productosalmacen a ON p.idEmpresa = a.idEmpresa AND a.idSucursal = '" & glsSucursal & "' " & _
'                                                                      "AND p.idProducto = a.idProducto " & _
'                                                                      "AND p." & strCampoUM & " = a.idUMCompra " & strTablaPresentaciones & _
'                                          "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
'                                          "WHERE p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "')" & _
'                                          "AND p.idEmpresa = '" & glsEmpresa & "' "
'
'                            If indValidaStock Then
'                                 setSqlAlm = setSqlAlm & "AND " & strCantidad & "  > 0 "
'                            End If
'                            g.Columns.ColumnByFieldName("Stock").Visible = True
'                            setSqlAlm = setSqlAlm & "AND (p.GlsProducto "
'
'                        Else
'                            setSqlAlm = "SELECT p.CodigoRapido,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto, Format(" & strCantidad & ",2) as Stock, t.GlsTallaPeso, p.idfabricante " & _
'                                        "FROM productos p " & _
'                                             "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
'                                             "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
'                                             "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
'                                             "INNER JOIN productosalmacen a ON p.idEmpresa = a.idEmpresa AND a.idSucursal = '" & glsSucursal & "' " & _
'                                                                           "AND a.idAlmacen = '" & strAlm & "' AND p.idProducto = a.idProducto " & _
'                                                                           "AND p." & strCampoUM & " = a.idUMCompra " & strTablaPresentaciones & _
'                                             "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
'                                        "WHERE p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "')" & _
'                                        "AND p.idEmpresa = '" & glsEmpresa & "' "
'
'                            If indValidaStock Then
'                                 setSqlAlm = setSqlAlm & "AND " & strCantidad & "  > 0 "
'                            End If
'                            g.Columns.ColumnByFieldName("Stock").Visible = True
'                            setSqlAlm = setSqlAlm & "AND (p.GlsProducto "
'                        End If
'                    End If
                
                Else
                    'If Trim("" & traerCampo("Empresas", "RUC", "idEmpresa", glsEmpresa, False)) = "20257354041" Then
'                        If opt_Producto.Value Then
'                            strTipoProducto = "06001"
'                        ElseIf opt_MateriaPrima.Value Then
'                            strTipoProducto = "06003"
'                        ElseIf OptFormulas.Value Then
'                            strTipoProducto = "06004"
'                        End If
'
'
'                        setSqlAlm = "SELECT p.idempresa,p.idProducto,p.CodigoRapido,p.GlsProducto,m.GlsMarca,p.idUMCompra AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto,0 as Separacion,Format(IfNull(s.Stock,0),2) as Disponible,t.GlsTallaPeso , p.idfabricante, p.idTipoProducto, p.estProducto,Format(IfNull(s.Stock,0),2) as Stock " & _
'                                   "FROM productos p " & _
'                                   "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
'                                   "INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM " & _
'                                   "INNER JOIN monedas o ON p.idMoneda = o.idMoneda " & _
'                                   "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
'                                   "Left Join " & _
'                                    "( " & _
'                                        "Select P.idEmpresa,IfNull(vd.idSucursal,'') Idsucursal,P.idProducto, " & _
'                                        "sum(if(vd.tipovale = 'I',Cantidad,Cantidad * -1)) Stock " & _
'                                        "From Valescab vc " & _
'                                        "Inner Join PeriodosINV pi on vc.idEmpresa = pi.idEmpresa And vc.idPeriodoINV = pi.idPeriodoINV And vc.idSucursal = pi.idSucursal " & _
'                                        "Inner Join ValesDet vd On vc.idEmpresa = vd.idEmpresa And vc.idSucursal = vd.idSucursal And vc.tipoVale = vd.tipoVale And vc.idValesCab = vd.idValesCab " & _
'                                        "Inner Join Productos P On P.IdEmpresa = vd.IdEmpresa And P.IdProducto = vd.IdProducto " & _
'                                        "Where vc.idEmpresa = '" & glsEmpresa & "' And vc.EstValeCab <> 'ANU' And (vc.IdAlmacen = '" & strAlm & "' Or '' = '" & strAlm & "') AND p.idTipoProducto = '" & strTipoProducto & "'  AND estProducto = 'A' AND (p.IdProducto like '%" & Trim(TxtBusq.Text) & "%' OR p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') And pi.estPeriodoInv = 'ACT' " & _
'                                        "Group bY P.idEmpresa,P.idProducto order by P.idEmpresa,P.idProducto " & _
'                                    ") S " & _
'                                    "On P.idEmpresa = S.idEmpresa And P.idProducto = S.idProducto " & _
'                                   "Where p.idEmpresa = '" & glsEmpresa & "' " & _
'                                   "AND (P.IdProducto like '%" & Trim(TxtBusq.Text) & "%' Or p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') " & _
'                                   "AND idTipoProducto =  '" & strTipoProducto & "' AND estProducto = 'A' " & _
'                                   "AND p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa " & _
'                                   "WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "')  "
'
'                        g.Columns.ColumnByFieldName("Stock").Visible = True
                                   
                    'Else
                                   
                            If opt_Producto.Value Then
                                strTipoProducto = "06001"
                            ElseIf opt_MateriaPrima.Value Then
                                strTipoProducto = "06003"
                            ElseIf OptFormulas.Value Then
                                strTipoProducto = "06004"
                            End If
                        
                        If NIdStock = "" Then
                    
                            nPC = ComputerName
                            nPC = Replace(nPC, "-", "")
                            nPC = Replace(nPC, "", Trim(nPC))
                            nPC = Trim(nPC)
            
                            NIdStock = Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_TS")
                            
                        End If
                        
                        '**************************** LUIS LARA 13/01/2018 ***************************
                        'Cn.Execute "Call Spu_CalculaStock('" & glsEmpresa & "','" & NIdStock & "','','0','" & glsSucursal & "','','','" & strAlm & "',SysDate(),'')"
                        '*****************************************************************************
                        
                         setSqlAlm = "SELECT p.CodigoRapido,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda,CASE WHEN p.afectoIGV = 1 THEN 'S' ELSE 'N' END Afecto, CAST(ISNULL(" & strCantidad & ",0) AS NUMERIC(12,2)) as Stock, t.GlsTallaPeso, p.idfabricante " & _
                           "FROM productos p " & _
                           "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                           "INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM " & _
                           "INNER JOIN monedas o ON p.idMoneda = o.idMoneda " & _
                           "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                           "Left Join " & _
                           "( " & _
                           "Select P.idEmpresa,IsNull(vd.idSucursal,'') Idsucursal,P.idProducto, " & _
                           "sum(CASE WHEN vd.tipovale = 'I' THEN Cantidad ELSE Cantidad * -1 END) as CantidadStock " & _
                           "From Valescab vc " & _
                           "Inner Join PeriodosINV pi on vc.idEmpresa = pi.idEmpresa And vc.idPeriodoINV = pi.idPeriodoINV And vc.idSucursal = pi.idSucursal " & _
                           "Left Join ValesDet vd On vc.idEmpresa = vd.idEmpresa And vc.idSucursal = vd.idSucursal And vc.tipoVale = vd.tipoVale And vc.idValesCab = vd.idValesCab " & _
                           "left Join Productos P On P.IdEmpresa = vd.IdEmpresa And P.IdProducto = vd.IdProducto " & _
                           "Where vc.idEmpresa = '" & glsEmpresa & "' And vc.EstValeCab <> 'ANU' And (vc.IdAlmacen = '" & strAlm & "' Or '' = '" & strAlm & "') AND p.idTipoProducto = '" & strTipoProducto & "'  AND estProducto = 'A' AND (p.IdProducto like '%" & Trim(TxtBusq.Text) & "%' OR p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') And pi.estPeriodoInv = 'ACT' " & _
                           "Group bY P.idEmpresa,vd.idSucursal,P.idProducto" & _
                           ") A " & _
                           "On P.idEmpresa = A.idEmpresa And P.idProducto = A.idProducto " & _
                           "Where p.idEmpresa = '" & glsEmpresa & "' " & _
                           "AND p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "')  AND (p.IdProducto like '%" & Trim(TxtBusq.Text) & "%' OR p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') " & _
                           "AND estProducto = 'A' "
                                   
                    '**************************** LUIS LARA 13/01/2018 ***************************
''''                        setSqlAlm = "SELECT p.CodigoRapido,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto, Format(A.Stock,2) as Stock, t.GlsTallaPeso, p.idfabricante " & _
''''                           "FROM productos p " & _
''''                           "Inner Join PreciosVenta V " & _
''''                                "On P.IdEmpresa = V.IdEmpresa And P.IdProducto = V.IdProducto And P.IdUMVenta = V.IdUM " & _
''''                           "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
''''                           "INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM " & _
''''                           "INNER JOIN monedas o ON p.idMoneda = o.idMoneda " & _
''''                           "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
''''                           "Left Join " & NIdStock & " A " & _
''''                           "On P.idEmpresa = A.idEmpresa And P.idProducto = A.idProducto " & _
''''                           "Where p.idEmpresa = '" & glsEmpresa & "' " & _
''''                           "And V.IdLista = '" & strCodLista & "' AND (p.IdProducto like '%" & Trim(TxtBusq.Text) & "%' OR p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') " & _
''''                           "AND estProducto = 'A' "
                    '*******************************************************************************
                    
                    If Trim("" & CCodMotivo) = "" Then
                        StrWhereMotivo = "AND p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "') "
                    Else
                        'If Val(Trim(traerCampo("motivostraslados", "indlprecio", "idmotivotraslado", Trim("" & CCodMotivo), False) & "")) = 1 Then
                        '    StrWhereMotivo = "AND p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "') "
                        'Else
                        '    StrWhereMotivo = ""
                        'End If
                    End If
                    
'''''''''''                    '********************** LUIS LARA 12/01/2018 **************************
'''''''''''                    setSqlAlm = "SELECT p.CodigoRapido,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda," & _
'''''''''''                                "if(p.afectoIGV = 1,'S','N') Afecto,Format(ifnull(XZ.sc_stock,0) + ifnull(s.Stock,0),2) as Stock,t.GlsTallaPeso , p.idfabricante " & _
'''''''''''                               "FROM productos p " & _
'''''''''''                               "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
'''''''''''                               "INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM " & _
'''''''''''                               "INNER JOIN monedas o ON p.idMoneda = o.idMoneda " & _
'''''''''''                               "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
'''''''''''                               "Left Join ( " & _
'''''''''''                               "Select P.idEmpresa,IfNull(vd.idSucursal,'') Idsucursal,P.idProducto,vc.idAlmacen, " & _
'''''''''''                               "sum(If(vd.idempresa is null,0,if(vd.tipovale = 'I',Cantidad,Cantidad * -1))) as Stock " & _
'''''''''''                               "From (SELECT * FROM Valescab vc WHERE vc.idAlmacen = '" & strAlm & "' AND DATE_FORMAT(vc.fechaemision, '%Y%m%d')  = DATE_FORMAT(sysdate(), '%Y%m%d')   ) vc " & _
'''''''''''                               "Inner Join ValesDet vd On vd.idEmpresa = vc.idEmpresa And vd.idSucursal = vc.idSucursal And vd.tipoVale = vc.tipoVale And vd.idValesCab = vc.idValesCab " & _
'''''''''''                               "inner join Productos P On P.IdEmpresa = vd.IdEmpresa And P.IdProducto = vd.IdProducto " & _
'''''''''''                               "Where P.idEmpresa = '" & glsEmpresa & "' AND idTipoProducto = '" & strTipoProducto & "'  AND estProducto = 'A' " & _
'''''''''''                               "AND vc.idSucursal = '" & glsSucursal & "' And vc.estValeCab <> 'ANU' " & _
'''''''''''                               "AND vc.idAlmacen = '" & strAlm & "' AND DATE_FORMAT(vc.fechaemision, '%Y%m%d')  = DATE_FORMAT(sysdate(), '%Y%m%d') " & _
'''''''''''                               "AND (p.idProducto like '%" & Trim(TxtBusq.Text) & "%' OR p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') " & _
'''''''''''                               "Group bY P.idEmpresa,P.idProducto,vc.idAlmacen order by P.idEmpresa,P.idProducto " & _
'''''''''''                               ") S On P.idEmpresa = S.idEmpresa And P.idProducto = S.idProducto " & _
'''''''''''                               "Left Join (SELECT sc_periodo,sc_codalm,sc_codart,sc_stock,idempresa FROM tbsaldo_costo_kardex z where sc_codalm = '" & strAlm & "' and sc_periodo = DATE_FORMAT(sysdate(), '%Y%m') and sc_stock <> 0 " & _
'''''''''''                               ") XZ On P.idEmpresa  = xz.idempresa And P.idProducto = xz.sc_codart Where p.idEmpresa = '" & glsEmpresa & "' " & _
'''''''''''                               StrWhereMotivo & _
'''''''''''                               "AND (p.idProducto like '%" & Trim(TxtBusq.Text) & "%' OR p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') " & _
'''''''''''                               "AND estProducto = 'A' "
                               
                               'Luis 01/04/2019 "AND p.idProducto IN (SELECT preciosventa.idProducto FROM preciosventa WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & strCodLista & "')  " & _
                    '**********************************************************************
                    
                    
                        'If indValidaStock Then
                        '    setSqlAlm = setSqlAlm & "AND " & strCantidad & "  > 0 "
                        'End If
                    End If
                                   
                    G.Columns.ColumnByFieldName("Stock").Visible = True
                    setSqlAlm = setSqlAlm & "AND (p.GlsProducto "
                'End If
                
            Else
                setSqlAlm = "SELECT p.CodigoRapido,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto, Format(0,2) as Stock, t.GlsTallaPeso,P.IdFabricante " & _
                         "FROM productos p " & _
                                "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                                "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                                "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
                                "INNER JOIN productosalmacen a ON p.idEmpresa = a.idEmpresa AND p.idProducto = a.idProducto AND a.idAlmacen  = '" & strAlm & "' " & _
                                "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                         "WHERE p.idEmpresa = '" & glsEmpresa & "' AND (p.GlsProducto "
                                
            End If
        End If
    
    Else
'        If opt_Servicios.Value = True And CodDcto = "40" Or CodDcto = "92" Then
'            setSqlAlm = "SELECT p.idProducto,p.GlsProducto,'' as GlsMarca, p.idUMVenta, u.GlsUM, o.idMoneda as GlsMoneda,CASE WHEN p.afectoIGV = 1 THEN 'S' ELSE 'N' END Afecto, CAST(0 AS NUMERIC(12,2)) as Stock, '' AS GlsTallaPeso ,'' AS CodigoRapido,0 as Separacion,0 as Disponible,P.IdFabricante " & _
'                        "FROM productos p " & _
'                        "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
'                        "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
'                        "WHERE p.idEmpresa = '" & glsEmpresa & "' AND p.idTipoProducto = '06002' AND (p.GlsProducto "
'
'        Else
            setSqlAlm = "SELECT p.idProducto,p.GlsProducto,'' as GlsMarca, p.idUMVenta, u.GlsUM, o.idMoneda as GlsMoneda,CASE WHEN p.afectoIGV = 1 THEN 'S' ELSE 'N' END Afecto, CAST(0 AS NUMERIC(12,2)) as Stock, '' AS GlsTallaPeso ,'' AS CodigoRapido,P.IdFabricante " & _
                        "FROM productos p,monedas o, unidadMedida u " & _
                        "WHERE p.idMoneda  = o.idMoneda AND p.idEmpresa = '" & glsEmpresa & "' AND p.idTipoProducto = '06002' AND p.idUMVenta = u.idUM AND (p.GlsProducto "
        
        'End If
    End If
    
End Function

Public Sub ExecuteReturnTextAlm(ByVal strAlm As String, ByRef rspa As ADODB.Recordset, ByRef strCod As String, ByRef StrDes As String, ByRef strCodUM As String, ByVal ValidaStock As Boolean, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef TipoDcto As String, ByRef StrMsgError As String, Optional strParAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1, Optional PIdCliente As String, Optional PCodMotivo As String)
On Error GoTo Err
Dim IndAgregaAux                    As Boolean
    
    MousePointer = 0
    CodDcto = TipoDcto
    
    CIdCliente = PIdCliente
    CCodMotivo = PCodMotivo
    '--- Iniciamos Variables
    indAlmacen = True
    SRptBus(0) = ""
    indPedido = indVarPedido
    
    '--- Pasamos valores de parametros a las variables privadas a nivel de form
    indValidaStock = ValidaStock
    strCodLista = strVarCodLista
    strCodAlmacen = strAlm
    SqlAdic = strParAdic
    indUMVenta = indVarUMVenta
    indMostrarPresentaciones = indVarMostrarPresentaciones

    '--- Asignamos valores
    fraTipoProd.Visible = indMostrarTP
    
    If indMostrarPresentaciones = False Then
        Me.Height = fraPresentaciones.top + 350
        lblPresentaciones.Visible = False
    End If
    
    Select Case TipoProd
        Case 1 'productos
            opt_Producto.Value = True
        Case 2 'servicios
            opt_Servicios.Value = True
        Case 3 'materia prima
            opt_MateriaPrima.Value = True
    End Select
    
    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
'    '--- Filtramos
'    If Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "VISUALIZA_LISTA_PRODUCTOS", True) & "") = "S" Then
'        If opt_Producto.Value Or opt_MateriaPrima.Value Then
'
'            If leeParametro("AYUDA_PRODUCTOS_CLIENTE") = "S" Then
'
'                CargarProductosCliente StrMsgError
'
'            Else
'
'                fill StrMsgError
'
'            End If
'
'        Else
        
            fill StrMsgError
        
'        End If
'        If StrMsgError <> "" Then StrMsgError = Err.Description
'    End If
    
    Me.Show vbModal
    
    IndAgregaAux = IndAgrega
    
    '--- Devolvemos valores
    If SRptBus(0) <> "" Then
        If IndAgrega Then
            strCod = SRptBus(0)
            StrDes = SRptBus(1)
            strCodUM = SRptBus(2)
        Else
            strCod = ""
            StrDes = ""
            strCodUM = ""
        End If
    End If
    
    Set G.DataSource = Nothing
    Set gPresentaciones.DataSource = Nothing

    '--- Quitamos Filtros existentes
    G.Dataset.Filter = ""
    G.Dataset.Filtered = True
    
    gPresentaciones.Dataset.Filter = ""
    gPresentaciones.Dataset.Filtered = True
    
    Set G.DataSource = Nothing
    Set gPresentaciones.DataSource = Nothing
    
    If TypeName(rsg) = "Nothing" Then
        Exit Sub
    Else
        If rsg.State = 0 Then
            indEvaluaVacio = True
            Exit Sub
        End If
    End If
    
    '--- Eliminamos los registros q no estan marcados
    IndAgrega = IndAgregaAux
    
    rsg.Filter = ""
    rsg.MoveFirst
    If rsg.RecordCount > 0 Then
        rsg.MoveFirst
        Do While Not rsg.EOF
            If IndAgrega Then
                If rsg.Fields("CHK") = "0" Then
                    rsg.Delete adAffectCurrent
                    rsg.Update
                End If
            Else
                rsg.Delete adAffectCurrent
                rsg.Update
            End If
            rsg.MoveNext
        Loop
    End If
    
    Set rspa = rsg.Clone(adLockReadOnly)
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
        
    Unload Me
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub ExecuteKeyasciiReturnTextAlm(ByVal KeyAscii As Integer, strAlm As String, ByRef strCod As String, ByRef StrDes As String, ByRef strCodUM As String, ByVal ValidaStock As String, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef StrMsgError As String, Optional strParAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
On Error GoTo Err

    MousePointer = 0
    
    '--- Iniciamos Variables
    indAlmacen = True
    SRptBus(0) = ""
    indPedido = indVarPedido
    
    '--- Pasamos valores de parametros a las variables privadas a nivel de form
    strCodAlmacen = strAlm
    indValidaStock = ValidaStock
    strCodLista = strVarCodLista
    SqlAdic = strParAdic
    indUMVenta = indVarUMVenta
    indMostrarPresentaciones = indVarMostrarPresentaciones
    
    '--- Asignamos valores
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    fraTipoProd.Visible = indMostrarTP
    
    If indMostrarPresentaciones = False Then
        Me.Height = fraPresentaciones.top + 350
        lblPresentaciones.Visible = False
    End If
    
    Select Case TipoProd
        Case 1 'productos
            opt_Producto.Value = True
        Case 2 'servicios
            opt_Servicios.Value = True
        Case 3 'materia prima
            opt_MateriaPrima.Value = True
    End Select
    
    '--- Filtramos
    If opt_Producto.Value Or opt_MateriaPrima.Value Then
            
        If leeParametro("AYUDA_PRODUCTOS_CLIENTE") = "S" Then
    
            CargarProductosCliente StrMsgError
        
        Else
        
            fill StrMsgError
        
        End If
    
    Else
    
        fill StrMsgError
    
    End If
    If StrMsgError <> "" Then GoTo Err
    
    Me.Show vbModal
    
    '--- Devolvemos valores
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        StrDes = SRptBus(1)
        strCodUM = SRptBus(2)
    End If

    Set G.DataSource = Nothing
    Set gPresentaciones.DataSource = Nothing
    Unload Me
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub listaOtrasPresentaciones(ByRef StrMsgError As String)
Dim rsdatos                     As New ADODB.Recordset
On Error GoTo Err

    If indMostrarPresentaciones = False Then Exit Sub
    
    csql = "SELECT p.idUM,u.abreUM as GlsUM,CAST(r.factor AS NUMERIC(12,2)) AS factor,p.VVUnit AS VVUnit,p.IGVUnit AS IGVUnit,p.PVUnit AS PVUnit " & _
             "FROM preciosventa p,unidadMedida u, presentaciones r " & _
             "WHERE p.idUM = u.idUM " & _
               "AND p.idProducto = '" & G.Columns.ColumnByFieldName("idProducto").Value & "' " & _
               "AND p.idEmpresa = '" & glsEmpresa & "' " & _
               "AND p.idUM = r.idUM " & _
               "AND p.idProducto = r.idProducto " & _
               "AND r.idEmpresa = '" & glsEmpresa & "' " & _
               "AND p.idLista = '" & strCodLista & "' ORDER BY r.factor ASC"
               
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
        
    Set gPresentaciones.DataSource = rsdatos

'    With gPresentaciones
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idUM"
'    End With
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarNiveles(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsj As New ADODB.Recordset
Dim i As Integer

    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    fraNivel.Height = 355 * glsNumNiveles
    i = 0
    
    Do While Not rsj.EOF
        lblNivel(i).Caption = "" & rsj.Fields("GlsTipoNivel")
        rsj.MoveNext
        i = i + 1
    Loop
    
    TxtBusq.top = fraNivel.top + fraNivel.Height + 35
    lblBusq.top = TxtBusq.top
    
    fraContenido.Height = TxtBusq.top + TxtBusq.Height + 100
    If fraTipoProd.Height > fraContenido.Height Then
        fraContenido.Height = fraTipoProd.Height
    Else
        fraTipoProd.Height = fraContenido.Height
        fraGrilla.top = fraTipoProd.top + fraTipoProd.Height
        fraGrilla.Height = fraPresentaciones.top - (fraGrilla.top + lblPresentaciones.Height)
        G.Height = fraGrilla.Height - 200
    End If
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    
    Exit Sub
    
Err:
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Nivel_Change(Index As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If indMovNivel Then Exit Sub
    txtGls_Nivel(Index).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_Nivel(Index).Text, True)
    
    indMovNivel = True
    For i = Index + 1 To txtCod_Nivel.Count - 1
        txtCod_Nivel(i).Text = ""
        txtGls_Nivel(i).Text = ""
    Next
    
    indMovNivel = False
    If glsNumNiveles = Index + 1 Then
        If opt_Producto.Value Or opt_MateriaPrima.Value Then
            
            If leeParametro("AYUDA_PRODUCTOS_CLIENTE") = "S" Then
        
                CargarProductosCliente StrMsgError
            
            Else
            
                fill StrMsgError
            
            End If
        
        Else
        
            fill StrMsgError
        
        End If
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
