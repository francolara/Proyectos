VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmAyudaProdOCInv 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de Productos por Proveedor"
   ClientHeight    =   8865
   ClientLeft      =   3435
   ClientTop       =   1275
   ClientWidth     =   11985
   Icon            =   "frmAyudaProdOCInv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   36
      Top             =   8190
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Frame fraContenido 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   1860
      Left            =   90
      TabIndex        =   16
      Top             =   690
      Width           =   8745
      Begin VB.CommandButton BntOtrosProductos 
         Caption         =   "Más   Productos"
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
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1485
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Frame fraNivel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   60
         TabIndex        =   19
         Top             =   300
         Width           =   8625
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   0
            Left            =   8100
            Picture         =   "frmAyudaProdOCInv.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   0
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   1
            Left            =   8100
            Picture         =   "frmAyudaProdOCInv.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   360
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   2
            Left            =   8100
            Picture         =   "frmAyudaProdOCInv.frx":0720
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   720
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   3
            Left            =   8100
            Picture         =   "frmAyudaProdOCInv.frx":0AAA
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1080
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   4
            Left            =   8100
            Picture         =   "frmAyudaProdOCInv.frx":0E34
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1440
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   0
            Left            =   1305
            TabIndex        =   0
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
            Container       =   "frmAyudaProdOCInv.frx":11BE
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   25
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
            Container       =   "frmAyudaProdOCInv.frx":11DA
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   1
            Left            =   1305
            TabIndex        =   1
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
            Container       =   "frmAyudaProdOCInv.frx":11F6
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   26
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
            Container       =   "frmAyudaProdOCInv.frx":1212
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   2
            Left            =   1305
            TabIndex        =   2
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
            Container       =   "frmAyudaProdOCInv.frx":122E
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   2
            Left            =   2280
            TabIndex        =   27
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
            Container       =   "frmAyudaProdOCInv.frx":124A
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   3
            Left            =   1305
            TabIndex        =   3
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
            Container       =   "frmAyudaProdOCInv.frx":1266
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   3
            Left            =   2280
            TabIndex        =   28
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
            Container       =   "frmAyudaProdOCInv.frx":1282
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   4
            Left            =   1305
            TabIndex        =   4
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
            Container       =   "frmAyudaProdOCInv.frx":129E
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   4
            Left            =   2280
            TabIndex        =   29
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
            Container       =   "frmAyudaProdOCInv.frx":12BA
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
            TabIndex        =   34
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
            TabIndex        =   33
            Top             =   405
            Width           =   390
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   1485
            Width           =   345
         End
      End
      Begin CATControls.CATTextBox TxtBusq 
         Height          =   315
         Left            =   1365
         TabIndex        =   5
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
         Container       =   "frmAyudaProdOCInv.frx":12D6
         Vacio           =   -1  'True
      End
      Begin VB.Label lblBusq 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   195
         TabIndex        =   17
         Top             =   1200
         Width           =   645
      End
   End
   Begin VB.Frame fraPresentaciones 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   60
      TabIndex        =   15
      Top             =   6540
      Width           =   11865
      Begin DXDBGRIDLibCtl.dxDBGrid gPresentaciones 
         Height          =   1305
         Left            =   60
         OleObjectBlob   =   "frmAyudaProdOCInv.frx":12F2
         TabIndex        =   7
         Top             =   180
         Width           =   11715
      End
   End
   Begin VB.Frame fraTipoProd 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   1860
      Left            =   8880
      TabIndex        =   14
      Top             =   675
      Width           =   3045
      Begin VB.OptionButton opt_MateriaPrima 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   675
         TabIndex        =   10
         Top             =   1200
         Width           =   1290
      End
      Begin VB.OptionButton opt_Servicios 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   675
         TabIndex        =   9
         Top             =   810
         Width           =   1065
      End
      Begin VB.OptionButton opt_Producto 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   675
         TabIndex        =   8
         Top             =   420
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame fraGrilla 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   75
      TabIndex        =   13
      Top             =   2595
      Width           =   11865
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   3330
         Left            =   135
         OleObjectBlob   =   "frmAyudaProdOCInv.frx":3C62
         TabIndex        =   6
         Top             =   270
         Width           =   11685
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
            Picture         =   "frmAyudaProdOCInv.frx":82E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":867E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":8AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":8E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":9204
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":959E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":9938
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":9CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":A06C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":A406
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":A7A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOCInv.frx":B462
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1164
      ButtonWidth     =   3096
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "            Aceptar           "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Más Productos"
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
      Left            =   7245
      TabIndex        =   35
      Top             =   8520
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
      Left            =   60
      TabIndex        =   18
      Top             =   6360
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
      Left            =   10020
      TabIndex        =   12
      Top             =   6300
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
      Left            =   60
      TabIndex        =   11
      Top             =   8550
      Width           =   4680
   End
End
Attribute VB_Name = "FrmAyudaProdOCInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsg As New ADODB.Recordset
Private SqlAdic As String
Private sqlBus As String
Private sqlCond As String
Private SRptBus(3) As String
Private EsNuevo As Boolean
Private indAlmacen As Boolean
Private indValidaStock As Boolean
Private indPedido As Boolean
Private strCodAlmacen As String
Private cod_Prov As String
Private indUMVenta As Boolean
Private indMostrarPresentaciones As Boolean
Private strCodLista As String
Private indMovNivel As Boolean
Private intFoco As Integer '0 = Texto,1 = Grilla productos, 2 = Grilla presentaciones

Private Sub BntOtrosProductos_Click()
Dim strCod As String
Dim strDes As String

    Unload Me
    mostrarAyudaTexto "PRODUCTOS", strCod, strDes, " And IdProducto Not In(Select IdProducto From ProductosProveedor Where IdEmpresa = '" & glsEmpresa & "' And IdProveedor = '" & cod_Prov & "' Group By IdEmpresa,IdProducto)"
    SRptBus(0) = strCod
    SRptBus(1) = strDes
    SRptBus(2) = traerCampo("productos", "idUMCompra", "idProducto", strCod, True)

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

Private Sub cmbProdOtrasSucursales_Click()
On Error GoTo Err
Dim strCodProd As String, StrMsgError As String

    strCodProd = g.Columns.ColumnByFieldName("idProducto").Value
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
            g.SetFocus
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
Dim COptDefecto                 As String

    Me.Caption = "Ayuda de productos"
    Me.top = 0
    Me.left = 0
    
    ConfGrid g, True, False, False, False
    ConfGrid gPresentaciones, False, False, False, False
    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    EsNuevo = True
    If Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "VIZUALIZA_CODIGO_RAPIDO", True) & "") = "S" Then
        g.Columns.ColumnByFieldName("CodigoRapido").Visible = True
        g.Columns.ColumnByFieldName("IdProducto").Visible = False
    Else
        g.Columns.ColumnByFieldName("CodigoRapido").Visible = False
        g.Columns.ColumnByFieldName("IdProducto").Visible = True
    End If
    
    If leeParametro("VALIDA_PRODUCTOPROVEEDOR_OC") = "S" Then
        
        Toolbar1.Buttons(2).Visible = False
    
    End If
    
    COptDefecto = leeParametro("OPCION_AYUDA_PRODUCTOS")
    
    Select Case COptDefecto
        
        Case "P"
            opt_Producto.Value = True
        
        Case "S"
            opt_Servicios.Value = True
        
        Case "M"
            opt_MateriaPrima.Value = True
    
    End Select
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)

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

Private Sub g_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    
    If g.Dataset.State = dsEdit Then
        g.Dataset.Post
    End If

End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    Select Case KeyCode
        Case 13:
            SRptBus(0) = g.Columns.ColumnByFieldName("idProducto").Value
            SRptBus(1) = g.Columns.ColumnByFieldName("GlsProducto").Value '+ Space(150) + Trim(CStr(g.Columns.ColumnByFieldName("cod").Value))
            SRptBus(2) = g.Columns.ColumnByFieldName("idUMVenta").Value
            
            g.Dataset.Close
            g.Dataset.Active = False
            Me.Hide
    
        Case 27
            Unload Me
        Case 117
            Select Case intFoco
            Case 0
                g.SetFocus
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
            SRptBus(0) = g.Columns.ColumnByFieldName("idProducto").Value
            SRptBus(1) = g.Columns.ColumnByFieldName("GlsProducto").Value '+ Space(150) + Trim(CStr(g.Columns.ColumnByFieldName("cod").Value))
            SRptBus(2) = gPresentaciones.Columns.ColumnByFieldName("idUM").Value
            
            g.Dataset.Close
            g.Dataset.Active = False
            Me.Hide
        
        Case 27
            Unload Me
        
        Case 117
            Select Case intFoco
            Case 0
                g.SetFocus
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String
Dim strCod1 As String
Dim strDes1 As String
Dim CCodProducto            As String

    CCodProducto = IIf(leeParametro("VIZUALIZA_CODIGO_RAPIDO") = "S", "CodigoRapido", "IdProducto")
    Select Case Button.Index
        Case 1  'Aceptar
            If g.Count > 0 Then
                Me.Hide
            End If
            
        Case 2  'Otros Productos
            Unload Me
            mostrarAyudaTextoProducto "PRODUCTOS", strCod1, strDes1, " and idProducto not in(select idProducto from productosproveedor where idEmpresa = '" & glsEmpresa & "' and idProveedor = '" & cod_Prov & "' Group By IdEmpresa,IdProducto) "
             
            SRptBus(0) = strCod1
            SRptBus(1) = strDes1
            SRptBus(2) = traerCampo("productos", "idUMCompra", CCodProducto, strCod1, True)
        
        Case 3  'Salir
            'Unload Me
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
        fill StrMsgError
        If StrMsgError <> "" Then GoTo Err
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
    
    If KeyCode = vbKeyDown Then g.SetFocus
    If EsNuevo = True Then TxtBusq.SelStart = Len(TxtBusq.Text) + 1

End Sub

Private Sub TxtBusq_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyAscii = 13 Then
        fill StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        If g.Count > 1 Then g.SetFocus
    End If
    EsNuevo = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub fill(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsdatos As New ADODB.Recordset

    sqlBus = setSqlAlm(strCodAlmacen)
    sqlCond = sqlBus & " like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') "
    
    If txtCod_Nivel(glsNumNiveles - 1).Text <> "" Then
        sqlCond = sqlCond & " AND idNivel = '" & txtCod_Nivel(glsNumNiveles - 1).Text & "'"
    End If
    
    If opt_Producto.Value Then
        sqlCond = sqlCond & " AND idTipoProducto = '06001'"
    ElseIf opt_MateriaPrima.Value Then
        sqlCond = sqlCond & " AND idTipoProducto = '06003'"
    End If
    
    sqlCond = sqlCond & SqlAdic & " order by 1"
    
    If rsdatos.State = 1 Then rsdatos.Close
    rsdatos.Open sqlCond, Cn, adOpenStatic, adLockReadOnly
    If rsg.State = 1 Then rsg.Close
    
    rsg.Fields.Append "CHK", adChar, 1, adFldIsNullable
    rsg.Fields.Append "idProducto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 40, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "GlsMarca", adVarChar, 80, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 80, adFldIsNullable
    rsg.Fields.Append "GlsMoneda", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsTallaPeso", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "Afecto", adVarChar, 5, adFldIsNullable
    rsg.Fields.Append "idUMVenta", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "Stock", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idFabricante", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GLSDOCREFERENCIA", adVarChar, 200, adFldIsNullable
    rsg.Open
    
    If rsdatos.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("CHK") = 0
        rsg.Fields("idProducto") = ""
        rsg.Fields("CodigoRapido") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("GlsMarca") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("GlsMoneda") = ""
        rsg.Fields("GlsTallaPeso") = ""
        rsg.Fields("Afecto") = ""
        rsg.Fields("idUMVenta") = ""
        rsg.Fields("Stock") = 0#
        rsg.Fields("idFabricante") = ""
        rsg.Fields("GLSDOCREFERENCIA") = ""
    
    Else
        Do While Not rsdatos.EOF
            item = (rsg.RecordCount) + 1
            rsg.AddNew
            rsg.Fields("CHK") = 0
            rsg.Fields("idProducto") = "" & rsdatos.Fields("idProducto")
            rsg.Fields("CodigoRapido") = "" & rsdatos.Fields("CodigoRapido")
            rsg.Fields("GlsProducto") = "" & rsdatos.Fields("GlsProducto")
            rsg.Fields("GlsMarca") = "" & rsdatos.Fields("GlsMarca")
            rsg.Fields("GlsUM") = "" & rsdatos.Fields("GlsUM")
            rsg.Fields("GlsMoneda") = "" & rsdatos.Fields("GlsMoneda")
            rsg.Fields("GlsTallaPeso") = "" & rsdatos.Fields("GlsTallaPeso")
            rsg.Fields("Afecto") = rsdatos.Fields("Afecto")
            rsg.Fields("Stock") = rsdatos.Fields("Stock")
            rsdatos.MoveNext
       Loop
    End If
    
    Set g.DataSource = Nothing
    With g
        .DefaultFields = False
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        Set .DataSource = rsg
        .Dataset.Active = True
        .KeyField = "idProducto"
        .Dataset.Edit
        .Dataset.Post
    End With
    Set rsdatos = Nothing
    
    LblReg.Caption = "(" + Format(g.Count, "0") + ")Registros"
    listaOtrasPresentaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Function setSqlAlm(strAlm As String) As String
Dim strCampoUM As String
Dim strStockUM As String
Dim strCantidad As String
Dim strTablaPresentaciones As String

    If glsVisualizaCodFab = "N" Then
        g.Columns.ColumnByFieldName("IdFabricante").Visible = False
    End If

    strCampoUM = "idUMVenta"
    strStockUM = "CantidadStockUV"
    strCantidad = "(a.CantidadStock / f.Factor )" 'Es la cantidad de venta
    
    strTablaPresentaciones = " INNER JOIN presentaciones f ON p.idEmpresa = f.idEmpresa AND p.idProducto = f.idProducto AND p." & strCampoUM & " = f.idUM "
        
    If indUMVenta = False Then
        strCampoUM = "idUMCompra"
        strStockUM = "CantidadStockUC"
        strCantidad = "a.CantidadStock" 'Es la cantidad de compra
        strTablaPresentaciones = ""
    End If
    
    g.Columns.ColumnByFieldName("Stock").Visible = False
    If opt_Servicios.Value = False Then
        If indPedido = False Then
            setSqlAlm = "SELECT '0' AS CHK, p.idProducto,P.CodigoRapido,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda,iif(p.afectoIGV = 1,'S','N') Afecto, CAST(" & strCantidad & " AS DECIMAL(12,2)) as Stock, t.GlsTallaPeso, p.idfabricante " & _
                        "FROM productos p " & _
                        "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                        "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                        "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
                        "INNER JOIN productosalmacen a ON p.idEmpresa = a.idEmpresa AND a.idSucursal = '" & glsSucursal & "' " & _
                                                      "AND a.idAlmacen = '" & strAlm & "' AND p.idProducto = a.idProducto " & _
                                                      "AND p." & strCampoUM & " = a.idUMCompra " & strTablaPresentaciones & _
                        "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                        "WHERE p.idEmpresa = '" & glsEmpresa & "' "
                       
            If indValidaStock Then
                setSqlAlm = setSqlAlm & "AND " & strCantidad & "  > 0 "
            End If
            g.Columns.ColumnByFieldName("Stock").Visible = True
            setSqlAlm = setSqlAlm & "AND (p.GlsProducto "
                       
        Else
        
            If leeParametro("VALIDA_PRODUCTOPROVEEDOR_OC") = "S" Then
            
                setSqlAlm = "SELECT '0' AS CHK, p.idProducto,P.CodigoRapido,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda,iif(p.afectoIGV = 1,'S','N') Afecto, 0.00 as Stock, t.GlsTallaPeso " & _
                            "FROM productos p " & _
                            "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                            "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                            "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
                            "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                            "WHERE p.idEmpresa = '" & glsEmpresa & "' AND (p.GlsProducto "
                            
            Else
            
                setSqlAlm = "SELECT '0' AS CHK, p.idProducto,P.CodigoRapido,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda,iif(p.afectoIGV = 1,'S','N') Afecto, 0.00 as Stock, t.GlsTallaPeso " & _
                            "FROM productos p " & _
                            "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                            "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                            "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
                            "INNER JOIN (" & _
                                "Select A.IdEmpresa,A.IdProducto " & _
                                "From ProductosProveedor A " & _
                                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProveedor = '" & cod_Prov & "' " & _
                                "Group By A.IdEmpresa,A.IdProducto " & _
                            ") x " & _
                                "ON p.idEmpresa = x.idEmpresa and x.idProducto = p.idProducto " & _
                            "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                            "WHERE p.idEmpresa = '" & glsEmpresa & "' AND (p.GlsProducto "
                            
            End If
            
        End If
        
    Else
        setSqlAlm = "SELECT '0' AS CHK, p.idProducto,P.CodigoRapido,p.GlsProducto,'' as GlsMarca,'' as GlsUM,o.GlsMoneda,iif(p.afectoIGV = 1,'S','N') Afecto, 0.00 as Stock, '' AS GlsTallaPeso " & _
                    "FROM productos p,monedas o " & _
                    "WHERE p.idMoneda  = o.idMoneda AND p.idEmpresa = '" & glsEmpresa & "' AND p.idTipoProducto = '06002' AND (p.GlsProducto "
    End If
    
End Function

Public Sub ExecuteReturnTextAlm(ByVal codProv As String, ByVal strAlm As String, ByRef rsp As ADODB.Recordset, ByRef strCod As String, ByRef strDes As String, ByRef strCodUM As String, ByVal ValidaStock As Boolean, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef StrMsgError As String, Optional strParAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
On Error GoTo Err
    
    MousePointer = 0
    'Iniciamos Variables
    indAlmacen = True
    SRptBus(0) = ""
    indPedido = indVarPedido
    
    'Pasamos valores de parametros a las variables privadas a nivel de form
    indValidaStock = ValidaStock
    strCodLista = strVarCodLista
    strCodAlmacen = strAlm
    cod_Prov = codProv
    SqlAdic = strParAdic
    indUMVenta = indVarUMVenta
    indMostrarPresentaciones = indVarMostrarPresentaciones

    'Asignamos valores
    fraTipoProd.Visible = indMostrarTP
    
    If indMostrarPresentaciones = False Then
        Me.Height = fraPresentaciones.top + 350
        lblPresentaciones.Visible = False
    End If
    
'    Select Case TipoProd
'    Case 1 'productos
'        opt_Producto.Value = True
'    Case 2 'servicios
'        opt_Servicios.Value = True
'    Case 3 'materia prima
'        opt_MateriaPrima.Value = True
'    End Select
    
    'Filtramos
    fill StrMsgError
    If StrMsgError <> "" Then StrMsgError = Err.Description
    
    Me.Show vbModal
    
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
        strCodUM = SRptBus(2)
        
        Set g.DataSource = Nothing
        Set gPresentaciones.DataSource = Nothing
        If rsg.State = 1 Then rsg.Close
    
    Else
        'Quitamos Filtros existentes
        g.Dataset.Filter = ""
        g.Dataset.Filtered = True
        
        gPresentaciones.Dataset.Filter = ""
        gPresentaciones.Dataset.Filtered = True
        
        Set g.DataSource = Nothing
        Set gPresentaciones.DataSource = Nothing
        If TypeName(rsg) = "Nothing" Then
            Exit Sub
        Else
            If rsg.State = 0 Then
                Exit Sub
            End If
        End If
        
        'Eliminamos los registros q no estan marcados
        If rsg.RecordCount > 0 Then
            rsg.MoveFirst
            Do While Not rsg.EOF
                If rsg.Fields("CHK") = "0" Then
                    rsg.Delete adAffectCurrent
                    rsg.Update
                End If
                rsg.MoveNext
            Loop
        End If
    
        Set rsp = rsg.Clone(adLockReadOnly)
        If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    End If
    Unload Me
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub ExecuteKeyasciiReturnTextAlm(ByVal KeyAscii As Integer, strAlm As String, ByRef strCod As String, ByRef strDes As String, ByRef strCodUM As String, ByVal ValidaStock As String, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef StrMsgError As String, Optional strParAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
On Error GoTo Err

    MousePointer = 0
    'Iniciamos Variables
    indAlmacen = True
    SRptBus(0) = ""
    indPedido = indVarPedido
    
    'Pasamos valores de parametros a las variables privadas a nivel de form
    strCodAlmacen = strAlm
    indValidaStock = ValidaStock
    strCodLista = strVarCodLista
    SqlAdic = strParAdic
    indUMVenta = indVarUMVenta
    indMostrarPresentaciones = indVarMostrarPresentaciones
    
    'Asignamos valores
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
    
    'Filtramos
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Show vbModal

    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
        strCodUM = SRptBus(2)
        Set g.DataSource = Nothing
        Set gPresentaciones.DataSource = Nothing
        If rsg.State = 1 Then rsg.Close
    Else
        'Quitamos Filtros existentes
        g.Dataset.Filter = ""
        g.Dataset.Filtered = True
        
        gPresentaciones.Dataset.Filter = ""
        gPresentaciones.Dataset.Filtered = True
        
        Set g.DataSource = Nothing
        Set gPresentaciones.DataSource = Nothing
        If TypeName(rsg) = "Nothing" Then
            Exit Sub
        Else
            If rsg.State = 0 Then
                Exit Sub
            End If
        End If
        
        'Eliminamos los registros q no estan marcados
        If rsg.RecordCount > 0 Then
            rsg.MoveFirst
            Do While Not rsg.EOF
                If rsg.Fields("ok") = "0" Then
                    rsg.Delete adAffectCurrent
                    rsg.Update
                End If
                rsg.MoveNext
            Loop
        End If
        Set rsp = rsg.Clone(adLockReadOnly)
        If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    End If
    Unload Me

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub listaOtrasPresentaciones(ByRef StrMsgError As String)
Dim rsdatos                     As New ADODB.Recordset
On Error GoTo Err

    If indMostrarPresentaciones = False Then Exit Sub
    
    csql = "SELECT p.idUM,u.abreUM as GlsUM,CAST(r.factor AS DECIMAL(4,2)) AS factor,pp.ValorVenta AS VVUnit,pp.IGVVenta AS IGVUnit,pp.PrecioVenta AS PVUnit " & _
             "FROM preciosventa p " & _
             "Inner Join unidadMedida u " & _
                "On p.idUM = u.idUM " & _
            "Inner Join presentaciones r " & _
                "On p.idUM = r.idUM AND p.idProducto = r.idProducto AND p.idEmpresa = r.idEmpresa " & _
            "INNER JOIN (" & _
                "Select A.IdEmpresa,A.IdProducto,A.ValorVenta,A.IGVVenta,A.PrecioVenta " & _
                "From ProductosProveedor A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProveedor = '" & cod_Prov & "' " & _
                "Group By A.IdEmpresa,A.IdProducto,A.ValorVenta,A.IGVVenta,A.PrecioVenta " & _
            ") pp " & _
                "ON p.idEmpresa = pp.idEmpresa and p.idProducto = pp.idProducto " & _
             "WHERE p.idProducto = '" & g.Columns.ColumnByFieldName("idProducto").Value & "' " & _
               "AND p.idEmpresa = '" & glsEmpresa & "' " & _
               "AND p.idEmpresa = '" & glsEmpresa & "' " & _
               "AND p.idLista = '" & strCodLista & "' ORDER BY r.factor ASC"
               
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
'
If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
    
Set gPresentaciones.DataSource = rsdatos

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
        g.Height = fraGrilla.Height - 200
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
        fill StrMsgError
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
