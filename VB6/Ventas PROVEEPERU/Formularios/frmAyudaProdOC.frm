VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmAyudaProdOC 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Productos por Proveedor"
   ClientHeight    =   9360
   ClientLeft      =   10500
   ClientTop       =   3135
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAyudaProdOC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   12075
   Begin VB.CommandButton cmbProdOtrasSucursales 
      Caption         =   "Consultar en otras sucursales"
      Height          =   405
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8670
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Frame fraContenido 
      Appearance      =   0  'Flat
      Caption         =   " Filtros "
      ForeColor       =   &H00000000&
      Height          =   2220
      Left            =   90
      TabIndex        =   8
      Top             =   690
      Width           =   8745
      Begin VB.CommandButton BntOtrosProductos 
         Caption         =   "Otros Productos"
         Height          =   315
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1845
         Visible         =   0   'False
         Width           =   2475
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
         Height          =   360
         Left            =   60
         TabIndex        =   11
         Top             =   180
         Width           =   8625
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
            Left            =   8100
            Picture         =   "frmAyudaProdOC.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   30
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
            Left            =   8100
            Picture         =   "frmAyudaProdOC.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   390
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
            Left            =   8100
            Picture         =   "frmAyudaProdOC.frx":0720
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   750
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
            Left            =   8100
            Picture         =   "frmAyudaProdOC.frx":0AAA
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1110
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
            Index           =   4
            Left            =   8100
            Picture         =   "frmAyudaProdOC.frx":0E34
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1470
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   0
            Left            =   1305
            TabIndex        =   17
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
            Container       =   "frmAyudaProdOC.frx":11BE
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   18
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
            Container       =   "frmAyudaProdOC.frx":11DA
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   1
            Left            =   1305
            TabIndex        =   19
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
            Container       =   "frmAyudaProdOC.frx":11F6
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   20
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
            Container       =   "frmAyudaProdOC.frx":1212
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   2
            Left            =   1305
            TabIndex        =   21
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
            Container       =   "frmAyudaProdOC.frx":122E
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   2
            Left            =   2280
            TabIndex        =   22
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
            Container       =   "frmAyudaProdOC.frx":124A
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   3
            Left            =   1305
            TabIndex        =   23
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
            Container       =   "frmAyudaProdOC.frx":1266
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   3
            Left            =   2280
            TabIndex        =   24
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
            Container       =   "frmAyudaProdOC.frx":1282
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   4
            Left            =   1305
            TabIndex        =   25
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
            Container       =   "frmAyudaProdOC.frx":129E
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   4
            Left            =   2280
            TabIndex        =   26
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
            Container       =   "frmAyudaProdOC.frx":12BA
            Vacio           =   -1  'True
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   135
            TabIndex        =   31
            Top             =   45
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
            TabIndex        =   30
            Top             =   405
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
            TabIndex        =   29
            Top             =   765
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
            TabIndex        =   28
            Top             =   1125
            Width           =   345
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   135
            TabIndex        =   27
            Top             =   1485
            Width           =   345
         End
      End
      Begin CATControls.CATTextBox TxtBusq 
         Height          =   315
         Left            =   1365
         TabIndex        =   37
         Top             =   1500
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
         Container       =   "frmAyudaProdOC.frx":12D6
         Vacio           =   -1  'True
      End
      Begin VB.Label lblBusq 
         Appearance      =   0  'Flat
         Caption         =   "Producto"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   195
         TabIndex        =   9
         Top             =   1560
         Width           =   795
      End
   End
   Begin VB.Frame fraPresentaciones 
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
      Height          =   1605
      Left            =   60
      TabIndex        =   7
      Top             =   7020
      Width           =   11865
      Begin DXDBGRIDLibCtl.dxDBGrid gPresentaciones 
         Height          =   1305
         Left            =   60
         OleObjectBlob   =   "frmAyudaProdOC.frx":12F2
         TabIndex        =   32
         Top             =   180
         Width           =   11715
      End
   End
   Begin VB.Frame fraTipoProd 
      Appearance      =   0  'Flat
      Caption         =   " Tipo "
      ForeColor       =   &H00000000&
      Height          =   2220
      Left            =   8880
      TabIndex        =   3
      Top             =   690
      Width           =   3045
      Begin VB.OptionButton OptFormulas 
         Caption         =   "Fórmulas"
         Height          =   240
         Left            =   540
         TabIndex        =   38
         Top             =   1440
         Width           =   1290
      End
      Begin VB.OptionButton opt_MateriaPrima 
         Caption         =   "Materia Prima"
         Height          =   240
         Left            =   540
         TabIndex        =   6
         Top             =   885
         Width           =   1290
      End
      Begin VB.OptionButton opt_Servicios 
         Caption         =   "Servicios"
         Height          =   240
         Left            =   2115
         TabIndex        =   5
         Top             =   2070
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.OptionButton opt_Producto 
         Caption         =   "Productos"
         Height          =   240
         Left            =   540
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame fraGrilla 
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
      Height          =   3690
      Left            =   75
      TabIndex        =   2
      Top             =   2955
      Width           =   11865
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   3330
         Left            =   135
         OleObjectBlob   =   "frmAyudaProdOC.frx":3C62
         TabIndex        =   35
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
            Picture         =   "frmAyudaProdOC.frx":82C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":865E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":8AB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":8E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":91E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":957E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":9918
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":9CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":A04C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":A3E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":A780
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaProdOC.frx":B442
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
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   1164
      ButtonWidth     =   2858
      ButtonHeight    =   1005
      Appearance      =   1
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
   Begin VB.Label lblPresentaciones 
      Appearance      =   0  'Flat
      Caption         =   "Otras Presentaciones:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   60
      TabIndex        =   10
      Top             =   6720
      Width           =   3435
   End
   Begin VB.Label LblReg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "(0) Registros"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10020
      TabIndex        =   1
      Top             =   6660
      Width           =   1905
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Presionar Enter en el registro para obtener el resultado "
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7215
      TabIndex        =   0
      Top             =   8670
      Width           =   4680
   End
End
Attribute VB_Name = "FrmAyudaProdOC"
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
Private intFoco As Integer
Private valeind As Boolean
Private sw_todas_sucursales As Boolean
Private IndAgrega               As Boolean
Dim NIdStock                            As String

Private Sub CmdBusq_Click()
Dim StrMsgError As String
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

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
    
    IndAgrega = True
    Me.Caption = "Ayuda de productos"
    ConfGrid G, True, False, False, False
    ConfGrid gPresentaciones, False, False, False, False
    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    EsNuevo = True
    
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
On Error GoTo Err
Dim StrMsgError                     As String
Dim rsEli                           As New ADODB.Recordset

    IndAgrega = False
    SqlAdic = ""
    TxtBusq.Text = ""
    
    NIdStock = IIf(NIdStock = "", "0a", NIdStock)
    
    '************** LUIS LARA 13/01/2018 *****************
    'Set rsEli = DataProcedimiento("Spu_EliminaTemporales", strMsgError, NIdStock, "0a", "0a", "0a", "0a", "0a")
    'If strMsgError <> "" Then GoTo ERR
    '*****************************************************
    
    NIdStock = ""
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
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
    
    If G.Dataset.State = dsEdit Then
        G.Dataset.Post
    End If

End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    
    Select Case KeyCode
        Case 13:
            SRptBus(0) = G.Columns.ColumnByFieldName("idProducto").Value
            SRptBus(1) = G.Columns.ColumnByFieldName("GlsProducto").Value
            SRptBus(2) = G.Columns.ColumnByFieldName("idUMVenta").Value
            
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
            SRptBus(1) = G.Columns.ColumnByFieldName("GlsProducto").Value
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

    Select Case Button.Index
        Case 1  'Aceptar
            If G.Count > 0 Then
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

'    If EsNuevo = False Then
'        fill StrMsgError
'        If StrMsgError <> "" Then GoTo Err
'    End If
    
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
        If G.Count > 1 Then G.SetFocus
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
    sqlCond = sqlBus & " like '%" & Trim(TxtBusq.Text) & "%' OR p.IdProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') "
    
    If txtCod_Nivel(glsNumNiveles - 1).Text <> "" Then
        sqlCond = sqlCond & " AND idNivel = '" & txtCod_Nivel(glsNumNiveles - 1).Text & "'"
    End If
    
    If opt_Producto.Value Then
        sqlCond = sqlCond & " AND idTipoProducto = '06001'"
    ElseIf opt_MateriaPrima.Value Then
        sqlCond = sqlCond & " AND idTipoProducto = '06003'"
    ElseIf OptFormulas.Value Then
        sqlCond = sqlCond & " AND idTipoProducto = '06004'"
    End If
    
    sqlCond = sqlCond & SqlAdic & " order by 1"
    
    If rsdatos.State = 1 Then rsdatos.Close
    rsdatos.Open sqlCond, Cn, adOpenStatic, adLockReadOnly
    If rsg.State = 1 Then rsg.Close
    
    rsg.Fields.Append "CHK", adChar, 1, adFldIsNullable
    rsg.Fields.Append "idProducto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "GlsMarca", adVarChar, 80, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 80, adFldIsNullable
    rsg.Fields.Append "GlsMoneda", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsTallaPeso", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "Afecto", adVarChar, 5, adFldIsNullable
    rsg.Fields.Append "idUMVenta", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "Stock", adDouble, 11, adFldIsNullable
    rsg.Fields.Append "idFabricante", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GLSDOCREFERENCIA", adVarChar, 200, adFldIsNullable
    rsg.Fields.Append "AfectoIgv", adInteger, 4, adFldIsNullable
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
        rsg.Fields("AfectoIgv") = 0
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
            rsg.Fields("idFabricante") = rsdatos.Fields("idFabricante")
            rsg.Fields("AfectoIgv") = Val("" & rsdatos.Fields("AfectoIGV"))
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
    Set rsdatos = Nothing
    
    LblReg.Caption = "(" + Format(G.Count, "0") + ")Registros"
    
    listaOtrasPresentaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
   ' Resume
    Exit Sub
    Resume
End Sub

Public Sub Execute(ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
End Sub

Private Function setSqlAlm(strAlm As String) As String
Dim strCampoUM As String
Dim strStockUM As String
Dim strCantidad As String
Dim strTablaPresentaciones As String, csucursal As String
Dim strCadStk   As String
Dim nPC                             As String
    
    strCadStk = ""

    If glsVisualizaCodFab = "N" Then
        G.Columns.ColumnByFieldName("IdFabricante").Visible = False
    End If
    strCampoUM = "idUMVenta"
    strStockUM = "CantidadStockUV"
    'strCantidad = "(a.CantidadStock)"
    'strCantidad = "(stock)"
    
    strCantidad = IIf(indValidaStock = False, "(a.CantidadStock)", "(stock)")
    
    If indUMVenta = False Then
        strCampoUM = "idUMCompra"
        strStockUM = "CantidadStockUC"
        'strCantidad = "a.CantidadStock" 'Es la cantidad de compra
        'strCantidad = "(stock)"
        strCantidad = IIf(indValidaStock = False, "(a.CantidadStock)", "(stock)")
        strTablaPresentaciones = ""
    End If
    
    G.Columns.ColumnByFieldName("Stock").Visible = False
     
    If opt_Servicios.Value = False Then
        If valeind = True Then
            setSqlAlm = "SELECT '0' AS CHK, p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda," & _
                        "CASE WHEN p.afectoIGV = 1 THEN 'S' ELSE 'N' END Afecto, CAST(0 AS NUMERIC(12,2)) as Stock, '' AS GlsTallaPeso,p.idfabricante,p.AfectoIGV " & _
                        "FROM productos p " & _
                        "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                        "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                        "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
                        "WHERE p.idMoneda  = o.idMoneda AND p.idEmpresa = '" & glsEmpresa & "' And P.EstProducto = 'A' AND (p.GlsProducto "
        
        Else
            If indPedido = False Then
                If indValidaStock = False Then
                
                    If opt_Producto.Value Then
                        strTipoProducto = "06001"
                    ElseIf opt_MateriaPrima.Value Then
                        strTipoProducto = "06003"
                    ElseIf OptFormulas.Value Then
                        strTipoProducto = "06004"
                    End If
                    
                    If sw_todas_sucursales Then
                        csucursal = ""
                    Else
                        csucursal = " AND a.idSucursal = '" & glsSucursal & "' "
                    End If
    
                    setSqlAlm = "SELECT '0' AS CHK, p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM, " & _
                                "o.GlsMoneda,CASE WHEN p.afectoIGV = 1 THEN 'S' ELSE 'N' END Afecto, CAST(" & strCantidad & " AS NUMERIC(12,2)) as Stock, t.GlsTallaPeso, " & _
                                "p.idfabricante,p.AfectoIGV " & _
                                "FROM productos p " & _
                                "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                                "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                                "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
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
                           "Group bY P.idEmpresa,vd.idSucursal,P.idProducto " & _
                           ") A " & _
                           "On P.idEmpresa = A.idEmpresa And P.idProducto = A.idProducto " & _
                                "WHERE  " & _
                                " p.idEmpresa = '" & glsEmpresa & "' And P.EstProducto = 'A' "
    
                    If indValidaStock Then
                        setSqlAlm = setSqlAlm & "AND " & strCantidad & "  > 0 "
                    End If
                    G.Columns.ColumnByFieldName("Stock").Visible = True
                    setSqlAlm = setSqlAlm & "AND (p.GlsProducto "
                Else
                    If opt_Producto.Value Then
                        strTipoProducto = "06001"
                    ElseIf opt_MateriaPrima.Value Then
                        strTipoProducto = "06003"
                    ElseIf OptFormulas.Value Then
                        strTipoProducto = "06004"
                    End If
                    
                    If indValidaStock Then
                       strCadStk = "Having (stock) > 0"
                    End If
                    
                    If NIdStock = "" Then
                        
                        nPC = ComputerName
                        nPC = Replace(nPC, "-", "")
                        nPC = Replace(nPC, "", Trim(nPC))
                        nPC = Trim(nPC)
        
                        NIdStock = Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_TS")
                        
                    End If
                    
                    'setSqlAlm = "SELECT '0' AS CHK,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda," & _
                                "if(p.afectoIGV = 1,'S','N') Afecto,Format(s.Stock,2) as Stock,t.GlsTallaPeso , p.idfabricante,p.AfectoIGV " & _
                               "FROM productos p " & _
                               "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                               "INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM " & _
                               "INNER JOIN monedas o ON p.idMoneda = o.idMoneda " & _
                               "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                               "Left Join " & _
                               "( " & _
                               "Select P.idEmpresa,IfNull(vd.idSucursal,'') Idsucursal,P.idProducto,vc.idAlmacen, " & _
                               "sum(If(vd.idempresa is null,0,If(vc.idSucursal = '" & glsSucursal & "' And vc.estValeCab <> 'ANU' And idAlmacen = '" & strAlm & "'  ,if(vd.tipovale = 'I',Cantidad,Cantidad * -1),0))) as Stock " & _
                               "From Productos P " & _
                               "Inner Join ValesDet vd On P.IdEmpresa = vd.IdEmpresa And P.IdProducto = vd.IdProducto " & _
                               "Inner Join Valescab vc On vd.idEmpresa = vc.idEmpresa And vd.idSucursal = vc.idSucursal And vd.tipoVale = vc.tipoVale And vd.idValesCab = vc.idValesCab " & _
                               "Inner Join PeriodosINV pi on vc.idEmpresa = pi.idEmpresa And vc.idPeriodoINV = pi.idPeriodoINV And vc.idSucursal = pi.idSucursal " & _
                               "Where P.idEmpresa = '" & glsEmpresa & "' AND idTipoProducto = '" & strTipoProducto & "'  AND estProducto = 'A' AND (p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%')And (pi.estPeriodoInv = 'ACT' Or pi.estPeriodoInv Is Null) " & _
                               "Group bY P.idEmpresa,P.idProducto " & strCadStk & " order by P.idEmpresa,P.idProducto " & _
                               ") S " & _
                               "On P.idEmpresa = S.idEmpresa And P.idProducto = S.idProducto " & _
                               "Where p.idEmpresa = '" & glsEmpresa & "' " & _
                               "AND (p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') " & _
                               "AND estProducto = 'A' "
                               
                    '********************** LUIS LARA 12/01/2018 **************************
                    setSqlAlm = "SELECT '0' AS CHK,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda," & _
                                "if(p.afectoIGV = 1,'S','N') Afecto,Format(ifnull(XZ.sc_stock,0) + ifnull(s.Stock,0),2) as Stock,t.GlsTallaPeso , p.idfabricante,p.AfectoIGV " & _
                               "FROM productos p " & _
                               "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                               "INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM " & _
                               "INNER JOIN monedas o ON p.idMoneda = o.idMoneda " & _
                               "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                               "Left Join ( " & _
                               "Select P.idEmpresa,IfNull(vd.idSucursal,'') Idsucursal,P.idProducto,vc.idAlmacen, " & _
                               "sum(If(vd.idempresa is null,0,if(vd.tipovale = 'I',Cantidad,Cantidad * -1))) as Stock " & _
                               "From (SELECT * FROM Valescab vc WHERE vc.idAlmacen = '" & strAlm & "' AND DATE_FORMAT(vc.fechaemision, '%Y%m%d')  = DATE_FORMAT(sysdate(), '%Y%m%d')   ) vc " & _
                               "Inner Join ValesDet vd On vd.idEmpresa = vc.idEmpresa And vd.idSucursal = vc.idSucursal And vd.tipoVale = vc.tipoVale And vd.idValesCab = vc.idValesCab " & _
                               "inner join Productos P On P.IdEmpresa = vd.IdEmpresa And P.IdProducto = vd.IdProducto " & _
                               "Where P.idEmpresa = '" & glsEmpresa & "' AND idTipoProducto = '" & strTipoProducto & "'  AND estProducto = 'A' " & _
                               "AND vc.idSucursal = '" & glsSucursal & "' And vc.estValeCab <> 'ANU' " & _
                               "AND (p.idProducto like '%" & Trim(TxtBusq.Text) & "%' OR p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') " & _
                               "Group bY P.idEmpresa,P.idProducto,vc.idAlmacen order by P.idEmpresa,P.idProducto " & _
                               ") S " & _
                               "On P.idEmpresa = S.idEmpresa And P.idProducto = S.idProducto " & _
                               "Left Join (SELECT sc_periodo,sc_codalm,sc_codart,sc_stock,idempresa FROM tbsaldo_costo_kardex z where sc_codalm = '" & strAlm & "' and sc_periodo = DATE_FORMAT(sysdate(), '%Y%m') and sc_stock <> 0 " & _
                               ") XZ On P.idEmpresa  = xz.idempresa And P.idProducto = xz.sc_codart " & _
                               "Where p.idEmpresa = '" & glsEmpresa & "' " & _
                               "AND (p.idProducto like '%" & Trim(TxtBusq.Text) & "%' OR p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') " & _
                               "AND estProducto = 'A' "
                    '**********************************************************************
                               
                    '********************** LUIS LARA 12/01/2018 **************************
''''                    Cn.Execute "Call Spu_CalculaStock('" & glsEmpresa & "','" & NIdStock & "','','0','" & glsSucursal & "','','','" & strAlm & "',SysDate(),'')"
''''
''''                    setSqlAlm = "SELECT '0' AS CHK,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda," & _
''''                                "if(p.afectoIGV = 1,'S','N') Afecto,Format(s.Stock,2) as Stock,t.GlsTallaPeso , p.idfabricante,p.AfectoIGV " & _
''''                               "FROM productos p " & _
''''                               "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
''''                               "INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM " & _
''''                               "INNER JOIN monedas o ON p.idMoneda = o.idMoneda " & _
''''                               "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
''''                               "Left Join " & NIdStock & " S " & _
''''                               "On P.idEmpresa = S.idEmpresa And P.idProducto = S.idProducto " & _
''''                               "Where p.idEmpresa = '" & glsEmpresa & "' " & _
''''                               "AND estProducto = 'A' "
                    '************************************************************************
                    
                    
                               '"AND (p.GlsProducto like '%" & Trim(TxtBusq.Text) & "%' OR CodigoRapido like '%" & Trim(TxtBusq.Text) & "%' OR IdFabricante like '%" & Trim(TxtBusq.Text) & "%') "
                               
                                If indValidaStock Then
                                   setSqlAlm = setSqlAlm & "AND (ifnull(XZ.sc_stock,0) + ifnull(s.Stock,0)) > 0 "
                                End If
                                G.Columns.ColumnByFieldName("Stock").Visible = True
                                setSqlAlm = setSqlAlm & "AND (p.GlsProducto "
                 End If
                           
            Else
            
                setSqlAlm = "SELECT '0' AS CHK, p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda," & _
                            "if(p.afectoIGV = 1,'S','N') Afecto, Format(0,2) as Stock, t.GlsTallaPeso,p.AfectoIGV " & _
                            "FROM productos p " & _
                            "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                            "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                            "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
                            "INNER JOIN productosproveedor x ON x.idProducto = p.idProducto and x.idProducto = p.idProducto " & _
                            "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
                            "WHERE p.idEmpresa = '" & glsEmpresa & "' AND (p.GlsProducto "
            End If
        End If
    Else
        setSqlAlm = "SELECT '0' AS CHK, p.idProducto,p.GlsProducto,'' as GlsMarca,'' as GlsUM,o.GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto," & _
                    "Format(0,2) as Stock, '' AS GlsTallaPeso,p.idfabricante,p.AfectoIGV " & _
                    "FROM productos p,monedas o " & _
                    "WHERE p.idMoneda  = o.idMoneda AND p.idEmpresa = '" & glsEmpresa & "' AND p.idTipoProducto = '06002' AND (p.GlsProducto "
    End If
    
End Function

Public Sub ExecuteReturnTextAlm(ByVal strAlm As String, ByRef RsP As ADODB.Recordset, ByRef strCod As String, ByRef StrDes As String, ByRef strCodUM As String, ByVal ValidaStock As Boolean, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByVal indVarvale As Boolean, ByRef StrMsgError As String, Optional strParAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1, Optional psucursal As Boolean)
Dim IndAgregaAux                    As Boolean
On Error GoTo Err
    
    MousePointer = 0
    'Iniciamos Variables
    indAlmacen = True
    SRptBus(0) = ""
    indPedido = indVarPedido
    valeind = indVarvale
    
    sw_todas_sucursales = psucursal
    
    'Pasamos valores de parametros a las variables privadas a nivel de form
    indValidaStock = ValidaStock
    strCodLista = strVarCodLista
    strCodAlmacen = strAlm
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
        StrDes = SRptBus(1)
        strCodUM = SRptBus(2)
        
        Set G.DataSource = Nothing
        Set gPresentaciones.DataSource = Nothing
        
        If rsg.State = 1 Then rsg.Close
    Else
        
        IndAgregaAux = IndAgrega
        
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
                Exit Sub
            End If
        End If
        
        '--- Eliminamos los registros q no estan marcados
        IndAgrega = IndAgregaAux
        rsg.Filter = ""
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
        Set RsP = rsg.Clone(adLockReadOnly)
        If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    End If
    
    Unload Me
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub ExecuteKeyasciiReturnTextAlm(ByVal KeyAscii As Integer, strAlm As String, ByRef strCod As String, ByRef StrDes As String, ByRef strCodUM As String, ByVal ValidaStock As String, ByVal strVarCodLista As String, ByVal indVarUMVenta As Boolean, ByVal indVarMostrarPresentaciones As Boolean, ByVal indVarPedido As Boolean, ByRef StrMsgError As String, Optional strParAdic As String, Optional indMostrarTP As Boolean = True, Optional TipoProd As Integer = 1)
Dim IndAgregaAux                    As Boolean
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
        StrDes = SRptBus(1)
        strCodUM = SRptBus(2)
        
        Set G.DataSource = Nothing
        Set gPresentaciones.DataSource = Nothing
        
        If rsg.State = 1 Then rsg.Close
    Else
        
        IndAgregaAux = IndAgrega
        
        'Quitamos Filtros existentes
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
                Exit Sub
            End If
        End If
        
        'Eliminamos los registros q no estan marcados
        IndAgrega = IndAgregaAux
        rsg.Filter = ""
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
    
        Set RsP = rsg.Clone(adLockReadOnly)
        If rsg.State = 1 Then rsg.Close
        Set rsg = Nothing
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
    
     csql = "SELECT p.idUM,u.abreUM as GlsUM,CAST(r.factor AS NUMERIC(12,2)) AS factor, VVUnit, IGVUnit,PVUnit " & _
            "FROM PreciosVenta p " & _
            "INNER JOIN unidadMedida u " & _
              "ON p.idUM = u.idUM " & _
            "INNER JOIN  presentaciones r " & _
              "ON p.idUM = r.idUM " & _
              "AND p.idProducto = r.idProducto " & _
              "AND p.idEmpresa = r.idEmpresa " & _
            "WHERE p.idProducto = '" & G.Columns.ColumnByFieldName("idProducto").Value & "' " & _
            "AND p.idEmpresa = '" & glsEmpresa & "' " & _
            "ORDER BY r.factor ASC "
               
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
Dim i   As Integer

    '--- Limpiando Tag
    For i = 0 To 4
        txtCod_Nivel(i).Tag = ""
        txtCod_Nivel(i).Visible = False
    Next
    
    '--- Tipos nivel
    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    numPesos = Val("" & rsj.RecordCount)
    For i = 0 To numPesos - 1
        txtCod_Nivel(i).Tag = ""
        txtCod_Nivel(i).Visible = True
    Next
    fraNivel.Height = 355 * numPesos
    
    i = 0
    Do While Not rsj.EOF
        If (i + 1) = numPesos Then
            txtCod_Nivel(i).Tag = "TidNivel"
        End If
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
