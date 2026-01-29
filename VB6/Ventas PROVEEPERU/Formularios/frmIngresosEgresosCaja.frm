VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmIngresosEgresosCaja 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos y Egresos de Caja"
   ClientHeight    =   5010
   ClientLeft      =   4635
   ClientTop       =   1710
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8550
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
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
            Caption         =   "Anular"
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
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   4320
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
            Picture         =   "frmIngresosEgresosCaja.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresosEgresosCaja.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      Caption         =   "Listado"
      ForeColor       =   &H00000000&
      Height          =   4305
      Left            =   90
      TabIndex        =   3
      Top             =   660
      Width           =   8415
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   90
         TabIndex        =   29
         Top             =   180
         Width           =   8235
         Begin VB.CommandButton cmbAyudaCaja 
            Height          =   315
            Left            =   7740
            Picture         =   "frmIngresosEgresosCaja.frx":3518
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   180
            Width           =   390
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   780
            TabIndex        =   30
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42795009
            CurrentDate     =   38667
         End
         Begin CATControls.CATTextBox txtCod_CajaBuscar 
            Height          =   315
            Left            =   3300
            TabIndex        =   33
            Top             =   210
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
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
            Container       =   "frmIngresosEgresosCaja.frx":38A2
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_CajaBuscar 
            Height          =   315
            Left            =   4260
            TabIndex        =   34
            Top             =   210
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
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
            Container       =   "frmIngresosEgresosCaja.frx":38BE
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Caja:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   2820
            TabIndex        =   35
            Top             =   270
            Width           =   360
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   31
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   90
         TabIndex        =   5
         Top             =   840
         Width           =   8235
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   285
            Left            =   1140
            TabIndex        =   6
            Top             =   210
            Width           =   6960
            _ExtentX        =   12277
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
            MaxLength       =   255
            Container       =   "frmIngresosEgresosCaja.frx":38DA
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Busqueda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   255
            Width           =   765
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   2670
         Left            =   120
         OleObjectBlob   =   "frmIngresosEgresosCaja.frx":38F6
         TabIndex        =   4
         Top             =   1560
         Width           =   8220
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4305
      Left            =   90
      TabIndex        =   0
      Top             =   660
      Width           =   8415
      Begin VB.CommandButton cmbAyudaTipoMovCaja 
         Height          =   315
         Left            =   7740
         Picture         =   "frmIngresosEgresosCaja.frx":5E2D
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1180
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   7740
         Picture         =   "frmIngresosEgresosCaja.frx":61B7
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1530
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_MovCaja 
         Height          =   285
         Left            =   5160
         TabIndex        =   1
         Tag             =   "TidMovCaja"
         Top             =   240
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   12632319
         Enabled         =   0   'False
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
         Container       =   "frmIngresosEgresosCaja.frx":6541
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Caja 
         Height          =   285
         Left            =   1710
         TabIndex        =   9
         Top             =   840
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "frmIngresosEgresosCaja.frx":655D
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Caja 
         Height          =   285
         Left            =   2700
         TabIndex        =   10
         Top             =   840
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   503
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "frmIngresosEgresosCaja.frx":6579
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   285
         Left            =   1710
         TabIndex        =   13
         Tag             =   "TidMoneda"
         Top             =   1560
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
         Container       =   "frmIngresosEgresosCaja.frx":6595
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   285
         Left            =   2700
         TabIndex        =   14
         Top             =   1560
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   503
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "frmIngresosEgresosCaja.frx":65B1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtObs 
         Height          =   510
         Left            =   1710
         TabIndex        =   17
         Tag             =   "TGlsObs"
         Top             =   2310
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   900
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
         MaxLength       =   255
         Container       =   "frmIngresosEgresosCaja.frx":65CD
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_TipoCambio 
         Height          =   285
         Left            =   6810
         TabIndex        =   19
         Tag             =   "NValTipoCambio"
         Top             =   1920
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "frmIngresosEgresosCaja.frx":65E9
         Text            =   "0"
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_TipoMovCaja 
         Height          =   285
         Left            =   1710
         TabIndex        =   22
         Tag             =   "TidTipoMovCaja"
         Top             =   1200
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
         Container       =   "frmIngresosEgresosCaja.frx":6605
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TipoMovCaja 
         Height          =   285
         Left            =   2700
         TabIndex        =   23
         Top             =   1200
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   503
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "frmIngresosEgresosCaja.frx":6621
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_MovCajaDet 
         Height          =   285
         Left            =   6780
         TabIndex        =   25
         Tag             =   "TidMovCajaDet"
         Top             =   330
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "frmIngresosEgresosCaja.frx":663D
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_Monto 
         Height          =   285
         Left            =   1710
         TabIndex        =   26
         Tag             =   "NValMonto"
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
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
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "frmIngresosEgresosCaja.frx":6659
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   285
         Left            =   4200
         TabIndex        =   27
         Tag             =   "TidSucursal"
         Top             =   240
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   12632319
         Enabled         =   0   'False
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
         Container       =   "frmIngresosEgresosCaja.frx":6675
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.Label lblAnulado 
         Appearance      =   0  'Flat
         Caption         =   "Anulado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   420
         TabIndex        =   28
         Top             =   180
         Width           =   3015
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Tipo Mov. Caja:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   300
         TabIndex        =   24
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5580
         TabIndex        =   20
         Top             =   1980
         Width           =   930
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Obs:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   18
         Top             =   2430
         Width           =   330
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   1950
         Width           =   495
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         Caption         =   "Moneda:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   300
         TabIndex        =   15
         Top             =   1590
         Width           =   765
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   870
         Width           =   360
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nº:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6420
         TabIndex        =   2
         Top             =   360
         Width           =   225
      End
   End
End
Attribute VB_Name = "frmIngresosEgresosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public indIngresoSalida As String
Private strMovCaja As String
Private strEstado As String

Dim indCargando As Boolean


Private Sub cmbAyudaCaja_Click()
mostrarAyuda "CAJASUSUARIO", txtCod_CajaBuscar, txtGls_CajaBuscar
End Sub

Private Sub dtpFecha_Change()
If indCargando Then Exit Sub

Dim StrMsgError As String

On Error GoTo Err

    strMovCaja = traerCajaFecha(StrMsgError)
    If StrMsgError <> "" Then GoTo Err
    
    txtCod_MovCaja.Text = strMovCaja

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_CajaBuscar_Change()
Dim StrMsgError As String

On Error GoTo Err

    txtGls_CajaBuscar.Text = traerCampo("cajas", "GlsCaja", "idCaja", txtCod_CajaBuscar.Text, True)
    
    If indCargando Then Exit Sub
    
    strMovCaja = traerCajaFecha(StrMsgError)
    If StrMsgError <> "" Then GoTo Err
    
    txtCod_MovCaja.Text = strMovCaja
    
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub


Private Sub txtCod_CajaBuscar_KeyPress(KeyAscii As Integer)
 If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CAJASUSUARIO", txtCod_CajaBuscar, txtGls_CajaBuscar
        KeyAscii = 0
 End If
End Sub


Private Sub cmbAyudaMoneda_Click()
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
End Sub

Private Sub cmbAyudaTipoMovCaja_Click()
    mostrarAyuda "TIPOMOVCAJA", txtCod_TipoMovCaja, txtGls_TipoMovCaja, " AND indIngresoSalida = '" & indIngresoSalida & "'"
End Sub

Private Sub Form_Load()
Dim StrMsgError As String

Dim strCodCaja As String
Dim strGlsCaja As String

On Error GoTo Err

indCargando = False

DtpFecha.Value = Format(getFechaSistema, "dd/mm/yyyy")

FraListado.Visible = True
FraGeneral.Visible = False
habilitaBotones 7

If indIngresoSalida = "I" Then
    Me.Caption = "Ingresos de Caja "
Else
    Me.Caption = "Egresos de Caja "
End If

ConfGrid GLista, False, False, False, False

strMovCaja = CajaAperturadaUsuario(0, StrMsgError)
If StrMsgError <> "" Then GoTo Err

indCargando = True

txtCod_CajaBuscar.Text = traerCampo("movcajas", "idCaja", "idMovCaja", strMovCaja, True, " idSucursal = '" & glsSucursal & "'")
DtpFecha.Value = Format(traerCampo("movcajas", "FecCaja", "idMovCaja", strMovCaja, True, " idSucursal = '" & glsSucursal & "'"), "dd/mm/yyyy")

listaIngresosEgresos strMovCaja, StrMsgError
If StrMsgError <> "" Then GoTo Err

indCargando = False

txtVal_TipoCambio.Decimales = glsDecimalesTC

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
Dim StrCodigo As String
Dim strMsg      As String

On Error GoTo Err

validaFormSQL Me, StrMsgError
If StrMsgError <> "" Then GoTo Err

If left(txtCod_TipoMovCaja.Text, 4) <> "9999" Then
    If Val("" & txtVal_Monto.Value) <= 0 Then
        StrMsgError = "Debe ingresar un monto"
        txtVal_Monto.OnError = True
        GoTo Err
    End If
End If

If txtCod_MovCajaDet.Text = "" Then 'graba
    txtCod_MovCajaDet.Text = GeneraCorrelativoAnoMes("movcajasdet", "idMovCajaDet")
    
    EjecutaSQLForm Me, 0, True, "movcajasdet", StrMsgError, , , , , , True
    If StrMsgError <> "" Then GoTo Err
    
    strMsg = "Grabo"
Else 'modifica

    EjecutaSQLForm Me, 1, True, "movcajasdet", StrMsgError, "idMovCajaDet"
    If StrMsgError <> "" Then GoTo Err
    
    strMsg = "Modifico"
End If
MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title

listaIngresosEgresos txtCod_MovCaja.Text, StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo(ByRef StrMsgError As String)
On Error GoTo Err

    limpiaFormCaja
    
    strEstado = "ACT"
    
    lblAnulado.Visible = False
    
    txtCod_Sucursal.Text = glsSucursal

    strMovCaja = traerCajaFecha(StrMsgError)
    If StrMsgError <> "" Then GoTo Err
    
    If strMovCaja = "" Then
        txtCod_MovCaja.Text = ""
        Exit Sub
    End If
    
    txtCod_MovCaja.Text = strMovCaja
    
    txtCod_Caja.Text = traerCampo("movcajas", "idCaja", "idMovCaja", strMovCaja, True, " idSucursal = '" & glsSucursal & "'")
    
    txtVal_TipoCambio.Text = glsTC
    
    txtCod_TipoMovCaja.Enabled = True
    cmbAyudaTipoMovCaja.Enabled = True
    txtCod_Moneda.Enabled = True
    cmbAyudaMoneda.Enabled = True
    
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gLista_OnDblClick()
Dim StrMsgError As String
On Error GoTo Err

If GLista.Columns.ColumnByName("idMovCajaDet").Value = "" Then Exit Sub

mostrarIngresosEgresos GLista.Columns.ColumnByName("idMovCajaDet").Value, StrMsgError
If StrMsgError <> "" Then GoTo Err

If left(txtCod_TipoMovCaja.Text, 4) = "9999" Then
    txtCod_TipoMovCaja.Enabled = False
    cmbAyudaTipoMovCaja.Enabled = False
    txtCod_Moneda.Enabled = False
    cmbAyudaMoneda.Enabled = False
Else
    txtCod_TipoMovCaja.Enabled = True
    cmbAyudaTipoMovCaja.Enabled = True
    txtCod_Moneda.Enabled = True
    cmbAyudaMoneda.Enabled = True
End If

FraListado.Visible = False
FraGeneral.Visible = True
FraGeneral.Enabled = False
habilitaBotones 2
Exit Sub
Err:
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
Dim indEvaluacion As Integer
Dim strCodUsuarioAutorizacion As String
On Error GoTo Err
Select Case Button.Index
    Case 1 'Nuevo
        nuevo StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        If strMovCaja = "" Then Exit Sub
        
        If traerCampo("movcajas", "indEstado", "idMovCaja", strMovCaja, True, " idSucursal = '" & glsSucursal & "'") = "C" Then
            StrMsgError = "La caja se encuentra cerrada"
            GoTo Err
        End If
        
        FraListado.Visible = False
        FraGeneral.Visible = True
        FraGeneral.Enabled = True
    Case 2 'Grabar
    
        If indIngresoSalida = "S" Then
            'indEvaluacion = 0
                
            'frmAprobacion.MostrarForm "07",  indEvaluacion, strCodUsuarioAutorizacion, strMsgError
            'If strMsgError <> "" Then GoTo ERR
            
            'If indEvaluacion = 0 Then
             '   Exit Sub
            'End If
        End If
    
        Grabar StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Case 3 'Modificar
    
        If traerCampo("movcajas", "indEstado", "idMovCaja", txtCod_MovCaja.Text, True, " idSucursal = '" & glsSucursal & "'") = "C" Then
            StrMsgError = "La caja se encuentra cerrada"
            GoTo Err
        End If

        FraGeneral.Enabled = True
    Case 5 'Anular
        anularMovCaja StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Case 4, 7  'Cancelar
        FraListado.Visible = True
        FraGeneral.Visible = False
        FraGeneral.Enabled = False
    Case 6 'Imprimir
        'imprimeReciboCaja txtCod_MovCajaDet.Text, StrMsgError
        'If StrMsgError <> "" Then GoTo ERR
        
        mostrarReporte "rptImpRecibosCaja.rpt", "parEmpresa|parSucursal|parMovCajaDet|parTipoMovCaja", glsEmpresa & "|" & glsSucursal & "|" & txtCod_MovCajaDet.Text & "|" & indIngresoSalida, GlsForm, StrMsgError
        If StrMsgError <> "" Then GoTo Err
  
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
        Toolbar1.Buttons(5).Visible = indHabilitar 'Anular
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

If strEstado = "ANU" And indexBoton <> 8 Then
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(6).Visible = False
End If

End Sub

Private Sub txt_TextoBuscar_Change()
Dim StrMsgError As String
On Error GoTo Err
If strCodMovCaja <> "" Then
    listaIngresosEgresos strCodMovCaja, StrMsgError
    If StrMsgError <> "" Then GoTo Err
End If
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then GLista.SetFocus
End Sub

Private Sub listaIngresosEgresos(ByVal strVarMovCaja As String, ByRef StrMsgError As String)
Dim strCond As String
On Error GoTo Err
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND m.GlsTipoMovCaja LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT m.idMovCajaDet ,t.GlsTipoMovCaja,o.GlsMoneda ,Format(m.ValMonto,2) AS ValMonto,m.estMovCajaDet " & _
           "FROM movcajasdet m,tiposmovcaja t,monedas o " & _
           " WHERE m.idTipoMovCaja = t.idTipoMovCaja " & _
           " AND m.idMoneda = o.idMoneda " & _
           " AND m.idEmpresa = '" & glsEmpresa & "'" & _
           " AND m.idSucursal = '" & glsSucursal & "'" & _
           " AND t.indIngresoSalida = '" & indIngresoSalida & "'" & _
           " AND m.idMovCaja = '" & strVarMovCaja & "'"
           
    If strCond <> "" Then csql = csql + strCond

    csql = csql + " ORDER BY m.idMovCajaDet"
    With GLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idMovCajaDet"
End With
Me.Refresh
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarIngresosEgresos(StrCod As String, ByRef StrMsgError As String)
Dim rst As New ADODB.Recordset
On Error GoTo Err
    csql = "SELECT m.idMovCaja,m.idMovCajaDet,c.idCaja,m.idTipoMovCaja,m.idMoneda,m.ValMonto,m.ValTipoCambio,m.GlsObs,c.idSucursal,m.estMovCajaDet " & _
           "FROM movcajasdet m,MovCajas c " & _
           "WHERE m.idMovCaja = c.idMovCaja AND m.idMovCajaDet = '" & StrCod & "' AND m.idEmpresa = '" & glsEmpresa & "' " & _
             "AND m.idSucursal = '" & glsSucursal & "'  AND c.idEmpresa = '" & glsEmpresa & "' AND c.idSucursal = '" & glsSucursal & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    txtCod_Caja.Tag = "TidCaja"
    
    If Not rst.EOF Then strEstado = "" & rst.Fields("estMovCajaDet")
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If strEstado = "ANU" Then
        lblAnulado.Visible = True
    Else
        lblAnulado.Visible = False
    End If
    
    txtCod_Caja.Tag = ""
    
Me.Refresh
Exit Sub
Err:
txtCod_Caja.Tag = ""
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Caja_Change()
    txtGls_Caja.Text = traerCampo("cajas", "GlsCaja", "idCaja", txtCod_Caja.Text, True)
End Sub

Private Sub txtCod_Moneda_Change()
    txtGls_Moneda.Text = traerCampo("monedas", "glsMoneda", "idMoneda", txtCod_Moneda.Text, False)
End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    mostrarAyudaKeyascii KeyAscii, "MONEDA", txtCod_Moneda, txtGls_Moneda
    KeyAscii = 0
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
End If
End Sub

Private Sub txtCod_TipoMovCaja_Change()
    txtGls_TipoMovCaja.Text = traerCampo("tiposmovcaja", "GlsTipoMovCaja", "idTipoMovCaja", txtCod_TipoMovCaja.Text, False)
End Sub

Private Sub txtCod_TipoMovCaja_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    mostrarAyudaKeyascii KeyAscii, "TIPOMOVCAJA", txtCod_TipoMovCaja, txtGls_TipoMovCaja, " AND indIngresoSalida = '" & indIngresoSalida & "'"
    KeyAscii = 0
    If txtCod_TipoMovCaja.Text <> "" Then SendKeys "{tab}"
End If
End Sub

Private Sub anularMovCaja(ByRef StrMsgError As String)

On Error GoTo Err

    If left(txtCod_TipoMovCaja.Text, 4) = "9999" Then
        StrMsgError = "Este tipo de movimiento de caja no se puede anular"
        GoTo Err
    End If

    If traerCampo("movcajas", "indEstado", "idMovCaja", txtCod_MovCaja.Text, True, " idSucursal = '" & glsSucursal & "'") = "C" Then
        StrMsgError = "La caja se encuentra cerrada"
        GoTo Err
    End If

    If MsgBox("Seguro de Anular el Documento", vbQuestion + vbYesNo, App.Title) = vbYes Then
    
        csql = "UPDATE movcajasdet SET estMovCajaDet = 'ANU' " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' " & _
                 "AND idMovCaja = '" & txtCod_MovCaja.Text & "' AND idMovCajaDet = '" & txtCod_MovCajaDet.Text & "'"
                 
        Cn.Execute csql
        
        strEstado = "ANU"
        
        lblAnulado.Visible = True
        
        listaIngresosEgresos txtCod_MovCaja.Text, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub


Private Function traerCajaFecha(ByRef StrMsgError As String) As String
Dim rsTemp As New ADODB.Recordset
Dim strCodMovCaja As String
Dim strGlsCaja As String
Dim strEstado As String

On Error GoTo Err

strCodMovCaja = ""

csql = "SELECT m.idMovCaja,m.indEstado " & _
        "FROM movcajas m " & _
        "WHERE m.idUsuario = '" & glsUser & "' " & _
         "AND m.idEmpresa = '" & glsEmpresa & "' " & _
         "AND m.idSucursal = '" & glsSucursal & "' " & _
         "AND m.idCaja = '" & txtCod_CajaBuscar.Text & "' " & _
         "AND DATE_FORMAT(m.FecCaja ,'%d/%m/%Y') = DATE_FORMAT('" & Format(DtpFecha.Value, "yyyy-mm-dd") & "','%d/%m/%Y')"
         
rsTemp.Open csql, Cn, adOpenKeyset, adLockOptimistic
If Not rsTemp.EOF Then
    strCodMovCaja = "" & rsTemp.Fields("idMovCaja")
    strEstado = "" & rsTemp.Fields("indEstado")
Else
    If txtCod_CajaBuscar.Text <> "" Then
        MsgBox "No hay caja disponible para la fecha y caja indicada", vbInformation, App.Title
    End If
End If

strGlsCaja = traerCampo("cajas", "GlsCaja", "idCaja", txtCod_CajaBuscar.Text, True)

If indIngresoSalida = "I" Then
    Me.Caption = "Ingresos de Caja - "
Else
    Me.Caption = "Egresos de Caja - "
End If

Me.Caption = Me.Caption & strGlsCaja & " al " & Format(DtpFecha.Value, "dd/mm/yyyy")
If strEstado = "A" Then Me.Caption = Me.Caption & " (Aperturada)"
If strEstado = "C" Then Me.Caption = Me.Caption & " (Cerrada)"


listaIngresosEgresos strCodMovCaja, StrMsgError
If StrMsgError <> "" Then GoTo Err

traerCajaFecha = strCodMovCaja

If rsTemp.State = 1 Then rsTemp.Close
Set rsTemp = Nothing
Exit Function
Err:
If rsTemp.State = 1 Then rsTemp.Close
Set rsTemp = Nothing
If StrMsgError = "" Then StrMsgError = Err.Description
End Function

Public Sub limpiaFormCaja()
Dim C As Object

For Each C In Me.Controls

    If TypeOf C Is TextBox Or TypeOf C Is CATTextBox Then
        If C.Name <> txtCod_CajaBuscar.Name And C.Name <> txtGls_CajaBuscar.Name Then
            C.Text = ""
        End If
    End If
    
    If TypeOf C Is DTPicker Then
        If C.Name <> DtpFecha.Name Then
            C.Value = getFechaSistema
        End If
    End If
Next
End Sub
