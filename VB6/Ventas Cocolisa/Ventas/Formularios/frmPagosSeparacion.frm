VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmPagosSeparacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagos por Separacion"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTotales 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Totales"
      ForeColor       =   &H00C00000&
      Height          =   1725
      Left            =   5640
      TabIndex        =   24
      Top             =   5460
      Width           =   4060
      Begin CATControls.CATTextBox txt_TotalRecibido 
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":0000
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Vuelto 
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   600
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":001C
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total Recibido:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   1020
         TabIndex        =   28
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Vuelto S/.:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   1020
         TabIndex        =   27
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.Frame fraVuelto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vuelto - Entregado"
      ForeColor       =   &H00C00000&
      Height          =   1725
      Left            =   60
      TabIndex        =   20
      Top             =   5460
      Width           =   5535
      Begin CATControls.CATTextBox txtVal_VueltoEntregado 
         Height          =   285
         Left            =   3720
         TabIndex        =   21
         Top             =   1380
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":0038
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gVuelto 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmPagosSeparacion.frx":0054
         TabIndex        =   22
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total Vuelto Entregado S/.:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   1620
         TabIndex        =   23
         Top             =   1440
         Width           =   2025
      End
   End
   Begin VB.Frame frmPagos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Pagos"
      ForeColor       =   &H00C00000&
      Height          =   2805
      Left            =   60
      TabIndex        =   18
      Top             =   2640
      Width           =   9675
      Begin DXDBGRIDLibCtl.dxDBGrid gPagos 
         Height          =   2475
         Left            =   60
         OleObjectBlob   =   "frmPagosSeparacion.frx":2130
         TabIndex        =   19
         Top             =   240
         Width           =   9495
      End
   End
   Begin VB.Frame fraDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Datos del Documento"
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   1875
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   9675
      Begin CATControls.CATTextBox txt_NumDoc 
         Height          =   315
         Left            =   7860
         TabIndex        =   2
         Top             =   150
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   16775664
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         ForeColor       =   -2147483640
         Container       =   "frmPagosSeparacion.frx":5A8D
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Serie 
         Height          =   315
         Left            =   5685
         TabIndex        =   3
         Top             =   150
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         BackColor       =   16775664
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         ForeColor       =   -2147483640
         Container       =   "frmPagosSeparacion.frx":5AA9
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TipoCambio 
         Height          =   285
         Left            =   7860
         TabIndex        =   4
         Top             =   675
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":5AC5
         Text            =   "0"
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalBruto 
         Height          =   285
         Left            =   900
         TabIndex        =   5
         Top             =   1050
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":5AE1
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalIGV 
         Height          =   285
         Left            =   4335
         TabIndex        =   6
         Top             =   1050
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":5AFD
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalNeto 
         Height          =   285
         Left            =   7860
         TabIndex        =   7
         Top             =   1050
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":5B19
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   285
         Left            =   900
         TabIndex        =   8
         Top             =   675
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":5B35
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   285
         Left            =   1875
         TabIndex        =   9
         Top             =   675
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":5B51
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Adelantos 
         Height          =   285
         Left            =   885
         TabIndex        =   32
         Top             =   1440
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":5B6D
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Saldo 
         Height          =   285
         Left            =   4320
         TabIndex        =   33
         Top             =   1440
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   16775664
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
         Container       =   "frmPagosSeparacion.frx":5B89
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Adelantos:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   60
         TabIndex        =   35
         Top             =   1515
         Width           =   795
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Saldo:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3645
         TabIndex        =   34
         Top             =   1500
         Width           =   645
      End
      Begin VB.Label lbl_TC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "T/C:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7035
         TabIndex        =   17
         Top             =   675
         Width           =   765
      End
      Begin VB.Label lbl_NumDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Numero:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6960
         TabIndex        =   16
         Top             =   225
         Width           =   915
      End
      Begin VB.Label lbl_Serie 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Serie:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   4935
         TabIndex        =   15
         Top             =   225
         Width           =   615
      End
      Begin VB.Label lbl_TotalNeto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7035
         TabIndex        =   14
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label lbl_TotalIGV 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "IGV:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3660
         TabIndex        =   13
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label lbl_TotalBruto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bruto:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   75
         TabIndex        =   12
         Top             =   1125
         Width           =   615
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Moneda:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   75
         TabIndex        =   11
         Top             =   750
         Width           =   765
      End
      Begin VB.Label lblDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Boleta de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   225
         Width           =   3765
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":5BA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":5F3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":6391
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":672B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":6AC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":6E5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":71F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":7593
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":792D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":7CC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":8061
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":8D23
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":90BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":950F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":98A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagosSeparacion.frx":A2BB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   1164
      ButtonWidth     =   1111
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin CATControls.CATTextBox txtCod_Caja 
      Height          =   285
      Left            =   1140
      TabIndex        =   29
      Top             =   7260
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      BackColor       =   16775664
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
      Container       =   "frmPagosSeparacion.frx":A98D
      Estilo          =   1
      EnterTab        =   -1  'True
   End
   Begin CATControls.CATTextBox txtGls_Caja 
      Height          =   285
      Left            =   2100
      TabIndex        =   30
      Top             =   7275
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   503
      BackColor       =   16775664
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
      Container       =   "frmPagosSeparacion.frx":A9A9
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Caja Activa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   7320
      Width           =   975
   End
End
Attribute VB_Name = "frmPagosSeparacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strTD As String
Private strNumDoc As String
Private strSerie As String

Private strCodCliente As String

Private strEstDoc As String
Private indInserta  As Boolean

Private Sub Form_Load()
Dim strMsgError As String
On Error GoTo ERR

indInserta = False

txt_TipoCambio.Decimales = glsDecimalesTC
txt_TotalBruto.Decimales = glsDecimalesCaja
txt_TotalIGV.Decimales = glsDecimalesCaja
txt_TotalNeto.Decimales = glsDecimalesCaja
txt_TotalRecibido.Decimales = glsDecimalesCaja
txt_Vuelto.Decimales = glsDecimalesCaja
txtVal_VueltoEntregado.Decimales = glsDecimalesCaja
txt_Adelantos.Decimales = glsDecimalesCaja
txt_Saldo.Decimales = glsDecimalesCaja

gPagos.Columns.ColumnByFieldName("MontoOri").DecimalPlaces = glsDecimalesCaja
gPagos.Columns.ColumnByFieldName("MontoSoles").DecimalPlaces = glsDecimalesCaja

gVuelto.Columns.ColumnByFieldName("Vuelto").DecimalPlaces = glsDecimalesCaja

ConfGrid gPagos, True, False, False, False
ConfGrid gVuelto, True, False, False, False

Exit Sub
ERR:
If strMsgError = "" Then strMsgError = ERR.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub grabar(ByRef strMsgError As String)
Dim strCodigo As String
Dim strMsg      As String
Dim strCodMovCaja As String

Dim cselect         As String
Dim rsp             As New ADODB.Recordset

On Error GoTo ERR

validaFormSQL Me, strMsgError
If strMsgError <> "" Then GoTo ERR

eliminaNulosGrilla

If gPagos.Count >= 1 Then
    If gPagos.Count = 1 And gPagos.Columns.ColumnByFieldName("idFormadePago").Value = "" Then
        strMsgError = "Falta Ingresar Pagos"
        GoTo ERR
    End If
End If

'''If strTD <> "91" Then
'''    If Val(txt_TotalRecibido.Value) < Val(txt_TotalNeto.Value) Then
'''        strMsgError = "El Monto recibido es menor al Monto por pagar"
'''        GoTo err
'''    End If
'''End If

'''strCodMovCaja = Trim(traerCampo("docventas", "idMovCaja", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))
'''If strCodMovCaja = "" Then
    strCodMovCaja = CajaAperturadaUsuario(0, strMsgError)
    If strMsgError <> "" Then GoTo ERR
'''End If

EjecutaSQLFormPagosSeparacion Me, strMsgError, strTD, strNumDoc, strSerie, strCodMovCaja, gPagos, gVuelto
If strMsgError <> "" Then GoTo ERR

Unload Me

Exit Sub
ERR:
If strMsgError = "" Then strMsgError = ERR.Description
End Sub

Private Sub nuevo(ByRef strMsgError As String)
    Dim rst As New ADODB.Recordset
    Dim rsv As New ADODB.Recordset

    '********FORMATO GRILLA
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "idFormadePago", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "glsFormadePago", adVarChar, 185, adFldIsNullable
    rst.Fields.Append "idMoneda", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsMoneda", adVarChar, 185, adFldIsNullable
    rst.Fields.Append "MontoOri", adDouble, 14, adFldIsNullable
    rst.Fields.Append "MontoSoles", adDouble, 14, adFldIsNullable
    rst.Fields.Append "idTipoFormaPago", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "fecVctos", adVarChar, 185, adFldIsNullable

    rst.Open
    rst.AddNew

    rst.Fields("Item") = 1
    rst.Fields("idFormadePago") = glsFormaPagoVentas
    rst.Fields("glsFormadePago") = traerCampo("formaspagos", "GlsFormaPago", "idFormaPago", glsFormaPagoVentas, True)
    ''''rst.Fields("idMoneda") = "PEN"
    rst.Fields("idMoneda") = txtCod_Moneda.Text
    rst.Fields("GlsMoneda") = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
    ''''rst.Fields("MontoOri") = txt_TotalNeto.Value
    ''''rst.Fields("MontoSoles") = txt_TotalNeto.Value
    rst.Fields("MontoOri") = txt_Saldo.Value
    
    If txtCod_Moneda.Text = "PEN" Then
        rst.Fields("MontoSoles") = txt_Saldo.Value
    Else
        rst.Fields("MontoSoles") = txt_Saldo.Value * txt_TipoCambio.Text
    End If
    
    rst.Fields("idTipoFormaPago") = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", glsFormaPagoVentas, True)
    
'''    If rst.Fields("idTipoFormaPago") = "06090002" Then
'''        rst.Fields("fecVctos") = Format(DateAdd("D", traerCampo("formaspagos", "diasVcto", "idFormaPago", glsFormaPagoVentas, True), CDate(fecEmision)), "dd/mm/yyyy")
'''    Else
        rst.Fields("fecVctos") = ""
'''    End If
        
'''    listaPagos strMsgError
'''    If strMsgError <> "" Then GoTo err

    mostrarDatosGridSQL gPagos, rst, strMsgError
    If strMsgError <> "" Then GoTo ERR
    
    
    'Formato grilla de vueltos
    rsv.Fields.Append "Item", adInteger, , adFldRowID
    rsv.Fields.Append "idMoneda", adVarChar, 8, adFldIsNullable
    rsv.Fields.Append "GlsMoneda", adVarChar, 185, adFldIsNullable
    rsv.Fields.Append "Vuelto", adDouble, 14, adFldIsNullable
        
    rsv.Open
    rsv.AddNew

    rsv.Fields("Item") = 1
    rsv.Fields("idMoneda") = ""
    rsv.Fields("GlsMoneda") = ""
    rsv.Fields("Vuelto") = 0
    
    mostrarDatosGridSQL gVuelto, rsv, strMsgError
    If strMsgError <> "" Then GoTo ERR
    
    
    calcularTotales
    
    gPagos.Columns.FocusedIndex = gPagos.Columns.ColumnByFieldName("idFormadePago").Index
    Exit Sub
ERR:
    If strMsgError = "" Then strMsgError = ERR.Description
End Sub

Private Sub gPagos_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer
If Action = daInsert Then
    gPagos.Columns.ColumnByFieldName("item").Value = gPagos.Count
    If txt_TotalNeto.Value - txt_TotalRecibido.Value Then
        gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN"
        gPagos.Columns.ColumnByFieldName("GlsMoneda").Value = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
        gPagos.Columns.ColumnByFieldName("MontoSoles").Value = txt_TotalNeto.Value - txt_TotalRecibido.Value
        gPagos.Columns.ColumnByFieldName("MontoOri").Value = txt_TotalNeto.Value - txt_TotalRecibido.Value
    Else
        gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN"
        gPagos.Columns.ColumnByFieldName("GlsMoneda").Value = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
        gPagos.Columns.ColumnByFieldName("MontoSoles").Value = 0
        gPagos.Columns.ColumnByFieldName("MontoOri").Value = 0
    End If
    gPagos.Dataset.Post
End If
End Sub

Private Sub gPagos_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
If Action = daInsert Then
    If gPagos.Columns.ColumnByFieldName("idFormadePago").Value = "" And indInserta = False Then
        Allow = False
    Else
        gPagos.Columns.FocusedIndex = gPagos.Columns.ColumnByFieldName("idFormadePago").Index
    End If
End If
End Sub

Private Sub gPagos_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Dim strCod As String
    Dim strDes As String
    
    Select Case Column.Index
        Case gPagos.Columns.ColumnByFieldName("idFormadePago").Index
            strCod = gPagos.Columns.ColumnByFieldName("idFormadePago").Value
            strDes = gPagos.Columns.ColumnByFieldName("glsFormadePago").Value
            
            mostrarAyudaTexto "FORMASPAGO", strCod, strDes
            
            gPagos.Dataset.Edit
            gPagos.Columns.ColumnByFieldName("idFormadePago").Value = strCod
            gPagos.Columns.ColumnByFieldName("glsFormadePago").Value = strDes
            gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", strCod, True)
            If gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = "06090002" Then
                gPagos.Columns.ColumnByFieldName("fecVctos").Value = Format(DateAdd("D", traerCampo("formaspagos", "diasVcto", "idFormaPago", strCod, True), CDate(fecEmision)), "dd/mm/yyyy")
            Else
                gPagos.Columns.ColumnByFieldName("fecVctos").Value = ""
            End If
            gPagos.Dataset.Post
        Case gPagos.Columns.ColumnByFieldName("idMoneda").Index
            strCod = gPagos.Columns.ColumnByFieldName("idMoneda").Value
            strDes = gPagos.Columns.ColumnByFieldName("glsMoneda").Value
            
            mostrarAyudaTexto "MONEDA", strCod, strDes
            
            gPagos.Dataset.Edit
            gPagos.Columns.ColumnByFieldName("idMoneda").Value = strCod
            gPagos.Columns.ColumnByFieldName("glsMoneda").Value = strDes
            If strCod = "PEN" Then 'soles
                gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value
            Else 'dolares
                gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * txt_TipoCambio.Value
            End If
            gPagos.Dataset.Post
    End Select
    calcularTotales
End Sub

Private Sub gPagos_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Select Case gPagos.Columns.FocusedColumn.Index
        Case gPagos.Columns.ColumnByFieldName("MontoOri").Index
            gPagos.Dataset.Edit
            If gPagos.Columns.ColumnByFieldName("idMoneda").Value = "PEN" Then 'soles
                gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value
            Else 'dolares
                gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * txt_TipoCambio.Value
            End If
            gPagos.Dataset.Post
            calcularTotales
End Select
End Sub

Private Sub gPagos_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer
If KeyCode = 46 Then
    If gPagos.Count > 0 Then
        If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                       
            If gPagos.Count = 1 Then
                gPagos.Dataset.Edit
                gPagos.Columns.ColumnByFieldName("Item").Value = 1
                gPagos.Columns.ColumnByFieldName("idFormadePago").Value = ""
                gPagos.Columns.ColumnByFieldName("glsFormadePago").Value = ""
                gPagos.Columns.ColumnByFieldName("idMoneda").Value = ""
                gPagos.Columns.ColumnByFieldName("GlsMoneda").Value = ""
                gPagos.Columns.ColumnByFieldName("MontoOri").Value = 0#
                gPagos.Columns.ColumnByFieldName("MontoSoles").Value = 0#
                gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = ""
                gPagos.Columns.ColumnByFieldName("fecVctos").Value = ""
                gPagos.Dataset.Post
            Else
                gPagos.Dataset.Delete
                
                gPagos.Dataset.First
                Do While Not gPagos.Dataset.EOF
                    i = i + 1
                    
                    gPagos.Dataset.Edit
                    gPagos.Columns.ColumnByFieldName("Item").Value = i
                    gPagos.Dataset.Post
                    
                    gPagos.Dataset.Next
                Loop
                If gPagos.Dataset.State = dsEdit Or gPagos.Dataset.State = dsInsert Then
                    gPagos.Dataset.Post
                End If
            End If
            calcularTotales
        End If
    End If
End If
If KeyCode = 13 Then
    If gPagos.Dataset.State = dsEdit Or gPagos.Dataset.State = dsInsert Then
          gPagos.Dataset.Post
    End If
End If
End Sub

Private Sub gPagos_OnKeyPress(Key As Integer)
Dim strCod As String
Dim strDes As String
If Key <> 9 And Key <> 13 And Key <> 27 Then
    Select Case gPagos.Columns.FocusedColumn.Index
        Case gPagos.Columns.ColumnByFieldName("idFormadePago").Index
            strCod = gPagos.Columns.ColumnByFieldName("idFormadePago").Value
            strDes = gPagos.Columns.ColumnByFieldName("glsFormadePago").Value
            
            mostrarAyudaKeyasciiTexto Key, "FORMASPAGO", strCod, strDes
            Key = 0
            gPagos.Dataset.Edit
            gPagos.Columns.ColumnByFieldName("idFormadePago").Value = strCod
            gPagos.Columns.ColumnByFieldName("glsFormadePago").Value = strDes
            gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", strCod, True)
            If gPagos.Columns.ColumnByFieldName("idTipoFormaPago").Value = "06090002" Then 'CREDITO
                gPagos.Columns.ColumnByFieldName("fecVctos").Value = Format(DateAdd("D", traerCampo("formaspagos", "diasVcto", "idFormaPago", strCod, True), CDate(fecEmision)), "dd/mm/yyyy")
            Else
                gPagos.Columns.ColumnByFieldName("fecVctos").Value = ""
            End If
            gPagos.Dataset.Post
            calcularTotales
        Case gPagos.Columns.ColumnByFieldName("idMoneda").Index
            strCod = gPagos.Columns.ColumnByFieldName("idMoneda").Value
            strDes = gPagos.Columns.ColumnByFieldName("glsMoneda").Value
            
            mostrarAyudaKeyasciiTexto Key, "MONEDA", strCod, strDes
            Key = 0
            gPagos.Dataset.Edit
            gPagos.Columns.ColumnByFieldName("idMoneda").Value = strCod
            gPagos.Columns.ColumnByFieldName("glsMoneda").Value = strDes
            If strCod = "PEN" Then 'soles
                gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value
            Else 'dolares
                gPagos.Columns.ColumnByFieldName("MontoSoles").Value = gPagos.Columns.ColumnByFieldName("MontoOri").Value * txt_TipoCambio.Value
            End If
            gPagos.Dataset.Post
            calcularTotales
    End Select
End If
End Sub


'***************************************************
Private Sub gvuelto_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer
If Action = daInsert Then
    gVuelto.Columns.ColumnByFieldName("item").Value = gVuelto.Count
'    If txt_TotalNeto.Value - txt_TotalRecibido.Value Then
'        gVuelto.Columns.ColumnByFieldName("idMoneda").Value = "PEN"
'        gVuelto.Columns.ColumnByFieldName("GlsMoneda").Value = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
'        gVuelto.Columns.ColumnByFieldName("MontoSoles").Value = txt_TotalNeto.Value - txt_TotalRecibido.Value
'        gVuelto.Columns.ColumnByFieldName("MontoOri").Value = txt_TotalNeto.Value - txt_TotalRecibido.Value
'    Else
'        gVuelto.Columns.ColumnByFieldName("idMoneda").Value = "PEN"
'        gVuelto.Columns.ColumnByFieldName("GlsMoneda").Value = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
'        gVuelto.Columns.ColumnByFieldName("MontoSoles").Value = 0
'        gVuelto.Columns.ColumnByFieldName("MontoOri").Value = 0
'    End If
    gVuelto.Dataset.Post
End If
End Sub

Private Sub gvuelto_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
If Action = daInsert Then
    If gVuelto.Columns.ColumnByFieldName("idMoneda").Value = "" And indInserta = False Then
        Allow = False
    Else
        gVuelto.Columns.FocusedIndex = gVuelto.Columns.ColumnByFieldName("idMoneda").Index
    End If
End If
End Sub

Private Sub gvuelto_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Dim strCod As String
    Dim strDes As String
    
    Select Case Column.Index
        Case gVuelto.Columns.ColumnByFieldName("idMoneda").Index
            strCod = gVuelto.Columns.ColumnByFieldName("idMoneda").Value
            strDes = gVuelto.Columns.ColumnByFieldName("glsMoneda").Value
            
            mostrarAyudaTexto "MONEDA", strCod, strDes
            
            If existeEnGrilla(gVuelto, "idMoneda", strCod) = False Then
                gVuelto.Dataset.Edit
                gVuelto.Columns.ColumnByFieldName("idMoneda").Value = strCod
                gVuelto.Columns.ColumnByFieldName("glsMoneda").Value = strDes
            Else
                MsgBox "La Moneda ya fue ingresada", vbInformation, App.Title
                Exit Sub
            End If
            If gVuelto.Count = 1 Then txtVal_VueltoEntregado.Text = 0
            
            'regresa valores para recalcular
            gVuelto.Columns.ColumnByFieldName("Vuelto").Value = 0#
            
            gVuelto.Dataset.Post
            
            calcularTotalVueltoEntregado
            
            gVuelto.Dataset.Edit
            
            'calcula saldo
            If strCod = "PEN" Then
                gVuelto.Columns.ColumnByFieldName("Vuelto").Value = (txt_Vuelto.Value - txtVal_VueltoEntregado.Value)
            Else
                gVuelto.Columns.ColumnByFieldName("Vuelto").Value = ((txt_Vuelto.Value - txtVal_VueltoEntregado.Value) / txt_TipoCambio.Value)
            End If
            
            gVuelto.Dataset.Post
            
            gVuelto.Columns.FocusedIndex = gVuelto.Columns.ColumnByFieldName("Vuelto").Index
            
            calcularTotalVueltoEntregado
    End Select
    
End Sub

Private Sub gvuelto_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Select Case gVuelto.Columns.FocusedColumn.Index
        Case gVuelto.Columns.ColumnByFieldName("Vuelto").Index
'            gVuelto.Dataset.Edit
'            If gVuelto.Columns.ColumnByFieldName("idMoneda").Value = "PEN" Then 'soles
'                gVuelto.Columns.ColumnByFieldName("MontoSoles").Value = gVuelto.Columns.ColumnByFieldName("MontoOri").Value
'            Else 'dolares
'                gVuelto.Columns.ColumnByFieldName("MontoSoles").Value = gVuelto.Columns.ColumnByFieldName("MontoOri").Value * txt_TipoCambio.Value
'            End If
'            gVuelto.Dataset.Post
            calcularTotalVueltoEntregado
End Select
End Sub

Private Sub gvuelto_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer
If KeyCode = 46 Then
    If gVuelto.Count > 0 Then
        If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                       
            If gVuelto.Count = 1 Then
                gVuelto.Dataset.Edit
                gVuelto.Columns.ColumnByFieldName("Item").Value = 1
                gVuelto.Columns.ColumnByFieldName("idMoneda").Value = ""
                gVuelto.Columns.ColumnByFieldName("GlsMoneda").Value = ""
                gVuelto.Columns.ColumnByFieldName("Vuelto").Value = 0#
                gVuelto.Dataset.Post
            Else
                gVuelto.Dataset.Delete
                
                gVuelto.Dataset.First
                Do While Not gVuelto.Dataset.EOF
                    i = i + 1
                    
                    gVuelto.Dataset.Edit
                    gVuelto.Columns.ColumnByFieldName("Item").Value = i
                    gVuelto.Dataset.Post
                    
                    gVuelto.Dataset.Next
                Loop
                If gVuelto.Dataset.State = dsEdit Or gVuelto.Dataset.State = dsInsert Then
                    gVuelto.Dataset.Post
                End If
            End If
            calcularTotales
        End If
    End If
End If
If KeyCode = 13 Then
    If gVuelto.Dataset.State = dsEdit Or gVuelto.Dataset.State = dsInsert Then
          gVuelto.Dataset.Post
    End If
End If
End Sub

Private Sub gvuelto_OnKeyPress(Key As Integer)
Dim strCod As String
Dim strDes As String
If Key <> 9 And Key <> 13 And Key <> 27 Then
    Select Case gVuelto.Columns.FocusedColumn.Index
        Case gVuelto.Columns.ColumnByFieldName("idMoneda").Index
            strCod = gVuelto.Columns.ColumnByFieldName("idMoneda").Value
            strDes = gVuelto.Columns.ColumnByFieldName("glsMoneda").Value
            
            mostrarAyudaKeyasciiTexto Key, "MONEDA", strCod, strDes
            Key = 0
            gVuelto.Dataset.Edit
            gVuelto.Columns.ColumnByFieldName("idMoneda").Value = strCod
            gVuelto.Columns.ColumnByFieldName("glsMoneda").Value = strDes
'            If strCod = "PEN" Then 'soles
'                gVuelto.Columns.ColumnByFieldName("MontoSoles").Value = gVuelto.Columns.ColumnByFieldName("MontoOri").Value
'            Else 'dolares
'                gVuelto.Columns.ColumnByFieldName("MontoSoles").Value = gVuelto.Columns.ColumnByFieldName("MontoOri").Value * txt_TipoCambio.Value
'            End If
            gVuelto.Dataset.Post
'            calcularTotales
    End Select
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strMsgError As String
On Error GoTo ERR
Select Case Button.Index
    Case 1 'Grabar
        grabar strMsgError
        If strMsgError <> "" Then GoTo ERR
    Case 2 'Cancelar
        Unload Me
    Case 3 'Salir
        Unload Me
End Select
Exit Sub
ERR:
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub listaPagos(ByRef strMsgError As String)
Dim strCond As String
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim rsv As New ADODB.Recordset
On Error GoTo ERR
'TRAE EL LISTADO DE FORMAS DE PAGO DEL DOCUMENTO Y LO ALMACENA EN UN RECORSET
    csql = "SELECT p.item, p.idFormadePago,f.GlsFormaPago, p.idMoneda,m.GlsMoneda, p.MontoOri, p.MontoSoles, " & _
                  "P.idTipoFormaPago , P.fecVctos " & _
           "FROM pagosdocventas p,formaspagos f,monedas m " & _
           "WHERE p.idEmpresa = '" & glsEmpresa & "' AND p.idSucursal = '" & glsSucursal & "' " & _
             "AND f.idEmpresa = '" & glsEmpresa & "' " & _
             "AND P.idFormadePago = f.idFormaPago " & _
             "AND p.idMoneda = m.idMoneda " & _
             "AND p.idDocumento = '" & strTD & "' AND  p.idDocVentas = '" & strNumDoc & "' AND p.idSerie = '" & strSerie & "' ORDER BY p.item"
             
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idFormadePago", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "glsFormadePago", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idMoneda", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsMoneda", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "MontoOri", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "MontoSoles", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idTipoFormaPago", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "fecVctos", adVarChar, 185, adFldIsNullable
    
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
    
        rsg.Fields("Item") = 1
        rsg.Fields("idFormadePago") = glsFormaPagoVentas
        rsg.Fields("glsFormadePago") = traerCampo("formaspagos", "GlsFormaPago", "idFormaPago", glsFormaPagoVentas, True)
        ''''rsg.Fields("idMoneda") = "PEN"
        rsg.Fields("idMoneda") = txtCod_Moneda.Text
        rsg.Fields("GlsMoneda") = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
        ''''rsg.Fields("MontoOri") = txt_TotalNeto.Value
        ''''rsg.Fields("MontoSoles") = txt_TotalNeto.Value
        rsg.Fields("MontoOri") = txt_TotalNeto.Value
        
        If txtCod_Moneda.Text = "PEN" Then
            rsg.Fields("MontoSoles") = txt_TotalNeto.Value
        Else
            rsg.Fields("MontoSoles") = txt_TotalNeto.Value * txt_TipoCambio.Text
        End If
        
        rsg.Fields("idTipoFormaPago") = traerCampo("formaspagos", "idTipoFormaPago", "idFormaPago", glsFormaPagoVentas, True)
        
        If rsg.Fields("idTipoFormaPago") = "06090002" Then
            rsg.Fields("fecVctos") = Format(DateAdd("D", traerCampo("formaspagos", "diasVcto", "idFormaPago", glsFormaPagoVentas, True), CDate(fecEmision)), "dd/mm/yyyy")
        Else
            rsg.Fields("fecVctos") = ""
        End If
            
        
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = rst.Fields("Item")
            rsg.Fields("idFormadePago") = "" & rst.Fields("idFormadePago")
            rsg.Fields("glsFormadePago") = "" & rst.Fields("GlsFormaPago")
            rsg.Fields("idMoneda") = "" & rst.Fields("idMoneda")
            rsg.Fields("GlsMoneda") = "" & rst.Fields("GlsMoneda")
            rsg.Fields("MontoOri") = "" & rst.Fields("MontoOri")
            rsg.Fields("MontoSoles") = "" & rst.Fields("MontoSoles")
            rsg.Fields("idTipoFormaPago") = "" & rst.Fields("idTipoFormaPago")
            rsg.Fields("fecVctos") = "" & rst.Fields("fecVctos")
            rst.MoveNext
        Loop
    End If
      
    mostrarDatosGridSQL gPagos, rsg, strMsgError
    If strMsgError <> "" Then GoTo ERR
    
    
'TRAE LOS VUELTOS REGISTRADOS
    csql = "SELECT m.idMoneda,o.GlsMoneda, m.ValMonto " & _
           "FROM movcajasdet m,monedas o " & _
           "WHERE m.idMoneda = o.idMoneda " & _
             "AND m.idDocumento = '" & strTD & "' AND  m.idDocVentas = '" & strNumDoc & "' AND m.idSerie = '" & strSerie & "' " & _
             "AND m.idTipoMovCaja = '99990003' ORDER BY m.idMoneda"
             
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
    rsv.Fields.Append "Item", adInteger, , adFldRowID
    rsv.Fields.Append "idMoneda", adVarChar, 8, adFldIsNullable
    rsv.Fields.Append "GlsMoneda", adVarChar, 185, adFldIsNullable
    rsv.Fields.Append "Vuelto", adDouble, 14, adFldIsNullable
        
    rsv.Open
    
    If rst.RecordCount = 0 Then
        rsv.AddNew
    
        rsv.Fields("Item") = 1
        rsv.Fields("idMoneda") = ""
        rsv.Fields("GlsMoneda") = ""
        rsv.Fields("Vuelto") = 0
        
    Else
        Do While Not rst.EOF
            rsv.AddNew
            rsv.Fields("Item") = rst.RecordCount
            rsv.Fields("idMoneda") = "" & rst.Fields("idMoneda")
            rsv.Fields("GlsMoneda") = "" & rst.Fields("GlsMoneda")
            rsv.Fields("Vuelto") = "" & rst.Fields("ValMonto")
            rst.MoveNext
        Loop
    End If
    
    mostrarDatosGridSQL gVuelto, rsv, strMsgError
    If strMsgError <> "" Then GoTo ERR
    
    
    calcularTotales
    
    
Me.Refresh
If rst.State = 1 Then rst.Close
Set rst = Nothing
Exit Sub
ERR:
If rst.State = 1 Then rst.Close
Set rst = Nothing
If strMsgError = "" Then strMsgError = ERR.Description

End Sub

Private Sub txtCod_Caja_Change()
    txtGls_Caja.Text = traerCampo("cajas", "GlsCaja", "idCaja", txtCod_Caja.Text, True)
End Sub

Private Sub txtCod_Moneda_Change()
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
End Sub

Private Sub calcularTotales()

Dim intFila As Integer
Dim dblTotalNeto As Double
    
    intFila = gPagos.Dataset.RecNo
    intFila = gPagos.Dataset.RecNo
    intFila = gPagos.Dataset.RecNo
    
    txt_TotalRecibido.Text = 0#
    txt_Vuelto.Text = 0#
    gPagos.Dataset.First
    Do While Not gPagos.Dataset.EOF
        txt_TotalRecibido.Text = txt_TotalRecibido.Value + gPagos.Columns.ColumnByFieldName("MontoSoles").Value
        gPagos.Dataset.Next
    Loop
    
    gPagos.Dataset.RecNo = intFila
    
    If txtCod_Moneda.Text = "PEN" Then
        dblTotalNeto = Val(txt_TotalNeto.Value)
    Else
        dblTotalNeto = Val(txt_TotalNeto.Value) * Val(txt_TipoCambio.Value)
    End If
    
    If Val(txt_TotalRecibido.Value) > dblTotalNeto Then
        txt_Vuelto.Text = Val(txt_TotalRecibido.Value) - dblTotalNeto
    End If
    
    gVuelto.Dataset.Edit
    If txt_Vuelto.Value > 0 Then
        fraVuelto.Enabled = True
        
'        gVuelto.Dataset.Edit
        gVuelto.Columns.ColumnByFieldName("idMoneda").Value = "PEN"
        gVuelto.Columns.ColumnByFieldName("GlsMoneda").Value = traerCampo("monedas", "GlsMoneda", "idMoneda", "PEN", False)
        gVuelto.Columns.ColumnByFieldName("Vuelto").Value = txt_Vuelto.Value
'        gVuelto.Dataset.Post
        
    Else
        
        
'        gVuelto.Dataset.Edit
        gVuelto.Columns.ColumnByFieldName("idMoneda").Value = ""
        gVuelto.Columns.ColumnByFieldName("GlsMoneda").Value = ""
        gVuelto.Columns.ColumnByFieldName("Vuelto").Value = 0#
'        gVuelto.Dataset.Post
        
        fraVuelto.Enabled = False
    End If
    
    gVuelto.Dataset.Post
    
    calcularTotalVueltoEntregado
End Sub

Private Sub calcularTotalVueltoEntregado()

Dim dblVuelto As Double

Dim intFila As Integer
    
    intFila = gVuelto.Dataset.RecNo
    intFila = gVuelto.Dataset.RecNo
    intFila = gVuelto.Dataset.RecNo

    txtVal_VueltoEntregado.Text = 0#
    
    gVuelto.Dataset.First
    Do While Not gVuelto.Dataset.EOF
    
        If gVuelto.Columns.ColumnByFieldName("idMoneda").Value <> "" Then
            If gVuelto.Columns.ColumnByFieldName("idMoneda").Value = "PEN" Then
                dblVuelto = dblVuelto + Val("" & gVuelto.Columns.ColumnByFieldName("Vuelto").Value)
            Else
                dblVuelto = dblVuelto + (Val("" & gVuelto.Columns.ColumnByFieldName("Vuelto").Value) * txt_TipoCambio.Value)
            End If
        End If
        gVuelto.Dataset.Next
    Loop
    
    txtVal_VueltoEntregado.Text = dblVuelto
    
    gVuelto.Dataset.RecNo = intFila
    
End Sub

Private Sub eliminaNulosGrilla()
    Dim indWhile As Boolean
    Dim indEntro As Boolean
    Dim i As Integer
    indWhile = True
    Do While indWhile = True
        If gPagos.Count >= 1 Then
            gPagos.Dataset.First
            indEntro = False
            Do While Not gPagos.Dataset.EOF
                If Trim(gPagos.Columns.ColumnByFieldName("idFormadePago").Value) = "" Then
                    gPagos.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gPagos.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gPagos.Count >= 1 Then
        gPagos.Dataset.First
        i = 0
        Do While Not gPagos.Dataset.EOF
            i = i + 1
            gPagos.Dataset.Edit
            gPagos.Columns.ColumnByFieldName("item").Value = i
            If gPagos.Dataset.State = dsEdit Then gPagos.Dataset.Post
            gPagos.Dataset.Next
        Loop
    Else
        indInserta = True
        gPagos.Dataset.Append
        indInserta = False
    End If
    
End Sub

Public Sub mostrarForm(ByVal strVarTipoDoc As String, ByVal strVarNumDoc As String, ByVal strVarSerie As String, ByRef strMsgError As String)
Dim rst As New ADODB.Recordset
Dim strCodCaja As String
Dim strCodMovCaja As String
On Error GoTo ERR

strTD = strVarTipoDoc
strNumDoc = strVarNumDoc
strSerie = strVarSerie

'''strCodMovCaja = Trim(traerCampo("docventas", "idMovCaja", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))
'''If strCodMovCaja = "" Then
    txtCod_Caja.Text = CajaAperturadaUsuario(1, strMsgError)
    If strMsgError <> "" Then GoTo ERR
'''Else
'''    strCodMovCaja = traerCampo("movcajas", "idCaja", "idMovCaja", strCodMovCaja, True, "idSucursal = '" & glsSucursal & "'")
'''    txtCod_Caja.Text = strCodMovCaja
'''End If

txt_Serie.Text = strSerie
txt_NumDoc.Text = strNumDoc
lblDoc.Caption = traerCampo("documentos", "GlsDocumento", "idDocumento", strTD, False)

csql = "SELECT d.idPerCliente,d.idMoneda, d.TipoCambio, d.TotalValorVenta, d.TotalIGVVenta, d.TotalPrecioVenta, d.estDocVentas, " & _
       "d.FecPago, d.RUCCliente, d.dirCliente, d.GlsCliente, " & _
       "(SELECT SUM(c.ValMonto) FROM movcajasdet c WHERE c.idEmpresa = '" & glsEmpresa & "' AND c.idSucursal = '" & glsSucursal & "' AND c.idDocumento = '" & strTD & "' AND c.idSerie = d.idSerie AND c.idDocVentas = d.idDocVentas) AS Adelantos " & _
       "FROM docventas d " & _
       "WHERE d.idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND d.idDocumento = '" & strTD & "' AND d.idDocVentas = '" & strNumDoc & "' AND d.idSerie = '" & strSerie & "'"

rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly

If Not rst.EOF Then

    strCodCliente = "" & rst.Fields("idPerCliente")
    
    txtCod_Moneda.Text = "" & rst.Fields("idMoneda")
    txt_TipoCambio.Text = "" & rst.Fields("TipoCambio")
    strEstDoc = "" & rst.Fields("estDocVentas")
'    If rst.Fields("idMoneda") = "PEN" Then 'SOLES
        txt_TotalBruto.Text = "" & rst.Fields("TotalValorVenta")
        txt_TotalIGV.Text = "" & rst.Fields("TotalIGVVenta")
        txt_TotalNeto.Text = "" & rst.Fields("TotalPrecioVenta")
        
        txt_Adelantos.Text = "" & rst.Fields("Adelantos")
        txt_Saldo.Text = Val("" & rst.Fields("TotalPrecioVenta")) - Val("" & rst.Fields("Adelantos"))
'    Else
'        txt_TotalBruto.Text = Val("" & rst.Fields("TotalValorVenta")) * txt_TipoCambio.Value
'        txt_TotalIGV.Text = Val("" & rst.Fields("TotalIGVVenta")) * txt_TipoCambio.Value
'        txt_TotalNeto.Text = Val("" & rst.Fields("TotalPrecioVenta")) * txt_TipoCambio.Value
'    End If
    
'    If strEstDoc = "IMP" Then
'        frmPagos.Enabled = False
'        Toolbar1.Buttons(1).Visible = False
'    Else
        frmPagos.Enabled = True
        Toolbar1.Buttons(1).Visible = True
'    End If
    
End If

nuevo strMsgError
If strMsgError <> "" Then GoTo ERR

frmPagosSeparacion.Show 1

Unload frmPagosSeparacion

If rst.State = 1 Then rst.Close
Set rst = Nothing
Exit Sub
ERR:
If rst.State = 1 Then rst.Close
Set rst = Nothing
If strMsgError = "" Then strMsgError = ERR.Description
End Sub

Public Sub EjecutaSQLFormPagosSeparacion(F As Form, ByRef strMsgError As String, strTD As String, strNumDoc As String, strSerie As String, strCodMovCaja As String, g As dxDBGrid, v As dxDBGrid)
Dim csql As String
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strCodMovDet As String

Dim strCodMoneda As String
Dim strCodFormadePago As String
Dim dblValMonto As Double

Dim strFormasPago As String
Dim strFecVactos  As String

Dim indTrans As Boolean

Dim intItem As Integer

On Error GoTo ERR

indTrans = True
Cn.BeginTrans

'Grabando Grilla
If TypeName(g) <> "Nothing" Then
'''    'Eliminamos pagosdocventas
'''    Cn.Execute "DELETE FROM pagosdocventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
'''
'''    'Eliminamos movcajasdet
'''    Cn.Execute "DELETE FROM movcajasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    
    g.Dataset.First
    Do While Not g.Dataset.EOF
        strCampos = ""
        strValores = ""
        For i = 0 To g.Columns.Count - 1
                If UCase(left(g.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(g.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(g.Columns(i).ObjectName, 3)
                    
                    strCampos = strCampos & strCampo & ","
                    
                    If strCampo = "idMoneda" Then
                        strCodMoneda = Trim(g.Columns(i).Value)
                    End If
                    
                    If strCampo = "idFormadePago" Then
                        strCodFormadePago = Trim(g.Columns(i).Value)
                    End If
                    
                    If strCampo = "MontoOri" Then
                        dblValMonto = g.Columns(i).Value
                    End If
                    
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
        
        
        'Buscamos en q item se quedo
        intItem = Val(traerCampo("pagosdocventas", "MAX(Item)", "idSucursal", glsSucursal, True, " idDocumento = '" & strTD & "' AND  idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'") & "")
        
        'Continuamos con la grabacion
        intItem = intItem + 1
        
        csql = "INSERT INTO pagosdocventas(" & strCampos & ",idDocumento,idDocVentas,idSerie,idEmpresa,idSucursal,item) VALUES(" & strValores & ",'" & strTD & "','" & strNumDoc & "','" & strSerie & "','" & glsEmpresa & "','" & glsSucursal & "'," & CStr(intItem) & ")"
                
        Cn.Execute csql
        
        'insertamos los montos en movimientos de caja
        strCodMovDet = generaCorrelativoAnoMes("movcajasdet", "idMovCajaDet")
        
        csql = "INSERT INTO movcajasdet (idMovCajaDet,idMovCaja,idTipoMovCaja,idMoneda,ValMonto,FecRegistro,idEmpresa,idSucursal,ValTipoCambio,idFormadePago,idDocumento,idDocVentas,idSerie) VALUES(" & _
               "'" & strCodMovDet & "','" & strCodMovCaja & "','99990002','" & strCodMoneda & "'," & dblValMonto & ",sysdate(),'" & glsEmpresa & "','" & glsSucursal & "'," & glsTC & ",'" & strCodFormadePago & "','" & strTD & "','" & strNumDoc & "','" & strSerie & "')"
               
        Cn.Execute csql
        
        g.Dataset.Next
    Loop
End If

If TypeName(v) <> "Nothing" Then

    'Eliminamos movcajasdet --- ya se elmino todos los movimiento arriba
    'Cn.Execute "DELETE FROM movcajasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    
    v.Dataset.First
    Do While Not v.Dataset.EOF
    
        strCodMoneda = v.Columns.ColumnByFieldName("idMoneda").Value
        dblValMonto = v.Columns.ColumnByFieldName("Vuelto").Value
                   
        If dblValMonto > 0 Then
            'insertamos los montos en movimientos de caja
            strCodMovDet = generaCorrelativoAnoMes("movcajasdet", "idMovCajaDet")
        
            csql = "INSERT INTO movcajasdet (idMovCajaDet,idMovCaja,idTipoMovCaja,idMoneda,ValMonto,FecRegistro,idEmpresa,idSucursal,ValTipoCambio,idFormadePago,idDocumento,idDocVentas,idSerie) VALUES(" & _
                   "'" & strCodMovDet & "','" & strCodMovCaja & "','99990003','" & strCodMoneda & "'," & dblValMonto & ",sysdate(),'" & glsEmpresa & "','" & glsSucursal & "'," & glsTC & ",'','" & strTD & "','" & strNumDoc & "','" & strSerie & "')"
                   
            Cn.Execute csql
        
        End If
        
        v.Dataset.Next
    Loop
End If

'Actualizamos el Documento de venta
'''g.Dataset.First
'''Do While Not g.Dataset.EOF
'''    If Trim(g.Columns.ColumnByFieldName("glsFormadePago").Value) <> "" Then
'''        strFormasPago = strFormasPago + g.Columns.ColumnByFieldName("glsFormadePago").Value & ","
'''    End If
'''    If Trim(g.Columns.ColumnByFieldName("fecVctos").Value) <> "" Then
'''        strFecVactos = strFecVactos + g.Columns.ColumnByFieldName("fecVctos").Value & ","
'''    End If
'''    g.Dataset.Next
'''Loop
'''
'''If Len(Trim(strFormasPago)) > 1 Then
'''    strFormasPago = left(strFormasPago, Len(strFormasPago) - 1)
'''End If
'''
'''If Len(Trim(strFecVactos)) > 1 Then
'''    strFecVactos = left(strFecVactos, Len(strFecVactos) - 1)
'''End If

'''csql = "UPDATE DOCVENTAS SET GlsFormasPago = '" & strFormasPago & "', GlsFecVectos = '" & strFecVactos & "',estDocVentas = 'CAN', idMovCaja = '" & strCodMovCaja & "' " & _
'''       "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND  idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"

'''Cn.Execute csql
    
    
Cn.CommitTrans
Exit Sub
ERR:
If strMsgError = "" Then strMsgError = ERR.Description
If indTrans Then Cn.RollbackTrans
'''Resume
End Sub
