VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmVale 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vale"
   ClientHeight    =   9210
   ClientLeft      =   3660
   ClientTop       =   3810
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13215
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   8475
      Left            =   90
      TabIndex        =   25
      Top             =   675
      Width           =   13065
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   120
         TabIndex        =   26
         Top             =   150
         Width           =   12840
         Begin VB.ComboBox cbx_Mes 
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
            ItemData        =   "frmVale.frx":0000
            Left            =   9225
            List            =   "frmVale.frx":0028
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   225
            Width           =   1845
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   915
            TabIndex        =   0
            Top             =   210
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
            Container       =   "frmVale.frx":0091
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Ano 
            Height          =   315
            Left            =   11955
            TabIndex        =   2
            Top             =   225
            Width           =   765
            _ExtentX        =   1349
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
            Container       =   "frmVale.frx":00AD
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Mes"
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
            Left            =   8820
            TabIndex        =   37
            Top             =   300
            Width           =   300
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Año"
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
            Left            =   11550
            TabIndex        =   36
            Top             =   300
            Width           =   300
         End
         Begin VB.Label Label21 
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
            TabIndex        =   27
            Top             =   255
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3960
         Left            =   135
         OleObjectBlob   =   "frmVale.frx":00C9
         TabIndex        =   34
         Top             =   945
         Width           =   12840
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   3330
         Left            =   120
         OleObjectBlob   =   "frmVale.frx":2F94
         TabIndex        =   35
         Top             =   4995
         Width           =   12840
      End
   End
   Begin VB.Frame fraTotales 
      Appearance      =   0  'Flat
      Caption         =   " Totales "
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
      Height          =   690
      Left            =   75
      TabIndex        =   41
      Top             =   8430
      Width           =   13065
      Begin CATControls.CATTextBox txt_TotalBruto 
         Height          =   315
         Left            =   3825
         TabIndex        =   42
         Tag             =   "NvalorTotal"
         Top             =   225
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         BackColor       =   12640511
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmVale.frx":6BC2
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalIGV 
         Height          =   315
         Left            =   6450
         TabIndex        =   43
         Tag             =   "NigvTotal"
         Top             =   225
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         BackColor       =   12640511
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmVale.frx":6BDE
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalNeto 
         Height          =   315
         Left            =   9150
         TabIndex        =   44
         Tag             =   "NprecioTotal"
         Top             =   225
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         BackColor       =   12640511
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmVale.frx":6BFA
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_DocReferencia 
         Height          =   285
         Left            =   150
         TabIndex        =   45
         Tag             =   "TGlsDocReferencia"
         Top             =   225
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   8438015
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
         Locked          =   -1  'True
         Container       =   "frmVale.frx":6C16
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtTotalPeso 
         Height          =   315
         Left            =   11280
         TabIndex        =   59
         Top             =   225
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         BackColor       =   12640511
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmVale.frx":6C32
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.Label LblPeso 
         Appearance      =   0  'Flat
         Caption         =   "Peso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   10800
         TabIndex        =   60
         Top             =   270
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lbl_SimbMonBruto 
         Appearance      =   0  'Flat
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3345
         TabIndex        =   51
         Top             =   270
         Width           =   240
      End
      Begin VB.Label lbl_SimbMonIGV 
         Appearance      =   0  'Flat
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6120
         TabIndex        =   50
         Top             =   270
         Width           =   240
      End
      Begin VB.Label lbl_SimbMonNeto 
         Appearance      =   0  'Flat
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   8850
         TabIndex        =   49
         Top             =   270
         Width           =   240
      End
      Begin VB.Label lbl_TotalNeto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   8340
         TabIndex        =   48
         Top             =   270
         Width           =   345
      End
      Begin VB.Label lbl_TotalIGV 
         Appearance      =   0  'Flat
         Caption         =   "IGV:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5745
         TabIndex        =   47
         Top             =   270
         Width           =   285
      End
      Begin VB.Label lbl_TotalBruto 
         Appearance      =   0  'Flat
         Caption         =   "Bruto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2820
         TabIndex        =   46
         Top             =   270
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   8235
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
            Picture         =   "frmVale.frx":6C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":6FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":743A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":77D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":7B6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":7F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":82A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":863C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":89D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":8D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":910A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":9DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":A166
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":A480
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmbImprimirBloque 
      Caption         =   "&Imprimir en Bloque"
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
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1575
      Visible         =   0   'False
      Width           =   2025
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   1164
      ButtonWidth     =   2540
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Nuevo         "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Grabar    "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Modificar    "
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Cancelar    "
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Anular    "
            Object.ToolTipText     =   "Anular"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Eliminar    "
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Importar    "
            Object.ToolTipText     =   "Importar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Imprimir    "
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Lista    "
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar Excel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Excel    "
            Object.ToolTipText     =   "Excel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Datos Tránsito"
            Object.ToolTipText     =   "Datos Tránsito"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Salir    "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   90
      TabIndex        =   14
      Top             =   690
      Width           =   13065
      Begin CATControls.CATTextBox txtGls_usuario 
         Height          =   315
         Left            =   1305
         TabIndex        =   62
         Top             =   2325
         Width           =   3660
         _ExtentX        =   6456
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
         Container       =   "frmVale.frx":A58A
         Vacio           =   -1  'True
      End
      Begin VB.CommandButton cmbAyudaCentroCosto 
         Height          =   315
         Left            =   7365
         Picture         =   "frmVale.frx":A5A6
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1980
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaCliente 
         Height          =   315
         Left            =   7365
         Picture         =   "frmVale.frx":A930
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1245
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaAlmacen 
         Height          =   315
         Left            =   7365
         Picture         =   "frmVale.frx":ACBA
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   915
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaConcepto 
         Height          =   315
         Left            =   7365
         Picture         =   "frmVale.frx":B044
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   540
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   7365
         Picture         =   "frmVale.frx":B3CE
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1620
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Vale 
         Height          =   315
         Left            =   11835
         TabIndex        =   17
         Tag             =   "TidValesCab"
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmVale.frx":B758
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Concepto 
         Height          =   315
         Left            =   1305
         TabIndex        =   3
         Tag             =   "TidConcepto"
         Top             =   540
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
         MaxLength       =   8
         Container       =   "frmVale.frx":B774
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Concepto 
         Height          =   315
         Left            =   2250
         TabIndex        =   18
         Top             =   540
         Width           =   5100
         _ExtentX        =   8996
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
         Container       =   "frmVale.frx":B790
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1305
         TabIndex        =   6
         Tag             =   "TidMoneda"
         Top             =   1605
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
         MaxLength       =   8
         Container       =   "frmVale.frx":B7AC
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2250
         TabIndex        =   19
         Top             =   1605
         Width           =   5100
         _ExtentX        =   8996
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
         Container       =   "frmVale.frx":B7C8
      End
      Begin CATControls.CATTextBox txtCod_Almacen 
         Height          =   315
         Left            =   1305
         TabIndex        =   4
         Tag             =   "TidAlmacen"
         Top             =   870
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
         MaxLength       =   8
         Container       =   "frmVale.frx":B7E4
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Almacen 
         Height          =   315
         Left            =   2250
         TabIndex        =   29
         Top             =   870
         Width           =   5100
         _ExtentX        =   8996
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
         Container       =   "frmVale.frx":B800
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1305
         TabIndex        =   5
         Tag             =   "TidProvCliente"
         Top             =   1230
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
         MaxLength       =   20
         Container       =   "frmVale.frx":B81C
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2250
         TabIndex        =   31
         Top             =   1230
         Width           =   5100
         _ExtentX        =   8996
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
         Container       =   "frmVale.frx":B838
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtObs 
         Height          =   315
         Left            =   1305
         TabIndex        =   8
         Tag             =   "TobsValesCab"
         Top             =   1965
         Width           =   6045
         _ExtentX        =   10663
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
         MaxLength       =   500
         Container       =   "frmVale.frx":B854
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtp_Emision 
         Height          =   315
         Left            =   8865
         TabIndex        =   9
         Tag             =   "FfechaEmision"
         Top             =   540
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   131792897
         CurrentDate     =   38955
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gDocReferencia 
         Height          =   1335
         Left            =   8055
         OleObjectBlob   =   "frmVale.frx":B870
         TabIndex        =   11
         Top             =   1350
         Width           =   4770
      End
      Begin CATControls.CATTextBox txt_TipoCambio 
         Height          =   315
         Left            =   8865
         TabIndex        =   10
         Tag             =   "NTipoCambio"
         Top             =   870
         Width           =   1185
         _ExtentX        =   2090
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
         Container       =   "frmVale.frx":DF14
         Text            =   "0.000"
         Decimales       =   3
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_CentroCosto 
         Height          =   315
         Left            =   1305
         TabIndex        =   7
         Tag             =   "TidCentroCosto"
         Top             =   1980
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
         MaxLength       =   8
         Container       =   "frmVale.frx":DF30
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_CentroCosto 
         Height          =   315
         Left            =   2250
         TabIndex        =   54
         Top             =   1980
         Visible         =   0   'False
         Width           =   5100
         _ExtentX        =   8996
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
         Container       =   "frmVale.frx":DF4C
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_AlmacenTrans 
         Height          =   315
         Left            =   1305
         TabIndex        =   56
         Top             =   2700
         Visible         =   0   'False
         Width           =   930
         _ExtentX        =   1640
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
         Container       =   "frmVale.frx":DF68
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_AlmacenTrans 
         Height          =   315
         Left            =   2250
         TabIndex        =   57
         Top             =   2700
         Visible         =   0   'False
         Width           =   5100
         _ExtentX        =   8996
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
         Container       =   "frmVale.frx":DF84
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_usuario 
         Height          =   315
         Left            =   9360
         TabIndex        =   61
         Tag             =   "TidUsuarioRegistro"
         Top             =   2850
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
         Container       =   "frmVale.frx":DFA0
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label lbl_Vendedor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   150
         TabIndex        =   63
         Top             =   2400
         Width           =   945
      End
      Begin VB.Label Lbl_AlmacenTrans 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén Dest."
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
         Left            =   135
         TabIndex        =   58
         Top             =   2790
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lblCentroCosto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "C. Costo"
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
         Left            =   150
         TabIndex        =   55
         Top             =   2070
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbl_Anulado 
         Appearance      =   0  'Flat
         Caption         =   "ANULADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   105
         TabIndex        =   39
         Top             =   180
         Width           =   2895
      End
      Begin VB.Label lbl_TC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "T/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   8055
         TabIndex        =   38
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   150
         TabIndex        =   33
         Top             =   2010
         Width           =   1200
      End
      Begin VB.Label lbl_FechaEmision 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   8055
         TabIndex        =   32
         Top             =   570
         Width           =   450
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
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
         Left            =   180
         TabIndex        =   24
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label lblProvClie 
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
         Left            =   150
         TabIndex        =   23
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Left            =   150
         TabIndex        =   22
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nº Vale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   11115
         TabIndex        =   21
         Top             =   210
         Width           =   570
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
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
         Left            =   150
         TabIndex        =   20
         Top             =   570
         Width           =   690
      End
   End
   Begin VB.Frame fraDetalle 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   4425
      Left            =   90
      TabIndex        =   52
      Top             =   3975
      Width           =   13065
      Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
         Height          =   4080
         Left            =   90
         OleObjectBlob   =   "frmVale.frx":DFBC
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   12900
      End
   End
End
Attribute VB_Name = "frmVale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsDetalle                                   As New ADODB.Recordset
Public indVale                                  As String
Private strEstVale                              As String
Private indInserta                              As Boolean
Private indNuevoDoc                             As Boolean
Private indCargando                             As Boolean
Private indInsertaDocRef                        As Boolean
Dim dblIgvNEw                                   As Double
Dim Sw_Documento                                As Boolean
Dim RsDetAtributos                              As New ADODB.Recordset
Dim CIdAlmacenAnt                               As String

Private Sub cbx_Mes_Click()
On Error GoTo Err
Dim StrMsgError As String

    If indNuevoDoc = False Then
        listaVales StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaAlmacen_Click()
    
    mostrarAyuda "ALMACEN", txtCod_Almacen, txtGls_Almacen
    'If txtCod_Almacen.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaCentroCosto_Click()
    
    mostrarAyuda "CENTROCOSTO", txtCod_CentroCosto, txtGls_CentroCosto

End Sub

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "PERSONA", txtCod_Cliente, txtGls_Cliente
    'If indVale = "I" Then
    '    mostrarAyuda "PROVEEDOR", txtCod_Cliente, txtGls_Cliente
    'Else
    '    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente
    'End If
    'If txtCod_Cliente.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaConcepto_Click()
    
    If indVale = "I" Then
        mostrarAyuda "CONCEPTOINGRESO", TxtCod_Concepto, TxtGls_Concepto
    Else
        mostrarAyuda "CONCEPTOSALIDA", TxtCod_Concepto, TxtGls_Concepto
    End If
    
    'If Trim("" & traerCampo("conceptos", "indCosto", "idConcepto", txtCod_Concepto.Text, False, " tipoConcepto = '" & indVale & "' ")) = "N" Then
    '    gDetalle.Columns.ColumnByFieldName("VVUnit").Visible = False
    '    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Visible = False
    'Else
    '    gDetalle.Columns.ColumnByFieldName("VVUnit").Visible = True
    '    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Visible = True
    'End If
    
    If Trim("" & traerCampo("conceptos", "IndDevolucion", "idConcepto", TxtCod_Concepto.Text, False, " tipoConcepto = '" & indVale & "' ")) = "1" Then
    
        gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = True
        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").DisableEditor = True
        gDetalle.Columns.ColumnByFieldName("PVUnit").DisableEditor = True
        gDetalle.Columns.ColumnByFieldName("TotalPVNeto").DisableEditor = True
    
    Else
    
        gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = False
        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").DisableEditor = False
        gDetalle.Columns.ColumnByFieldName("PVUnit").DisableEditor = False
        gDetalle.Columns.ColumnByFieldName("TotalPVNeto").DisableEditor = False
    
    End If
    
    If Trim("" & traerCampo("conceptos", "IndDevolucion", "idConcepto", TxtCod_Concepto.Text, False, " tipoConcepto = '" & indVale & "' ")) = "1" Then
    
        gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = True
        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").DisableEditor = True
        gDetalle.Columns.ColumnByFieldName("PVUnit").DisableEditor = True
        gDetalle.Columns.ColumnByFieldName("TotalPVNeto").DisableEditor = True
    
    Else
    
        gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = False
        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").DisableEditor = False
        gDetalle.Columns.ColumnByFieldName("PVUnit").DisableEditor = False
        gDetalle.Columns.ColumnByFieldName("TotalPVNeto").DisableEditor = False
    
    End If
    
    'If txtCod_Concepto.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbImprimirBloque_Click()
On Error GoTo Err
Dim StrMsgError As String

    If txtCod_Vale.Text <> "" Then
'        ImprimeCodigoBarra 1, "", txtCod_Vale.Text, StrMsgError
'        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub dtp_Emision_Change()
Dim strPeriodo      As String

    strPeriodo = Format(dtp_Emision.Value, "yyyymm")
    If Trim("" & traerCampo("parametros", "valparametro", "glsparametro", "PERIODO_CAMBIO_IGV", True)) > strPeriodo Then
        dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV_ANT", True), "0.00")) / 100, "0.00")
    Else
        dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV", True), "0.00")) / 100, "0.00")
    End If
    If txtCod_Vale.Text = "" Then
        txt_TipoCambio.Text = Val(Format(traerCampo("tiposdecambio", "tcVenta", "day(fecha)", Day(dtp_Emision.Value), False, " month(fecha)= " & Month(dtp_Emision.Value) & " and year(fecha)= " & Year(dtp_Emision.Value) & " "), "0.000"))
    End If
    
End Sub

Private Sub dtp_Emision_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then SendKeys "{tab}"

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError                             As String
Dim intDecimales                            As Integer

    If indVale = "I" Then
        Me.Caption = "Vale de Ingreso"
        'lblProvClie.Caption = "Proveedor"
    Else
        Me.Caption = "Vale de Salida"
        'lblProvClie.Caption = "Cliente"
    End If
    
    Me.top = 0
    Me.left = 0
    
    txt_Ano.Text = Year(getFechaSistema)
    cbx_Mes.ListIndex = Month(getFechaSistema) - 1
    
    ConfGrid_Inv gLista, False, False, False, False, True
    ConfGrid_Inv gListaDetalle, False, False, False, False
    
    ConfGrid_Inv gDetalle, True, False, False, False, False
    ConfGrid_Inv gDocReferencia, True, False, False, True, False
    
    intDecimales = leeParametro("DECIMALESVALES")
    
    txt_TipoCambio.Decimales = glsDecimalesTC
    txtCod_Moneda.Text = "PEN"
    
    muestraColumnasDetalle
        
    listaVales StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    fraDetalle.Visible = False
    fraTotales.Visible = False
    
    Toolbar1.Buttons(12).Visible = False
    
    habilitaBotones 9
    
    If gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Visible Then
        LblPeso.Visible = True
        TxtTotalPeso.Visible = True
    End If
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo   As String
Dim strMsg      As String
Dim objVales    As New clsVales
Dim sw_compra   As Integer
Dim StrFechaDoc As String
Dim rsValida    As New ADODB.Recordset
Dim StrCadSql   As String

    StrFechaDoc = ""
    
    getEstadoCierreMes CVDate(dtp_Emision.Value), StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If leeParametro("STOCK_POR_LOTE") = "S" Then
        Do While Not gDetalle.Dataset.EOF
            If Trim(gDetalle.Columns.ColumnByFieldName("idLote").Value) = "" Or Trim(gDetalle.Columns.ColumnByFieldName("idProducto").Value) = "" Or Val(gDetalle.Columns.ColumnByFieldName("Cantidad").Value) = 0 Then
                StrMsgError = "Falta Ingresar datos en el detalle, Verifique."
                GoTo Err
            End If
            gDetalle.Dataset.Next
        Loop
    End If
    eliminaNulosGrilla
    
    If gDetalle.Count >= 1 Then
        If gDetalle.Count = 1 And (gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" Or gDetalle.Columns.ColumnByFieldName("Cantidad").Value <= 0) Then
            StrMsgError = "Falta Ingresar datos en el Detalle"
            GoTo Err
        End If
    End If

    '-------------------------------------------
    Sw_Documento = False
    
    eliminaNulosGrillaDocRef
    generaSTRDocReferencia
    
    sw_compra = traerCampo("conceptos", "indcompra", "idConcepto", TxtCod_Concepto.Text, False)
    If sw_compra = 1 Then
        If Sw_Documento = False Then
            StrMsgError = "Debe Registrar como Documento de Referencia una Factura o Guía de Remisión."
            GoTo Err
        End If
    End If
    '-------------------------------------------
    
    If txtCod_Vale.Text = "" Then
    
        StrFechaDoc = Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "IND_DOC_FECHA_ACTUAL", True))
        If Trim("" & StrFechaDoc) = "S" Then
            
            If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
            StrCadSql = "select sysdate() Fecha  "
            rsValida.Open StrCadSql, Cn, adOpenStatic, adLockOptimistic
            
            If Not rsValida.EOF Then
                If Format(rsValida.Fields("Fecha"), "dd/mm/yyyy") <> Format(dtp_Emision.Value, "dd/mm/yyyy") Then
                    StrMsgError = "Solo está permitido registros con la fecha actual"
                    If StrMsgError <> "" Then GoTo Err
                End If
            End If
            
        End If
    
        EjecutaSQLFormVales Me, 0, StrMsgError, indVale, gDetalle, gDocReferencia, dtp_Emision.Value
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabó"
        
    Else '--- Modifica
        EjecutaSQLFormVales Me, 1, StrMsgError, indVale, gDetalle, gDocReferencia, dtp_Emision.Value
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modificó"
    End If
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    fraGeneral.Enabled = False
    habilitarColumas False
    fraTotales.Enabled = False
    
    listaVales StrMsgError
    If StrMsgError <> "" Then GoTo Err
     
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    '************** Luis 09/04/2019 *******************
    If InStr(1, StrMsgError, "Duplicate", vbTextCompare) > 0 Then
        txtCod_Vale.Text = ""
    End If
    '***************************************************
    
    Exit Sub
End Sub

Private Sub nuevo(StrMsgError As String)
On Error GoTo Err
Dim rsg                                 As New ADODB.Recordset
Dim RsD                                 As New ADODB.Recordset
Dim strAno                              As String
    
    Sw_Documento = False
    limpiaForm Me
    
    CIdAlmacenAnt = ""
    
    strEstVale = "GEN"
    txt_Ano.Text = Year(getFechaSistema)
    txt_TipoCambio.Text = glsTC
    txtCod_Moneda.Text = "PEN"
    lbl_Anulado.Caption = ""
    
    fraGeneral.Enabled = True
    
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adChar, 8, adFldIsNullable
    RsD.Open
    
    RsD.AddNew
    RsD.Fields("Item") = 1
    RsD.Fields("idDocumento") = ""
    RsD.Fields("GlsDocumento") = ""
    RsD.Fields("idSerie") = ""
    RsD.Fields("idNumDOc") = ""
    
    Set gDocReferencia.DataSource = Nothing
    mostrarDatosGridSQL gDocReferencia, RsD, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    gDocReferencia.Columns.FocusedIndex = gDocReferencia.Columns.ColumnByFieldName("idDocumento").ColIndex
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adChar, 15, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Cantidad2", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 45, adFldIsNullable
    rsg.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idSucursalOrigen", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "idDocumentoImp", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "idDocVentasImp", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerieImp", adVarChar, 4, adFldIsNullable
    rsg.Fields.Append "idLote", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "IdTallaPeso", adVarChar, 30, adFldIsNullable
    rsg.Open
    
    rsg.AddNew
    rsg.Fields("Item") = 1
    rsg.Fields("idProducto") = ""
    rsg.Fields("GlsProducto") = ""
    rsg.Fields("idUM") = ""
    rsg.Fields("GlsUM") = ""
    rsg.Fields("Factor") = 1
    rsg.Fields("Afecto") = 1
    rsg.Fields("Cantidad") = 0
    rsg.Fields("Cantidad2") = 0
    rsg.Fields("VVUnit") = 0
    rsg.Fields("IGVUnit") = 0
    rsg.Fields("PVUnit") = 0
    rsg.Fields("TotalVVNeto") = 0
    rsg.Fields("TotalIGVNeto") = 0
    rsg.Fields("TotalPVNeto") = 0
    rsg.Fields("NumLote") = ""
    rsg.Fields("FecVencProd") = ""
    rsg.Fields("idSucursalOrigen") = ""
    rsg.Fields("idDocumentoImp") = ""
    rsg.Fields("idDocVentasImp") = ""
    rsg.Fields("idSerieImp") = ""
    rsg.Fields("idLote") = ""
    rsg.Fields("IdTallaPeso") = ""
    
    Set gDetalle.DataSource = Nothing
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("idProducto").ColIndex
    
    Inicia_RecordSet RsDetAtributos, 0, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

 
Private Sub gdetalle_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    If Action = daInsert Then
        gDetalle.Columns.ColumnByFieldName("item").Value = gDetalle.Count
        gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
        gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
        gDetalle.Columns.ColumnByFieldName("idUM").Value = ""
        gDetalle.Columns.ColumnByFieldName("GlsUM").Value = ""
        gDetalle.Columns.ColumnByFieldName("Factor").Value = 1
        gDetalle.Columns.ColumnByFieldName("Afecto").Value = 1
        gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
        gDetalle.Columns.ColumnByFieldName("Cantidad2").Value = 0
        gDetalle.Columns.ColumnByFieldName("VVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("PVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("NumLote").Value = ""
        gDetalle.Columns.ColumnByFieldName("FecVencProd").Value = ""
        gDetalle.Columns.ColumnByFieldName("idLote").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Value = ""
        gDetalle.Dataset.Post
    End If

End Sub

Private Sub gdetalle_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If Action = daInsert Then
        If gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" And indInserta = False Then
            Allow = False
        Else
            gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("idProducto").ColIndex
        End If
    End If

End Sub

Private Sub gdetalle_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError                     As String
Dim rscd                            As New ADODB.Recordset
Dim strCod                          As String
Dim strDes                          As String
Dim strMoneda                       As String
Dim strDesMar                       As String
Dim strCodUM                        As String
Dim strDesUM                        As String
Dim intAfecto                       As Integer
Dim intFila                         As Integer
Dim dblTC                           As Double
Dim dblVVUnit                       As Double
Dim dblIGVUnit                      As Double
Dim dblPVUnit                       As Double
Dim dblFactor                       As Double
Dim WsVale                          As Boolean
Dim codigo                          As String
Dim Descripcion                     As String
Dim codproducto                     As String
Dim codalmacen                      As String
Dim indPedido                       As Boolean
Dim SwIngreso                       As Boolean

    intFila = Node.Index + 1
    
    Select Case Column.Index
        Case gDetalle.Columns.ColumnByFieldName("idProducto").Index
            strCod = gDetalle.Columns.ColumnByFieldName("idProducto").Value
            strDes = gDetalle.Columns.ColumnByFieldName("GlsProducto").Value
            strCodUM = gDetalle.Columns.ColumnByFieldName("idUM").Value
            
            If txtCod_Almacen.Text = "" Then
                StrMsgError = "Ingrese Almacen"
                txtCod_Almacen.OnError = True
                GoTo Err
            End If
            strCod = ""
            strDes = ""
            strCodUM = ""
            indPedido = False
            If indVale = "I" Then
                WsVale = True
            End If
            FrmAyudaProdOC.ExecuteReturnTextAlm txtCod_Almacen.Text, rscd, strCod, strDes, strCodUM, glsValidaStock, "", False, True, indPedido, WsVale, StrMsgError
            If rscd.RecordCount <> 0 Then
                mostrarDocImportado_Ayuda rscd, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            
        Case gDetalle.Columns.ColumnByName("Imprime").Index
'            If txtCod_Vale.Text <> "" And ("" & gDetalle.Columns.ColumnByFieldName("idProducto").Value) <> "" Then
'                ImprimeCodigoBarra 0, ("" & gDetalle.Columns.ColumnByFieldName("idProducto").Value), txtCod_Vale.Text, StrMsgError
'                If StrMsgError <> "" Then GoTo Err
'            End If
            
        Case gDetalle.Columns.ColumnByFieldName("idlote").Index
            codalmacen = "" & txtCod_Almacen.Text
            codproducto = "" & gDetalle.Columns.ColumnByFieldName("idproducto").Value
            FrmAyudaLotes_Vales.mostrar_from indVale, Descripcion, codigo, codproducto, codalmacen
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idLote").Value = Trim("" & codigo)
            gDetalle.Columns.ColumnByFieldName("NumLote").Value = Trim("" & Descripcion)
            gDetalle.Dataset.Post
'        Case gDetalle.Columns.ColumnByFieldName("IdAtributo").Index
'            FrmDetAtributos.MostrarForm StrMsgError, RsDetAtributos, SwIngreso, Val("" & gDetalle.Columns.ColumnByFieldName("Item").Value), Trim("" & gDetalle.Columns.ColumnByFieldName("GlsProducto").Value)
'            If StrMsgError <> "" Then GoTo Err
            
    End Select
    calcularTotales
    gDetalle.Dataset.RecNo = intFila
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gDetalle_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim dblVVUnit                   As Double
Dim dblIGVUnit                  As Double
Dim dblPVUnit                   As Double
Dim strCod                      As String
Dim strDes                      As String
Dim strCodFabri                 As String
Dim strCodMar                   As String
Dim strDesMar                   As String
Dim intAfecto                   As Integer
Dim strTipoProd                 As String
Dim strMoneda                   As String
Dim strCodUM                    As String
Dim strDesUM                    As String
Dim dblFactor                   As Double
Dim StrMsgError                 As String
Dim rsp                         As New ADODB.Recordset
Dim NCantidadOC                 As Double
Dim NCantidadVale               As Double
Dim intFila                     As Long

    If gDetalle.Dataset.Modified = False Then Exit Sub
    
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    
    Select Case gDetalle.Columns.FocusedColumn.Index
        Case gDetalle.Columns.ColumnByFieldName("VVUnit").Index
            procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 0, Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value), dblVVUnit, dblIGVUnit, dblPVUnit
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), dblVVUnit, dblIGVUnit, dblPVUnit, Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
            gDetalle.Dataset.Post
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("PVUnit").Index
            procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 1, Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value), dblVVUnit, dblIGVUnit, dblPVUnit
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), dblVVUnit, dblIGVUnit, dblPVUnit, Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
            gDetalle.Dataset.Post
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("Cantidad").Index
            
            If Trim("" & gDetalle.Columns.ColumnByFieldName("idDocumentoImp").Value) = "94" Then
            
                If leeParametro("VALIDA_CANTIDAD_OCOMPRA") = "1" Then
                    
                    If txtCod_Vale.Text = "" Then
                    
                        NCantidadOC = Val("" & traerCampo("DocVentasDet", "Cantidad - CantidadImp", "IdDocumento", Trim("" & gDetalle.Columns.ColumnByFieldName("IdDocumentoImp").Value), True, "IdSerie = '" & Trim("" & gDetalle.Columns.ColumnByFieldName("IdSerieImp").Value) & "' And IdDocVentas = '" & Trim("" & gDetalle.Columns.ColumnByFieldName("IdDocVentasImp").Value) & "' And IdSucursal = '" & Trim("" & gDetalle.Columns.ColumnByFieldName("IdSucursalOrigen").Value) & "' And IdProducto = '" & gDetalle.Columns.ColumnByFieldName("IdProducto").Value & "'"))
                        
                    Else
                        
                        NCantidadVale = Val("" & traerCampo("ValesDet", "Cantidad", "IdValesCab", txtCod_Vale.Text, True, "TipoVale = '" & indVale & "' And IdSucursal = '" & glsSucursal & "' And Item = " & Val("" & gDetalle.Columns.ColumnByFieldName("Item").Value) & ""))
                        NCantidadOC = Val("" & traerCampo("DocVentasDet", "Cantidad - (CantidadImp - " & NCantidadVale & ")", "IdDocumento", Trim("" & gDetalle.Columns.ColumnByFieldName("IdDocumentoImp").Value), True, "IdSerie = '" & Trim("" & gDetalle.Columns.ColumnByFieldName("IdSerieImp").Value) & "' And IdDocVentas = '" & Trim("" & gDetalle.Columns.ColumnByFieldName("IdDocVentasImp").Value) & "' And IdSucursal = '" & Trim("" & gDetalle.Columns.ColumnByFieldName("IdSucursalOrigen").Value) & "' And IdProducto = '" & gDetalle.Columns.ColumnByFieldName("IdProducto").Value & "'"))
                        
                    End If
                    
                    If Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value) > NCantidadOC Then
                        
                        gDetalle.Dataset.Edit
                        gDetalle.Columns.ColumnByFieldName("Cantidad").Value = NCantidadOC
                        calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("IGVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
                        gDetalle.Dataset.Post
                        
                        StrMsgError = "La Cantidad ingresada es mayor a la Cantidad(" & Trim("" & NCantidadOC) & ") de la Orden de Compra N° " & Trim("" & gDetalle.Columns.ColumnByFieldName("IdDocVentasImp").Value) & ""
                        
                        calcularTotales
                        
                        GoTo Err
                    
                    End If
                    
                End If
            
            End If
            
            gDetalle.Dataset.Edit
            
            gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Value = Val(Format(Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value) * Val("" & traerCampo("Productos", "IdTallaPeso", "IdProducto", gDetalle.Columns.ColumnByFieldName("IdProducto").Value, True)), "0.00"))
            
            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("IGVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
            gDetalle.Dataset.Post
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("Afecto").Index
            procesaMoneda Val("" & gDetalle.Columns.ColumnByFieldName("idMoneda").Value), txtCod_Moneda.Text, 0, Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value), dblVVUnit, dblIGVUnit, dblPVUnit

            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit

            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), dblVVUnit, dblIGVUnit, dblPVUnit, Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)

            gDetalle.Dataset.Post
            
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("IdProducto").Index
            If Len(Trim(gDetalle.Columns.ColumnByFieldName("idProducto").Value)) > 0 Then
                strCod = gDetalle.Columns.ColumnByFieldName("idProducto").Value
                strDes = gDetalle.Columns.ColumnByFieldName("GlsProducto").Value
        
                csql = "SELECT idProducto,GlsProducto,idUMVenta,AfectoIGV FROM Productos " & _
                        "WHERE idempresa = '" & glsEmpresa & _
                        "' AND (idProducto = '" & strCod & "' OR idFabricante = '" & strCod & "' OR CodigoRapido = '" & strCod & "')"
                rsp.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                If rsp.EOF Or rsp.BOF Then
                    StrMsgError = "No se encuentra registrado el producto"
                    gDetalle.Dataset.Edit
                    gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
                    gDetalle.Dataset.Post
                    GoTo Err
                Else
                    mostrarDocImportado_Ayuda rsp, StrMsgError
                End If
                gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").ColIndex
            End If
    End Select
    If rsp.State = 1 Then rsp.Close: Set rsp = Nothing
    
    Exit Sub

Err:
    If rsp.State = 1 Then rsp.Close: Set rsp = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
End Sub

Private Sub gdetalle_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gDetalle.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                
                With RsDetAtributos
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Filter = "Item = " & Val(gDetalle.Columns.ColumnByFieldName("Item").Value) & ""
                        Do While Not .EOF
                            .Delete adAffectCurrent
                            .Update
                            .MoveNext
                        Loop
                        .Filter = ""
                        .Filter = adFilterNone
                    End If
                End With
                    
                If gDetalle.Count = 1 Then
                    gDetalle.Dataset.Edit
                    gDetalle.Columns.ColumnByFieldName("item").Value = 1
                    gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idUM").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsUM").Value = ""
                    gDetalle.Columns.ColumnByFieldName("Factor").Value = 1
                    gDetalle.Columns.ColumnByFieldName("Afecto").Value = 1
                    gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
                    gDetalle.Columns.ColumnByFieldName("Cantidad2").Value = 0
                    gDetalle.Columns.ColumnByFieldName("VVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("PVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("NumLote").Value = ""
                    gDetalle.Columns.ColumnByFieldName("FecVencProd").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idLote").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Value = ""
                    gDetalle.Dataset.Post
                
                Else
                    gDetalle.Dataset.Delete
                    gDetalle.Dataset.First
                    Do While Not gDetalle.Dataset.EOF
                        i = i + 1
                        With RsDetAtributos
                            If .RecordCount > 0 Then
                                .MoveFirst
                                .Filter = "Item = " & Val(gDetalle.Columns.ColumnByFieldName("Item").Value) & ""
                                Do While Not .EOF
                                    .Fields("Item") = i
                                    .Update
                                    .MoveNext
                                Loop
                                .Filter = ""
                                .Filter = adFilterNone
                            End If
                        End With
                        gDetalle.Dataset.Edit
                        gDetalle.Columns.ColumnByFieldName("Item").Value = i
                        gDetalle.Dataset.Post
                        gDetalle.Dataset.Next
                    Loop
                    
                    If gDetalle.Dataset.State = dsEdit Or gDetalle.Dataset.State = dsInsert Then
                        gDetalle.Dataset.Post
                    End If
                End If
                calcularTotales
                gDetalle.SetFocus
                gDetalle.Columns.FocusedIndex = 1
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gDetalle.Dataset.State = dsEdit Or gDetalle.Dataset.State = dsInsert Then
            gDetalle.Dataset.Post
        End If
    End If
    
End Sub

Private Sub gDetalle_OnKeyPress(Key As Integer)
On Error GoTo Err
Dim StrMsgError As String
Dim strCod As String
Dim strDes As String
Dim strMoneda As String
Dim strDesMar As String
Dim strCodUM As String
Dim strDesUM As String
Dim intAfecto As Integer
Dim intFila As Integer
Dim dblTC As Double
Dim dblVVUnit As Double
Dim dblIGVUnit As Double
Dim dblPVUnit As Double
Dim dblFactor As Double
    
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    
    If Key <> 9 And Key <> 13 And Key <> 27 Then
        Select Case gDetalle.Columns.FocusedColumn.Index
            Case gDetalle.Columns.ColumnByFieldName("idProducto").Index
        End Select
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub gDocReferencia_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    If Action = daInsert Then
        gDocReferencia.Columns.ColumnByFieldName("item").Value = gDocReferencia.Count
        gDocReferencia.Dataset.Post
    End If

End Sub

Private Sub gDocReferencia_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If Action = daInsert Then
        If (gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value = "" Or gDocReferencia.Columns.ColumnByFieldName("idSerie").Value = "" Or gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value = "") And indInsertaDocRef = False Then
            Allow = False
        Else
            gDocReferencia.Columns.FocusedIndex = gDocReferencia.Columns.ColumnByFieldName("idDocumento").ColIndex
        End If
    End If

End Sub

Private Sub gDocReferencia_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim strDes As String
    
    Select Case Column.Index
        Case gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Index
            strCod = gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value
            strDes = gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value
            
            mostrarAyudaTexto IIf(indVale = "I", "DOCUMENTOSI", "DOCUMENTOS"), strCod, strDes
            
            gDocReferencia.Dataset.Edit
            gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value = strCod
            gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value = strDes
            gDocReferencia.Dataset.Post
            
            gDocReferencia.SetFocus
            
    End Select

End Sub

Private Sub gDocReferencia_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String
    
    If gDocReferencia.Dataset.Modified = False Then Exit Sub
    
    Select Case gDocReferencia.Columns.FocusedColumn.Index
        Case gDocReferencia.Columns.ColumnByFieldName("idSerie").Index
            gDocReferencia.Dataset.Edit
            gDocReferencia.Columns.ColumnByFieldName("idSerie").Value = Format(gDocReferencia.Columns.ColumnByFieldName("idSerie").Value, "0000")
            gDocReferencia.Dataset.Post
        
        Case gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Index
            gDocReferencia.Dataset.Edit
            gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value = Format(gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value, "00000000")
            gDocReferencia.Dataset.Post
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gDocReferencia_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer
    
    If KeyCode = 46 Then
        If gDocReferencia.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                           
                If gDocReferencia.Count = 1 Then
                    gDocReferencia.Dataset.Edit
                    gDocReferencia.Columns.ColumnByFieldName("Item").Value = 1
                    gDocReferencia.Columns.ColumnByFieldName("idAlmacen").Value = ""
                    gDocReferencia.Columns.ColumnByFieldName("GlsAlmacen").Value = ""
                    gDocReferencia.Dataset.Post
                Else
                    gDocReferencia.Dataset.Delete
                    gDocReferencia.Dataset.First
                    Do While Not gDocReferencia.Dataset.EOF
                        i = i + 1
                        gDocReferencia.Dataset.Edit
                        gDocReferencia.Columns.ColumnByFieldName("Item").Value = i
                        gDocReferencia.Dataset.Post
                        gDocReferencia.Dataset.Next
                    Loop
                    If gDocReferencia.Dataset.State = dsEdit Or gDocReferencia.Dataset.State = dsInsert Then
                        gDocReferencia.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    If KeyCode = 13 Then
        If gDocReferencia.Dataset.State = dsEdit Or gDocReferencia.Dataset.State = dsInsert Then
              gDocReferencia.Dataset.Post
        End If
    End If

End Sub

Private Sub gDocReferencia_OnKeyPress(Key As Integer)
Dim strCod As String
Dim strDes As String
    
    If Key <> 9 And Key <> 13 And Key <> 27 Then
    Select Case gDocReferencia.Columns.FocusedColumn.Index
        Case gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Index
            strCod = gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value
            strDes = gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value
            
            mostrarAyudaKeyasciiTexto Key, IIf(indVale = "I", "DOCUMENTOSI", "DOCUMENTOS"), strCod, strDes
            Key = 0
            
            gDocReferencia.Dataset.Edit
            gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value = strCod
            gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value = strDes
            gDocReferencia.Dataset.Post
    End Select
    End If

End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    ListaDetalle

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarVale gLista.Columns.ColumnByName("idValesCab").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraDetalle.Visible = True
    fraTotales.Visible = True
    fraGeneral.Enabled = False
    fraDetalle.Enabled = False

    habilitaBotones 2
    habilitarColumas False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    Exit Sub
    Resume
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim rscd                            As ADODB.Recordset
Dim rsdd                            As ADODB.Recordset
Dim StrMsgError                     As String
Dim strCodUsuarioAutorizacion       As String
Dim IndEvaluacion                   As Integer
Dim strPeriodo                      As String
Dim CIdValesCabRef                  As String
Dim strTipoDocImportado             As String
Dim CTipoVale                       As String
Dim CNumMov                         As String
Dim reporte                         As String
Dim CValeTrans                      As String

Dim rsReporte       As New ADODB.Recordset
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim xreporte        As CRAXDRT.Report

    Select Case Button.Index
        Case 1 'Nuevo
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
            habilitaBotones Button.Index
            habilitarColumas True
            
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraDetalle.Visible = True
            fraTotales.Visible = True
            
            fraGeneral.Enabled = True
            fraDetalle.Enabled = True
            fraTotales.Enabled = True
                        
            strPeriodo = Format(dtp_Emision.Value, "yyyymm")
            If Trim("" & traerCampo("parametros", "valparametro", "glsparametro", "PERIODO_CAMBIO_IGV", True)) > strPeriodo Then
                dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV_ANT", True), "0.00")) / 100, "0.00")
            Else
                dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV", True), "0.00")) / 100, "0.00")
            End If
            
            txtCod_usuario.Text = glsUser
            
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            fraGeneral.Enabled = False
            fraDetalle.Enabled = False
            fraTotales.Enabled = False
            
            habilitaBotones Button.Index
        
        Case 3 'Modificar
            getEstadoCierreMes CVDate(dtp_Emision.Value), StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            CValeTrans = traerCampo("ValesTrans", "IdValesTrans", IIf(indVale = "I", "IdValeIngreso", "IdValeSalida"), txtCod_Vale.Text, True)
            
            If Len(Trim("" & CValeTrans)) > 0 Then
                
                StrMsgError = "No se puede modificar el vale porque ha sido generado de la Transferencia " & CValeTrans: GoTo Err
            
            End If
            
            If strEstVale = "GEN" Then
                IndEvaluacion = 0
                
                If indVale = "S" Then
                    CIdValesCabRef = "" & traerCampo("ValesCab", "IdValesCabRef", "IdSucursal", glsSucursal, True, "TipoVale = 'S' And IdValesCab = '" & txtCod_Vale.Text & "'")
                    If Len(Trim(CIdValesCabRef)) > 0 Then
                        StrMsgError = "El Vale de Salida Nº " & txtCod_Vale.Text & " no se puede modificar porque ha sido generado del Vale de Ingreso Nº " & CIdValesCabRef & ", verifique.": GoTo Err
                    End If
                End If
                
                CNumMov = "" & traerCampo("TbComprasVales A Inner Join RegisDoc B On A.IdEmpresa = B.IdEmpresa And A.Annio_Mov = B.Annio_Mov And A.IdMesMov = B.IdMesMov And A.IdNumMov = B.IdNumMov", "A.IdNumMov", "A.IdValesCab", txtCod_Vale.Text, False, "A.IdEmpresa = '" & glsEmpresa & "' And B.TipoDcto " & IIf(indVale = "S", "", "Not ") & "In('07')")
                
                If Len(Trim(CNumMov)) > 0 Then
                    
                    StrMsgError = "El Vale ha sido importado en el Registro de Compras Mov.(" & CNumMov & "), no se puede modificar.": GoTo Err
                    
                End If
                
                frmAprobacion.MostrarForm "04", IndEvaluacion, strCodUsuarioAutorizacion, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                
                If IndEvaluacion = 0 Then Exit Sub
            
                fraGeneral.Enabled = True
                fraDetalle.Enabled = True
                fraTotales.Enabled = True
                
                habilitaBotones Button.Index
                habilitarColumas True
            Else
                StrMsgError = "No se puede Modificar el Vale"
                GoTo Err
            End If
            
        Case 4 'Cancelar
            fraGeneral.Enabled = False
            fraDetalle.Enabled = False
            fraTotales.Enabled = False
            
            habilitaBotones Button.Index
            habilitarColumas False
            
        Case 5 'Anular
            getEstadoCierreMes CVDate(dtp_Emision.Value), StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            CValeTrans = traerCampo("ValesTrans", "IdValesTrans", IIf(indVale = "I", "IdValeIngreso", "IdValeSalida"), txtCod_Vale.Text, True)
            
            If Len(Trim("" & CValeTrans)) > 0 Then
                
                StrMsgError = "No se puede anular el vale porque ha sido generado de la Transferencia " & CValeTrans: GoTo Err
            
            End If
            
            CNumMov = "" & traerCampo("TbComprasVales A Inner Join RegisDoc B On A.IdEmpresa = B.IdEmpresa And A.Annio_Mov = B.Annio_Mov And A.IdMesMov = B.IdMesMov And A.IdNumMov = B.IdNumMov", "A.IdNumMov", "A.IdValesCab", txtCod_Vale.Text, False, "A.IdEmpresa = '" & glsEmpresa & "' And B.TipoDcto " & IIf(indVale = "S", "", "Not ") & "In('07')")
                
            If Len(Trim(CNumMov)) > 0 Then
                
                StrMsgError = "El Vale ha sido importado en el Registro de Compras Mov.(" & CNumMov & "), no se puede anular.": GoTo Err
                
            End If
                
            anularDoc StrMsgError, Button.Index
            If StrMsgError <> "" Then GoTo Err
            
            fraGeneral.Enabled = False
            fraDetalle.Enabled = False
            fraTotales.Enabled = False
            
        Case 6 'Eliminar
        
            StrMsgError = "La opción se encuentra des habilitada.": GoTo Err
            
            getEstadoCierreMes CVDate(dtp_Emision.Value), StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            CValeTrans = traerCampo("ValesTrans", "IdValesTrans", IIf(indVale = "I", "IdValeIngreso", "IdValeSalida"), txtCod_Vale.Text, True)
            
            If Len(Trim("" & CValeTrans)) > 0 Then
                
                StrMsgError = "No se puede eliminar. el vale porque ha sido generado de la Transferencia " & CValeTrans: GoTo Err
            
            End If
            
            CNumMov = "" & traerCampo("TbComprasVales A Inner Join RegisDoc B On A.IdEmpresa = B.IdEmpresa And A.Annio_Mov = B.Annio_Mov And A.IdMesMov = B.IdMesMov And A.IdNumMov = B.IdNumMov", "A.IdNumMov", "A.IdValesCab", txtCod_Vale.Text, False, "A.IdEmpresa = '" & glsEmpresa & "' And B.TipoDcto " & IIf(indVale = "S", "", "Not ") & "In('07')")
                
            If Len(Trim(CNumMov)) > 0 Then
                
                StrMsgError = "El Vale ha sido importado en el Registro de Compras Mov.(" & CNumMov & "), no se puede eliminar.": GoTo Err
                
            End If
                
            EliminarVale StrMsgError, Button.Index
            If StrMsgError <> "" Then GoTo Err
            
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraDetalle.Visible = False
            fraTotales.Visible = False
            
            listaVales StrMsgError
            If StrMsgError <> "" Then GoTo Err
            '''habilitaBotones Button.Index
            cmbImprimirBloque.Visible = False
            
        Case 7 'Importar
            frmListaDocExportar_OC.MostrarForm "99", txtCod_Cliente.Text, rscd, rsdd, strTipoDocImportado, StrMsgError
            If StrMsgError <> "" Then GoTo Err
    
            If strTipoDocImportado <> "" Then
                mostrarDocImportado rscd, rsdd, strTipoDocImportado, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            Unload frmListaDocExportar
        
        Case 8 'Imprimir
            If txtCod_Vale.Text <> "" Then
'''''''                If Trim(Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "FORMATO_VALE_ALMACEN", True))) = "S" Then
'''''''                    If Trim("" & traerCampo("conceptos", "indCosto", "idConcepto", txtCod_Concepto.Text, False, " tipoConcepto = '" & indVale & "' ")) = "N" Then
'''''''                        reporte = "rptImpVale" & indVale & "2.rpt"
'''''''                    Else
'''''''                        If indVale = "S" Then
'''''''                            reporte = "rptImpValeCon" & indVale & "2.rpt"
'''''''                        Else
'''''''                            reporte = "rptImpVale" & indVale & "2.rpt"
'''''''                        End If
'''''''                    End If
'''''''                Else
'''''''                    reporte = "rptImpVale" & indVale & ".rpt"
'''''''                End If
                
                CTipoVale = ""
                CTipoVale = Trim(Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "FORMATO_VALE_ALMACEN", True)))
                Select Case CTipoVale
                    Case "S":
                        If Trim("" & traerCampo("conceptos", "indCosto", "idConcepto", TxtCod_Concepto.Text, False, " tipoConcepto = '" & indVale & "' ")) = "N" Then
                            reporte = "rptImpVale" & indVale & "2.rpt"
                        Else
                            If indVale = "S" Then
                                reporte = "rptImpValeCon" & indVale & "2.rpt"
                            Else
                                reporte = "rptImpVale" & indVale & "2.rpt"
                            End If
                        End If
                    Case "3":
                        If Trim("" & traerCampo("conceptos", "indCosto", "idConcepto", TxtCod_Concepto.Text, False, " tipoConcepto = '" & indVale & "' ")) = "N" Then
                            reporte = "rptImpVale" & indVale & "3.rpt"
                        Else
                            If indVale = "S" Then
                                reporte = "rptImpValeCon" & indVale & "3.rpt"
                            Else
                                reporte = "rptImpVale" & indVale & "3.rpt"
                            End If
                        End If
                    Case "4":
                        If Trim("" & traerCampo("conceptos", "indCosto", "idConcepto", TxtCod_Concepto.Text, False, " tipoConcepto = '" & indVale & "' ")) = "N" Then
                            reporte = "rptImpVale" & indVale & "4.rpt"
                        Else
                            If indVale = "S" Then
                                reporte = "rptImpValeCon" & indVale & "4.rpt"
                            Else
                                reporte = "rptImpVale" & indVale & "4.rpt"
                            End If
                        End If
                    Case Else
                        reporte = "rptImpVale" & indVale & ".rpt"
                End Select
                
                
                Screen.MousePointer = 1
    
                
                mostrarReporte reporte, "varEmpresa|varSucursal|varNumvale|varTipovale", glsEmpresa & "|" & glsSucursal & "|" & txtCod_Vale.Text & "|" & indVale, "vale", StrMsgError
                If StrMsgError <> "" Then GoTo Err
                
            End If
            
        Case 9 'Lista
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraDetalle.Visible = False
            fraTotales.Visible = False
            
            listaVales StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            habilitaBotones Button.Index
            cmbImprimirBloque.Visible = False
        
        Case 10: 'Importar Vales
            Dim F As New FrmImportaVales
            Load F
            F.Show
        
        Case 11: 'Excel
            gLista.m.ExportToXLS App.Path & "\Temporales\vales.xls"
            ShellEx App.Path & "\Temporales\vales.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        
        Case 12: 'Datos Transito
            FrmDatosTransito.MostrarForm txtCod_Vale.Text, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            
        Case 13: 'Salir
            Unload Me
    End Select
    
Exit Sub
Err:
    
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
Exit Sub
Resume
    
End Sub

Private Sub habilitaBotones(indexBoton As Integer)
Dim indHabilitar As Boolean
    
    Toolbar1.Buttons(10).Visible = False 'Importar Vales
    
    Select Case indexBoton
        Case 1 'Nuevo
            Toolbar1.Buttons(1).Visible = False 'Nuevo
            Toolbar1.Buttons(2).Visible = True 'Grabar
            Toolbar1.Buttons(3).Visible = False 'Modificar
            Toolbar1.Buttons(4).Visible = False 'Cancelar
            Toolbar1.Buttons(5).Visible = False 'Anular
            Toolbar1.Buttons(6).Visible = False 'Eliminar
            Toolbar1.Buttons(7).Visible = True 'Importar
            Toolbar1.Buttons(8).Visible = False 'Imprimir
            Toolbar1.Buttons(9).Visible = True 'Lista
            Toolbar1.Buttons(11).Visible = False 'excel
            Toolbar1.Buttons(12).Visible = False
        Case 2 'Grabar
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(4).Visible = False
            If traerCampo("Conceptos", "indAutomatico", "idConcepto", Trim(TxtCod_Concepto.Text), False) = "1" And Len(Trim(gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value)) > 0 Then
                Toolbar1.Buttons(3).Visible = False
                Toolbar1.Buttons(5).Visible = False
                Toolbar1.Buttons(6).Visible = False
            Else
                Toolbar1.Buttons(3).Visible = True
                Toolbar1.Buttons(5).Visible = True
                Toolbar1.Buttons(6).Visible = True
            End If
            Toolbar1.Buttons(7).Visible = False 'Importar
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = True
            Toolbar1.Buttons(11).Visible = False
            '--- 07/07/13
            If leeParametro("VISUALIZA_DATOS_TRANSITO") = "1" Then
            
                If indVale = "S" Then
                   Toolbar1.Buttons(12).Visible = True
                Else
                   Toolbar1.Buttons(12).Visible = False
                End If
                
            Else
                
                Toolbar1.Buttons(12).Visible = False
            
            End If
            
        Case 3 'Modificar
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = True
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False 'Eliminar
            Toolbar1.Buttons(7).Visible = True 'Importar
            Toolbar1.Buttons(8).Visible = False
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(11).Visible = False
            '--- 07/07/13
            If leeParametro("VISUALIZA_DATOS_TRANSITO") = "1" Then
            
                If indVale = "S" Then
                   Toolbar1.Buttons(12).Visible = True
                Else
                   Toolbar1.Buttons(12).Visible = False
                End If
                
            Else
                
                Toolbar1.Buttons(12).Visible = False
            
            End If
            
        Case 4 'Cancelar
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = True
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False 'Importar
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = True
            Toolbar1.Buttons(11).Visible = False
        Case 5 'Anular
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False 'Importar
            Toolbar1.Buttons(8).Visible = False
            Toolbar1.Buttons(9).Visible = True
            Toolbar1.Buttons(11).Visible = False
            Toolbar1.Buttons(12).Visible = False
        Case 6 'Eliminar
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False 'Importar
            Toolbar1.Buttons(8).Visible = False
            Toolbar1.Buttons(9).Visible = True
            Toolbar1.Buttons(11).Visible = False
            Toolbar1.Buttons(12).Visible = False
        Case 9 'Listar
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False 'Importar
            Toolbar1.Buttons(8).Visible = False
            Toolbar1.Buttons(9).Visible = False
            If indVale = "S" Then
                If Len(Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "CONCEPTO_CONSUMO_SALIDA", True))) > 0 Then
                    Toolbar1.Buttons(10).Visible = True 'Importar Vales
                End If
            End If
            Toolbar1.Buttons(11).Visible = True
            Toolbar1.Buttons(12).Visible = False
    End Select
    
'    If Len(Trim(traerCampo("PermisosUsuarios", "IdPermiso", "IdPermiso", "03", True, "IdUsuario = '" & glsUser & "' And CodSistema = '" & StrcodSistema & "'"))) = 0 Then
'        Toolbar1.Buttons(1).Visible = False
'    End If
            
    If Trim("" & traerCampo("Conceptos", "IndAutomatico", "IdConcepto", TxtCod_Concepto.Text, False)) = "1" Then
        Toolbar1.Buttons(3).Visible = False 'Modificar
        Toolbar1.Buttons(5).Visible = False 'Anular
        Toolbar1.Buttons(6).Visible = False 'Eliminar
    End If
            
End Sub

Private Sub txt_Ano_Change()
On Error GoTo Err
Dim StrMsgError As String

    If indNuevoDoc = False Then
        listaVales StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaVales StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaVales(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND (p.GlsPersona LIKE '%" & strCond & "%' or v.GlsDocReferencia LIKE '%" & strCond & "%' or v.idValesCab LIKE '%" & strCond & "%' )"
    End If
    
    csql = "SELECT idValesCab,fechaEmision,v.idConcepto,c.GlsConcepto,idProvCliente,p.GlsPersona,v.idAlmacen,a.GlsAlmacen,v.estValeCab, v.GlsDocReferencia  " & _
            "FROM valescab v inner join Conceptos c " & _
            "on v.idConcepto = c.idConcepto " & _
            "left join personas p on v.idProvCliente = p.idPersona " & _
            "inner join almacenes a on v.IdEmpresa = a.IdEmpresa And v.idAlmacen = a.idAlmacen " & _
            "WHERE tipoVale = '" & indVale & "' " & _
            "AND v.idEmpresa = '" & glsEmpresa & "' AND v.idSucursal = '" & glsSucursal & "' " & _
            "AND  year(fechaEmision) = " & Val(txt_Ano.Text) & " AND Month(fechaEmision) = " & cbx_Mes.ListIndex + 1
           
    If strCond <> "" Then csql = csql & strCond

    csql = csql & " ORDER BY idValesCab"
    
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
'        .KeyField = "idValesCab"
'    End With
    
    ListaDetalle
    Me.Refresh

Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarVale(strNum As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst                             As New ADODB.Recordset
Dim rsg                             As New ADODB.Recordset
Dim RsD                             As New ADODB.Recordset
Dim CSqlC                           As String
Dim RsC                             As New ADODB.Recordset

    indCargando = True
    
    TxtCod_Concepto.Text = ""
    
    CSqlC = "SELECT idValesCab, tipoVale, fechaEmision, valorTotal, igvTotal, precioTotal, idProvCliente, idConcepto, IdAlmacen, obsValesCab, idMoneda, GlsDocReferencia, TipoCambio, idEmpresa, idSucursal, estValeCab, idPeriodoInv, idCentroCosto, codAnula, obsAnulacion, fecAnulacion, usuAnula, IdValeTemp, TipoValeRef, IdValesCabRef,idUsuarioRegistro " & _
           "FROM valescab d " & _
           "WHERE d.idValesCab = '" & strNum & "' AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "' and d.tipoVale = '" & indVale & "' "
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        strEstVale = Trim("" & rst.Fields("estValeCab"))
        If strEstVale = "ANU" Then
            lbl_Anulado.Caption = "ANULADO"
            fraGeneral.Enabled = False
            habilitarColumas False
        Else
            lbl_Anulado.Caption = ""
        End If
        
        CIdAlmacenAnt = Trim("" & rst.Fields("IdAlmacen"))
        
    End If
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    txtCod_AlmacenTrans.Text = traerCampo("ValesTrans", IIf(indVale = "I", "IdAlmacenOrigen", "IdAlmacenDestino"), IIf(indVale = "I", "IdValeIngreso", "IdValeSalida"), strNum, True, "IdSucursal = '" & glsSucursal & "'")
    
    CSqlC = "SELECT * " & _
           "FROM valesdet " & _
           "WHERE idValesCab = '" & strNum & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' and tipoVale = '" & indVale & "' ORDER BY ITEM "
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adChar, 15, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Cantidad2", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 45, adFldIsNullable
    rsg.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idSucursalOrigen", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "idDocumentoImp", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "idDocVentasImp", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerieImp", adVarChar, 4, adFldIsNullable
    rsg.Fields.Append "idLote", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "IdTallaPeso", adVarChar, 30, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("idUM") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("Factor") = 1
        rsg.Fields("Afecto") = 1
        rsg.Fields("Cantidad") = 0
        rsg.Fields("Cantidad2") = 0
        rsg.Fields("VVUnit") = 0
        rsg.Fields("IGVUnit") = 0
        rsg.Fields("PVUnit") = 0
        rsg.Fields("TotalVVNeto") = 0
        rsg.Fields("TotalIGVNeto") = 0
        rsg.Fields("TotalPVNeto") = 0
        rsg.Fields("idMoneda") = ""
        rsg.Fields("NumLote") = ""
        rsg.Fields("FecVencProd") = ""
        rsg.Fields("idSucursalOrigen") = ""
        rsg.Fields("idDocumentoImp") = ""
        rsg.Fields("idDocVentasImp") = ""
        rsg.Fields("idSerieImp") = ""
        rsg.Fields("idLote") = ""
        rsg.Fields("IdTallaPeso") = ""
          
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idProducto") = "" & rst.Fields("idProducto")
            rsg.Fields("GlsProducto") = "" & rst.Fields("GlsProducto")
            rsg.Fields("idUM") = "" & rst.Fields("idUM")
            rsg.Fields("GlsUM") = traerCampo("unidadmedida", "abreUM", "idUM", ("" & rst.Fields("idUM")), False)
            rsg.Fields("Factor") = "" & rst.Fields("Factor")
            rsg.Fields("Afecto") = "" & rst.Fields("Afecto")
            rsg.Fields("Cantidad") = "" & rst.Fields("Cantidad")
            rsg.Fields("Cantidad2") = "" & rst.Fields("Cantidad2")
            rsg.Fields("VVUnit") = "" & rst.Fields("VVUnit")
            rsg.Fields("IGVUnit") = "" & rst.Fields("IGVUnit")
            rsg.Fields("PVUnit") = "" & rst.Fields("PVUnit")
            rsg.Fields("TotalVVNeto") = "" & rst.Fields("TotalVVNeto")
            rsg.Fields("TotalIGVNeto") = "" & rst.Fields("TotalIGVNeto")
            rsg.Fields("TotalPVNeto") = "" & rst.Fields("TotalPVNeto")
            rsg.Fields("idMoneda") = Trim("" & rst.Fields("idMoneda"))
            rsg.Fields("NumLote") = "" & rst.Fields("NumLote")
            rsg.Fields("FecVencProd") = "" & rst.Fields("FecVencProd")
            rsg.Fields("idSucursalOrigen") = "" & rst.Fields("idSucursalOrigen")
            rsg.Fields("idDocumentoImp") = "" & rst.Fields("idDocumentoImp")
            rsg.Fields("idDocVentasImp") = "" & rst.Fields("idDocVentasImp")
            rsg.Fields("idSerieImp") = "" & rst.Fields("idSerieImp")
            rsg.Fields("idLote") = "" & rst.Fields("idLote")
            rsg.Fields("IdTallaPeso") = "" & rst.Fields("IdTallaPeso")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Inicia_RecordSet RsDetAtributos, 0, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    CSqlC = "Select A.Item,A.IdAtributo,B.GlsAtributo,A.Valor " & _
            "From ValesDetAtributos A " & _
            "Inner Join Prod_Atributo B " & _
                "On A.IdEmpresa = B.IdEmpresa And A.IdAtributo = B.IdAtributo " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.TipoVale = '" & indVale & "' " & _
            "And A.IdValesCab = '" & strNum & "' " & _
            "Order By A.Item,A.IdAtributo"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockOptimistic
    
    Do While Not RsC.EOF
        
        RsDetAtributos.AddNew
        RsDetAtributos.Fields("Item") = Val("" & RsC.Fields("Item"))
        RsDetAtributos.Fields("IdAtributo") = RsC.Fields("IdAtributo")
        RsDetAtributos.Fields("GlsAtributo") = "" & RsC.Fields("GlsAtributo")
        RsDetAtributos.Fields("Valor") = Val("" & RsC.Fields("Valor"))
        
        RsC.MoveNext
        
    Loop
    
    RsC.Close: Set RsC = Nothing
    
    CSqlC = "SELECT r.item, r.tipoDocReferencia idDocumento,d.GlsDocumento, r.numDocReferencia idNumDoc, r.serieDocReferencia idSerie " & _
            "FROM docreferencia r , documentos d " & _
            "WHERE r.tipoDocReferencia = d.idDocumento AND tipoDocOrigen = '" & IIf(indVale = "I", "88", "99") & "' AND numDocOrigen = '" & strNum & "' AND serieDocOrigen = '000' AND r.idEmpresa = '" & glsEmpresa & "' AND r.idSucursal = '" & glsSucursal & "' ORDER BY ITEM"
    If rst.State = 1 Then rst.Close
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adChar, 8, adFldIsNullable
    RsD.Open
    
    If rst.RecordCount = 0 Then
        RsD.AddNew
        RsD.Fields("Item") = 1
        RsD.Fields("idDocumento") = ""
        RsD.Fields("GlsDocumento") = ""
        RsD.Fields("idSerie") = ""
        RsD.Fields("idNumDOc") = ""
    Else
        Do While Not rst.EOF
            RsD.AddNew
            RsD.Fields("Item") = "" & rst.Fields("Item")
            RsD.Fields("idDocumento") = "" & rst.Fields("idDocumento")
            RsD.Fields("GlsDocumento") = "" & rst.Fields("GlsDocumento")
            RsD.Fields("idSerie") = "" & rst.Fields("idSerie")
            RsD.Fields("idNumDOc") = "" & rst.Fields("idNumDOc")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gDocReferencia, RsD, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    indCargando = False
    dtp_Emision_Change
    
    calcularTotales
    
    Me.Refresh

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Function DatosProducto_Ayuda(strCodProd As String, ByRef strCodUM As String, ByRef strDesUM As String, ByRef dblFactor As Double) As Boolean
Dim rst As New ADODB.Recordset

    csql = "SELECT p.AfectoIGV,p.idMoneda,p.idUMCompra,u.abreUM " & _
            "FROM productos p,unidadmedida u " & _
            "WHERE p.idUMCompra = u.idUM " & _
            "AND p.idEmpresa = '" & glsEmpresa & "' " & _
            "AND p.idProducto = '" & strCodProd & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        DatosProducto_Ayuda = True
        strCodUM = "" & rst.Fields("idUMCompra")
        strDesUM = "" & rst.Fields("abreUM")
        dblFactor = 1
    Else
        DatosProducto_Ayuda = False
        strCodUM = ""
        strDesUM = ""
        dblFactor = 0
    End If
    rst.Close: Set rst = Nothing

End Function

Private Function DatosProducto(strCodProd As String, ByRef strGlsMarca As String, ByRef intAfecto As Integer, ByRef strMoneda As String, ByRef strCodUM As String, ByRef strDesUM As String, ByRef dblFactor As Double) As Boolean
Dim rst As New ADODB.Recordset

    csql = "SELECT m.GlsMarca,p.AfectoIGV,p.idMoneda,p.idUMCompra,u.abreUM AS GlsUM,x.factor " & _
            "FROM productos p,marcas m,unidadmedida u,presentaciones x " & _
            "WHERE p.idMarca = m.idMarca " & _
            "AND p.idUMCompra = u.idUM " & _
            "AND p.idProducto = x.idProducto " & _
            "AND p.idUMCompra = x.idUM " & _
            "AND p.idEmpresa = '" & glsEmpresa & "' " & _
            "AND m.idEmpresa = '" & glsEmpresa & "' " & _
            "AND x.idEmpresa = '" & glsEmpresa & "' " & _
            "AND p.idProducto = '" & strCodProd & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        DatosProducto = True
        strGlsMarca = "" & rst.Fields("GlsMarca")
        intAfecto = "" & rst.Fields("AfectoIGV")
        strMoneda = "" & rst.Fields("idMoneda")
        strCodUM = "" & rst.Fields("idUMCompra")
        strDesUM = "" & rst.Fields("GlsUM")
        dblFactor = Val("" & rst.Fields("factor"))
     Else
        DatosProducto = False
        strGlsMarca = ""
        intAfecto = 1
        strMoneda = ""
        strCodUM = ""
        strDesUM = ""
        dblFactor = 0
    End If
    rst.Close: Set rst = Nothing
    
End Function

Private Sub procesaMoneda(strMonProd As String, strMonDoc As String, intTipoValor As Integer, dblValor As Double, intAfecto As Integer, ByRef dblVVUnit As Double, ByRef dblIGVUnit As Double, ByRef dblPVUnit As Double)
Dim dblIGV As Double
Dim dblTC As Double
    
    dblIGV = dblIgvNEw
    dblTC = txt_TipoCambio.Value
    If intAfecto = 0 Then dblIGV = 0
    
    If strMonDoc = "USD" Then 'dolares
        If strMonProd = "PEN" Then 'soles
            dblValor = dblValor / dblTC
        End If
    Else 'soles
        If strMonProd = "USD" Then 'dolares
            dblValor = dblValor * dblTC
        End If
    End If
    
    If intAfecto = 1 Then
        If intTipoValor = 0 Then 'valor venta
            dblVVUnit = dblValor
            dblIGVUnit = dblValor * dblIGV
            dblPVUnit = dblVVUnit + dblIGVUnit
        Else 'precio venta
            dblVVUnit = dblValor / (dblIGV + 1)
            dblIGVUnit = dblValor - dblVVUnit
            dblPVUnit = dblValor
        End If
    Else
    
        dblVVUnit = dblValor
        dblIGVUnit = 0
        dblPVUnit = dblValor
        
    End If
    
End Sub

Private Sub calculaTotalesFila(dblCantidad As Double, dblVVUnit As Double, dblIGVUnit As Double, dblPVUnit As Double, intAfecto As Integer)
Dim dblTotalVVBruto As Double
Dim dblTotalPVBruto As Double
Dim dblTotalVVNeto As Double
Dim dblTotalIGVNeto As Double
Dim dblTotalPVNeto As Double
    
    dblTotalVVBruto = dblCantidad * dblVVUnit
    dblTotalPVBruto = dblCantidad * dblPVUnit
   
    dblTotalVVNeto = dblTotalVVBruto
    If intAfecto = 1 Then
        dblTotalIGVNeto = dblTotalVVNeto * dblIgvNEw
    Else
        dblTotalIGVNeto = 0
    End If
    dblTotalPVNeto = dblTotalVVNeto + dblTotalIGVNeto
    
    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = dblTotalVVNeto
    gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = dblTotalIGVNeto
    gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = dblTotalPVNeto

End Sub

Private Sub ListaDetalle()
Dim rsdatos                     As New ADODB.Recordset

    csql = "SELECT  V.Item,V.idProducto,V.GlsProducto,M.GlsMarca,V.idUM,U.GlsUM,CAST(V.Cantidad AS DECIMAL(12,2)) AS Cantidad, CAST(V.Cantidad2 AS DECIMAL(12,2)) AS Cantidad2, " & _
            "CAST(V.PVUnit AS DECIMAL(12,2)) AS PVUnit,CAST(V.TotalPVNeto AS DECIMAL(12,2)) AS TotalPVNeto,V.IdTallaPeso " & _
            "FROM valesdet V, productos P,marcas M,unidadmedida U " & _
            "WHERE V.idProducto = P.idProducto " & _
            "AND  P.idMarca = M.idMarca " & _
            "AND v.idEmpresa = '" & glsEmpresa & "' AND v.idSucursal = '" & glsSucursal & "' " & _
            "AND p.idEmpresa = '" & glsEmpresa & "' " & _
            "AND m.idEmpresa = '" & glsEmpresa & "' " & _
            "AND v.tipoVale = '" & indVale & "' " & _
            "AND  V.idUM = U.idUM AND idValesCab = '" & gLista.Columns.ColumnByFieldName("idValesCab").Value & "' " & _
            "ORDER BY V.Item"
    
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
        
    Set gListaDetalle.DataSource = rsdatos

'    With gListaDetalle
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "item"
'    End With

End Sub

Private Sub calcularTotales()
    
    txt_TotalBruto.Text = 0#
    txt_TotalIGV.Text = 0#
    txt_TotalNeto.Text = 0#
    TxtTotalPeso.Text = 0#
    gDetalle.Dataset.Refresh
    gDetalle.Dataset.First
    
    Do While Not gDetalle.Dataset.EOF
        txt_TotalBruto.Text = Val(txt_TotalBruto.Value) + Val("" & gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value)
        txt_TotalIGV.Text = Val(txt_TotalIGV.Value) + Val("" & gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value)
        txt_TotalNeto.Text = Val(txt_TotalNeto.Value) + Val("" & gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value)
        
        If gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Visible Then
            TxtTotalPeso.Text = Val(TxtTotalPeso.Value) + Val("" & gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Value)
        End If
        
        gDetalle.Dataset.Next
    Loop

End Sub

Private Sub eliminaNulosGrilla()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer
    
    indWhile = True
    Do While indWhile = True
        If gDetalle.Count >= 1 Then
            gDetalle.Dataset.First
            indEntro = False
            Do While Not gDetalle.Dataset.EOF
                If Trim(gDetalle.Columns.ColumnByFieldName("idProducto").Value) = "" Or gDetalle.Columns.ColumnByFieldName("Cantidad").Value <= 0 Then
                    gDetalle.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gDetalle.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gDetalle.Count >= 1 Then
        gDetalle.Dataset.First
        i = 0
        Do While Not gDetalle.Dataset.EOF
            i = i + 1
            
            With RsDetAtributos
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Filter = "Item = " & Val(gDetalle.Columns.ColumnByFieldName("Item").Value) & ""
                    Do While Not .EOF
                        .Fields("Item") = i
                        .Update
                        .MoveNext
                    Loop
                    .Filter = ""
                    .Filter = adFilterNone
                End If
            End With
            
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("item").Value = i
            If gDetalle.Dataset.State = dsEdit Then gDetalle.Dataset.Post
            gDetalle.Dataset.Next
        Loop
    Else
        indInserta = True
        gDetalle.Dataset.Append
        indInserta = False
    End If
    
End Sub

Private Sub eliminaNulosGrillaDocRef()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer
    
    indWhile = True
    Do While indWhile = True
        If gDocReferencia.Count >= 1 Then
            gDocReferencia.Dataset.First
            indEntro = False
            Do While Not gDocReferencia.Dataset.EOF
                If Trim(gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value) = "" Or gDocReferencia.Columns.ColumnByFieldName("idSerie").Value = "" Or gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value = "" Then
                    gDocReferencia.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gDocReferencia.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gDocReferencia.Count >= 1 Then
        gDocReferencia.Dataset.First
        i = 0
        Do While Not gDocReferencia.Dataset.EOF
            i = i + 1
            gDocReferencia.Dataset.Edit
            gDocReferencia.Columns.ColumnByFieldName("item").Value = i
            If gDocReferencia.Dataset.State = dsEdit Then gDocReferencia.Dataset.Post
            gDocReferencia.Dataset.Next
        Loop
        
    Else
        indInsertaDocRef = True
        gDocReferencia.Dataset.Append
        indInsertaDocRef = False
    End If
    
End Sub

Private Sub generaSTRDocReferencia()
Dim strAbre As String

    txt_DocReferencia.Text = ""
    If gDocReferencia.Count > 0 Then
        gDocReferencia.Dataset.First
        Do While Not gDocReferencia.Dataset.EOF
            If gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value <> "" Then
                strAbre = traerCampo("documentos", "AbreDocumento", "idDocumento", gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value, False)
                If txt_DocReferencia.Text = "" Then
                    txt_DocReferencia.Text = strAbre & " " & gDocReferencia.Columns.ColumnByFieldName("idSerie").Value & "-" & gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value
                Else
                    txt_DocReferencia.Text = txt_DocReferencia.Text & " / " & strAbre & " " & gDocReferencia.Columns.ColumnByFieldName("idSerie").Value & "-" & gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value
                End If
                If Trim(gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value) <> "01" Or Trim(gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value) <> "86" Then
                    Sw_Documento = True
                End If
            End If
            gDocReferencia.Dataset.Next
        Loop
    End If

End Sub

Private Sub txtCod_Almacen_Change()
    
    If txtCod_Almacen.Text = "" Then
        txtGls_Almacen.Text = ""
    Else
        txtGls_Almacen.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", txtCod_Almacen.Text, True)
    End If

End Sub

Private Sub txtCod_AlmacenTrans_Change()
    
    txtGls_AlmacenTrans.Text = traerCampo("Almacenes", "GlsAlmacen", "IdAlmacen", txtCod_AlmacenTrans.Text, True)
    If txtCod_AlmacenTrans.Text = "" Then
        Lbl_AlmacenTrans.Visible = False
        txtCod_AlmacenTrans.Visible = False
        txtGls_AlmacenTrans.Visible = False
    Else
        Lbl_AlmacenTrans.Visible = True
        Lbl_AlmacenTrans.Caption = "Almacén " & IIf(indVale = "I", "Ori.", "Dest.")
        txtCod_AlmacenTrans.Visible = True
        txtGls_AlmacenTrans.Visible = True
    End If
    
End Sub

Private Sub txtCod_CentroCosto_Change()
    
    If txtCod_CentroCosto.Text = "" Then
        txtGls_CentroCosto.Text = ""
    Else
        txtGls_CentroCosto.Text = traerCampo("centroscosto", "glsCentroCosto", "idCentroCosto", txtCod_CentroCosto.Text, True)
    End If

End Sub

Private Sub txtCod_Cliente_Change()
    
    If txtCod_Cliente.Text = "" Then
        txtGls_Cliente.Text = ""
    Else
        txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
    End If

End Sub

Private Sub txtCod_Cliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtCod_Cliente.Text = traerCampo("personas", "idPersona", "ruc", txtCod_Cliente.Text, False)
        If txtCod_Cliente.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Concepto_Change()
    
    If TxtCod_Concepto.Text = "" Then
        TxtGls_Concepto.Text = ""
    Else
        TxtGls_Concepto.Text = traerCampo("conceptos", "glsConcepto", "idConcepto", TxtCod_Concepto.Text, False, "TipoConcepto = '" & indVale & "'" & IIf(indCargando = True, "", " And IsNull(IndAutomatico,'') <> '1'"))
    End If

'    If Trim("" & traerCampo("conceptos", "indCosto", "idConcepto", txtCod_Concepto.Text, False, " tipoConcepto = '" & indVale & "'" & IIf(indCargando = True, "", " And IfNull(IndAutomatico,'') <> '1'"))) = "N" Then
'        gDetalle.Columns.ColumnByFieldName("VVUnit").Visible = False
'        'gDetalle.Columns.ColumnByFieldName("PVUnit").Visible = False
'        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Visible = False
'    Else
'        gDetalle.Columns.ColumnByFieldName("VVUnit").Visible = True
'        'gDetalle.Columns.ColumnByFieldName("PVUnit").Visible = True
'        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Visible = True
'    End If
    
End Sub

Private Sub txtCod_Concepto_LostFocus()
    
    If Len(Trim(TxtCod_Concepto.Text)) > 0 Then
        If Len(Trim(TxtGls_Concepto.Text)) = 0 Then
            MsgBox ("el Codigo de concepto no existe"), vbCritical
            If fraGeneral.Visible Then TxtCod_Concepto.SetFocus
        End If
    End If
    
End Sub

Private Sub txtCod_Moneda_Change()
    
    If txtCod_Moneda.Text = "" Then
        txtGls_Moneda.Text = ""
    Else
        txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
        lbl_SimbMonBruto.Caption = traerCampo("monedas", "Simbolo", "idMoneda", txtCod_Moneda.Text, False)
        lbl_SimbMonIGV.Caption = lbl_SimbMonBruto.Caption
        lbl_SimbMonNeto.Caption = lbl_SimbMonBruto.Caption
    End If

End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MONEDA", txtCod_Moneda, txtGls_Moneda
        KeyAscii = 0
        If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub anularDoc(ByRef StrMsgError As String, ByRef Button As Integer)
On Error GoTo Err
Dim rst                             As New ADODB.Recordset
Dim iniTrans                        As Boolean
Dim IndEvaluacion                   As Integer
Dim strCodUsuarioAutorizacion       As String
Dim motanula                        As String
Dim obs                             As String
Dim CIdValesCabRef                  As String
Dim User                            As String

    If MsgBox("Seguro de Anular el Vale", vbQuestion + vbYesNo, App.Title) = vbYes Then
        If indVale = "S" Then
            CIdValesCabRef = "" & traerCampo("ValesCab", "IdValesCabRef", "IdSucursal", glsSucursal, True, "TipoVale = 'S' And IdValesCab = '" & txtCod_Vale.Text & "'")
            If Len(Trim(CIdValesCabRef)) > 0 Then
                StrMsgError = "El Vale de Salida Nº " & txtCod_Vale.Text & " no se puede Anular porque ha sido generado del Vale de Ingreso Nº " & CIdValesCabRef & ", verifique.": GoTo Err
            End If
        End If
                
        IndEvaluacion = 0
    
        frmAprobacion.MostrarForm "05", IndEvaluacion, strCodUsuarioAutorizacion, StrMsgError
        If IndEvaluacion = 0 Then Exit Sub
        If StrMsgError <> "" Then GoTo Err
                    
        frmMotivosAnula.Motivos_Anulacion motanula, obs
        If motanula = 0 Then Exit Sub
        If StrMsgError <> "" Then GoTo Err
        
        Cn.BeginTrans
        iniTrans = True
        
        Graba_Logico_Vales "2", StrMsgError, glsSucursal, txtCod_Vale.Text, indVale
        If StrMsgError <> "" Then GoTo Err
        
        'actualizaStock txtCod_Vale.Text, 1, StrMsgError, indVale, False
        Actualiza_Stock_Nuevo StrMsgError, "E", glsSucursal, indVale, txtCod_Vale.Text, CIdAlmacenAnt
        If StrMsgError <> "" Then GoTo Err
      
        User = traerCampo("usuarios", "varusuario", "idusuario", glsUser, True)
        csql = "UPDATE valescab SET estValeCab = 'ANU',codanula='" & motanula & "',obsanulacion = '" & obs & "',fecanulacion=getdate(),usuanula='" & User & "' " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & txtCod_Vale.Text & "' " & _
               "AND tipoVale = '" & indVale & "' "
        Cn.Execute csql
        
        '--- Actualizamos la Cantidad importada
        csql = "SELECT idDocumentoImp,idDocVentasImp,idSerieImp,idProducto,idUM,Cantidad,idsucursalOrigen " & _
                 "FROM valesdet " & _
                 "WHERE idEmpresa = '" & glsEmpresa & "'" & _
                 "AND idSucursal = '" & glsSucursal & "'" & _
                 "AND idValesCab = '" & txtCod_Vale.Text & "'" & _
                 "AND idDocVentasImp <> '' AND tipoVale = '" & indVale & "' "
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            csql = "UPDATE docventasdet dd  SET dd.CantidadImp = dd.CantidadImp - " & CStr(rst.Fields("Cantidad")) & ", dd.estDocImportado = 'N' " & _
                   "WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idSucursal = '" & rst.Fields("idSucursalOrigen") & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                   "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "'"
            Cn.Execute csql
            rst.MoveNext
        Loop
        
        'Marcamos la Cabecera
       csql = "SELECT DISTINCT idDocumentoImp,idDocVentasImp,idSerieImp,idSucursalOrigen  FROM valesdet " & _
                "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & txtCod_Vale.Text & "' AND idDocVentasImp <> '' AND idDocVentasImp is not null " & _
                "AND tipoVale = '" & indVale & "' "
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        
        Do While Not rst.EOF
            csql = "UPDATE docVentas c  SET c.estDocImportado = 'N', c.estDocventas = 'GEN' " & _
                   "WHERE c.idEmpresa = '" & glsEmpresa & "' AND c.idSucursal = '" & rst.Fields("idSucursalOrigen") & "' AND c.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND c.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND c.idSerie = '" & rst.Fields("idSerieImp") & "'"
            Cn.Execute csql
            rst.MoveNext
        Loop
        
        If indVale = "I" Then
            CIdValesCabRef = "" & traerCampo("ValesCab", "IdValesCabRef", "IdSucursal", glsSucursal, True, "TipoVale = 'I' And IdValesCab = '" & txtCod_Vale.Text & "'")
            If Len(Trim(CIdValesCabRef)) > 0 Then
                csql = "UPDATE valescab SET estValeCab = 'ANU',codanula='" & motanula & "',obsanulacion = '" & obs & "',fecanulacion=getdate(),usuanula='" & User & "' " & _
                       "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & CIdValesCabRef & "' " & _
                       "AND tipoVale = 'S'"
                Cn.Execute csql
            End If
        End If
        
        Cn.CommitTrans
        
        strEstVale = "ANU"
        lbl_Anulado.Caption = "ANULADO"
        fraGeneral.Enabled = False
        habilitarColumas False
        habilitaBotones Button
        listaVales StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub
    
Err:
    If iniTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub habilitarColumas(ByVal indEstado As Boolean)
    
    gDetalle.Columns.ColumnByFieldName("idProducto").DisableEditor = Not indEstado
    gDetalle.Columns.ColumnByFieldName("Cantidad").DisableEditor = Not indEstado
    gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = Not indEstado
    gDetalle.Columns.ColumnByFieldName("PVUnit").DisableEditor = Not indEstado
    gDetalle.Columns.ColumnByFieldName("NumLote").DisableEditor = Not indEstado
    gDetalle.Columns.ColumnByFieldName("FecVencProd").DisableEditor = Not indEstado

End Sub

Private Sub mostrarDocImportado_Ayuda(ByVal rsdd As ADODB.Recordset, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg             As New ADODB.Recordset
Dim RsD             As New ADODB.Recordset
Dim strSerieDocVentas As String
Dim dblTC           As Double
Dim strCodFabri     As String
Dim strCodMar       As String
Dim strDesMar       As String
Dim intAfecto       As Integer
Dim strTipoProd     As String
Dim strMoneda       As String
Dim strCodUM        As String
Dim strDesUM        As String
Dim dblVVUnit       As Double
Dim dblIGVUnit      As Double
Dim dblPVUnit       As Double
Dim dblFactor       As Double
Dim intFila         As Integer
Dim i               As Integer
Dim indExisteDocRef As Boolean
Dim primero         As Boolean
Dim strInserta      As Boolean
Dim strFecIni       As String
    
    strFecIni = Format(dtp_Emision.Value, "yyyy-mm-dd")
    primero = True
    rsdd.MoveFirst
    Do While Not rsdd.EOF
        strInserta = True
        
        If traerCampo("Productos", "IdTipoProducto", "IdProducto", "" & rsdd.Fields("IdProducto"), True) = "06002" Then
            
            strInserta = False
            
        End If
        
        If strInserta = True Then
            If primero = True Then
                primero = False
            Else
                gDetalle.Dataset.Insert
            End If
        
            gDetalle.SetFocus
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" & rsdd.Fields("idProducto")
            gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = "" & rsdd.Fields("GlsProducto")
            gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = traerCampo("Productos", "CodigoRapido", "IDPRODUCTO", "" & rsdd.Fields("idProducto"), True)
            strCodUM = traerCampo("productos", "idUMCompra", "idProducto", "" & rsdd.Fields("idProducto"), True)
            If strDesUM = "" And strCodUM <> "" Then strDesUM = traerCampo("unidadMedida", "abreUM", "idUM", strCodUM, False)
            If Trim("" & rsdd.Fields("idProducto")) = "" Then Exit Sub
            
            If DatosProducto_Ayuda("" & rsdd.Fields("idProducto"), strCodUM, strDesUM, dblFactor) = False Then
            End If
            
            gDetalle.Columns.ColumnByFieldName("idUM").Value = strCodUM
            gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
            gDetalle.Columns.ColumnByFieldName("Afecto").Value = Val("" & rsdd.Fields("AfectoIGV"))
            gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
            gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
            gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Value = "0"
            
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = traerCostoUnit(Trim("" & rsdd.Fields("idProducto")), Trim("" & txtCod_Almacen.Text), strFecIni, txtCod_Moneda.Text, StrMsgError)
            If StrMsgError <> "" Then GoTo Err
            
            procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 0, Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value), dblVVUnit, dblIGVUnit, dblPVUnit
            
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            
            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), dblVVUnit, dblIGVUnit, dblPVUnit, Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
            
            gDetalle.Dataset.Post
            gDetalle.Dataset.RecNo = intFila
            gDetalle.Dataset.Edit
                            
            gDetalle.Dataset.Post
            
            If "" & rsdd.Fields("idProducto") <> "" Then
                gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").ColIndex
            End If
        End If
        rsdd.MoveNext
    Loop
    calcularTotales
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub mostrarDocImportado(ByVal rscd As ADODB.Recordset, ByVal rsdd As ADODB.Recordset, ByVal strTipoDocImportado As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim RsD As New ADODB.Recordset
Dim rsddtemp As New ADODB.Recordset
Dim strSerieDocVentas As String
Dim i As Integer
Dim indExisteDocRef As Boolean
Dim StrItem As String, strcodprod_aux As String, stridCodFabricante As String, strGlsproducto As String
Dim stridmarca As String, strGlsMarca As String, strIdDocventas As String, strIdSerie As String
Dim stridum As String, strglsum As String, strafecto As String, stridTipoProducto As String, stridMoneda As String
Dim strNumLote As String, strFecVencProd As String, stridSucursal As String
Dim nfactor As Double, ncantidad As Double, ncantidad2 As Double, NVVUnit As Double, NIGVUnit As Double, NPVUnit As Double
Dim nTotalVVBruto As Double, nTotalPVBruto As Double, nPorDcto As String, nDctoVV As Double, nDctoPV  As Double
Dim nTotalVVNeto As Double, nTotalIGVNeto As Double, nTotalPVNeto As Double
Dim nVVUnitLista As Double, nPVUnitLista As Double, nVVUnitNeto As Double, nPVUnitNeto As Double

    indCargando = True
    Set rsddtemp = rsdd
    i = 0

    'Formato Detalle
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adChar, 15, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Cantidad2", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 45, adFldIsNullable
    rsg.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idSucursalOrigen", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "idDocumentoImp", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "idDocVentasImp", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerieImp", adVarChar, 4, adFldIsNullable
    rsg.Fields.Append "idLote", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "IdTallaPeso", adVarChar, 30, adFldIsNullable
    rsg.Open

    '--- Formato Documento de Referencia
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adChar, 8, adFldIsNullable
    RsD.Open , , adOpenKeyset, adLockOptimistic

    If rsdd.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("idUM") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("Factor") = 1
        rsg.Fields("Afecto") = 1
        rsg.Fields("Cantidad") = 0
        rsg.Fields("Cantidad2") = 0
        rsg.Fields("VVUnit") = 0
        rsg.Fields("IGVUnit") = 0
        rsg.Fields("PVUnit") = 0
        rsg.Fields("TotalVVNeto") = 0
        rsg.Fields("TotalIGVNeto") = 0
        rsg.Fields("TotalPVNeto") = 0
        rsg.Fields("NumLote") = ""
        rsg.Fields("FecVencProd") = ""
        rsg.Fields("idDocumentoImp") = ""
        rsg.Fields("idDocVentasImp") = ""
        rsg.Fields("idSerieImp") = ""
        rsg.Fields("idSucursalOrigen") = ""
        rsg.Fields("idLote") = ""
        rsg.Fields("IdTallaPeso") = ""
        
    Else
    
        rscd.MoveFirst
        If Not rscd.EOF Then
            
            If Len(Trim("" & rscd.Fields("IdAlmacen"))) > 0 Then
                txtCod_Almacen.Text = "" & rscd.Fields("IdAlmacen")
            End If
            txtCod_Cliente.Text = "" & rscd.Fields("IdPersona")
            txtCod_Moneda.Text = "" & rscd.Fields("IdMoneda")
            txtObs.Text = "" & rscd.Fields("ObsDocVentas")
        End If
        
        rsdd.MoveFirst
        rsdd.Sort = "idProducto"
        'strSucursal_Origen = ""
        'strSucursal_Origen = RsDd.Fields("idSucursal")
        
        Do While Not rsdd.EOF
            strIdDocventas = "" & rsdd.Fields("idDocVentas")
            strIdSerie = "" & rsdd.Fields("idSerie")
            strcodprod_aux = rsdd.Fields("idProducto")
            stridCodFabricante = "" & rsdd.Fields("idCodFabricante")
            strGlsproducto = "" & rsdd.Fields("GlsProducto")
            stridmarca = "" & rsdd.Fields("idMarca")
            strGlsMarca = "" & rsdd.Fields("GlsMarca")
            stridum = "" & rsdd.Fields("idUM")
            strglsum = "" & rsdd.Fields("GlsUM")
            nfactor = "" & rsdd.Fields("Factor")
            strafecto = "" & rsdd.Fields("Afecto")
            ncantidad = "" & rsdd.Fields("Cantidad")
            ncantidad2 = "" & rsdd.Fields("Cantidad2")
            NVVUnit = "" & rsdd.Fields("VVUnit")
            NIGVUnit = "" & rsdd.Fields("IGVUnit")
            NPVUnit = "" & rsdd.Fields("PVUnit")
            nTotalVVBruto = "" & rsdd.Fields("TotalVVBruto")
            nTotalPVBruto = "" & rsdd.Fields("TotalPVBruto")
            nPorDcto = "" & rsdd.Fields("PorDcto")
            nDctoVV = "" & rsdd.Fields("DctoVV")
            nDctoPV = "" & rsdd.Fields("DctoPV")
            nTotalVVNeto = "" & rsdd.Fields("TotalVVNeto")
            nTotalIGVNeto = "" & rsdd.Fields("TotalIGVNeto")
            nTotalPVNeto = "" & rsdd.Fields("TotalPVNeto")
            stridTipoProducto = "" & rsdd.Fields("idTipoProducto")
            stridMoneda = "" & rsdd.Fields("idMoneda")
            strNumLote = "" & rsdd.Fields("NumLote")
            strFecVencProd = "" & rsdd.Fields("FecVencProd")
            nVVUnitLista = "" & rsdd.Fields("VVUnitLista")
            nPVUnitLista = "" & rsdd.Fields("PVUnitLista")
            nVVUnitNeto = "" & rsdd.Fields("VVUnitNeto")
            nPVUnitNeto = "" & rsdd.Fields("PVUnitNeto")
            stridSucursal = "" & rsdd.Fields("idSucursal")
            If strRepetirProductosGrid <> "S" Then
                ncantidad = 0#
                Do While rsdd.Fields("idProducto") & "" = strcodprod_aux And Not rsdd.EOF
                    ncantidad = ncantidad + rsdd.Fields("Cantidad")
                    strIdDocventas = "" & rsdd.Fields("idDocVentas")
                    strIdSerie = "" & rsdd.Fields("idSerie")
                    rsdd.MoveNext
                    If rsdd.EOF Then Exit Do
                    If rsdd.Fields("idProducto") <> strcodprod_aux Then Exit Do
                Loop
            Else
                rsdd.MoveNext
            End If

            rsg.AddNew
            i = i + 1
            rsg.Fields("Item") = i
            rsg.Fields("idProducto") = strcodprod_aux
            rsg.Fields("GlsProducto") = strGlsproducto
            rsg.Fields("idUM") = stridum
            rsg.Fields("GlsUM") = strglsum
            rsg.Fields("Factor") = nfactor
            rsg.Fields("Afecto") = strafecto
            rsg.Fields("Cantidad") = ncantidad
            rsg.Fields("IdTallaPeso") = Val(Format(ncantidad * Val("" & traerCampo("Productos", "IdTallaPeso", "IdProducto", strcodprod_aux, True)), "0.00"))
            rsg.Fields("Cantidad2") = ncantidad2
            rsg.Fields("VVUnit") = NVVUnit
            rsg.Fields("IGVUnit") = NIGVUnit
            rsg.Fields("PVUnit") = NPVUnit
            rsg.Fields("TotalVVNeto") = nTotalVVNeto
            rsg.Fields("TotalIGVNeto") = nTotalIGVNeto
            rsg.Fields("TotalPVNeto") = nTotalPVNeto
            rsg.Fields("NumLote") = strNumLote
            rsg.Fields("FecVencProd") = strFecVencProd
            rsg.Fields("idDocumentoImp") = "94"
            rsg.Fields("idDocVentasImp") = strIdDocventas
            rsg.Fields("idSerieImp") = strIdSerie
            rsg.Fields("IdSucursalOrigen") = stridSucursal
            rsg.Fields("idLote") = ""
        Loop
    End If

    rsddtemp.MoveFirst
    Do While Not rsddtemp.EOF
        If RsD.RecordCount > 0 Then RsD.MoveFirst
        indExisteDocRef = False
        Do While Not RsD.EOF
            If Trim("" & RsD.Fields("idDocumento")) = Trim("" & strTipoDocImportado) And Trim("" & RsD.Fields("idSerie")) = Trim("" & rsddtemp.Fields("idSerie")) And Trim("" & RsD.Fields("idNumDOc")) = Trim("" & rsddtemp.Fields("idDocVentas")) Then
                indExisteDocRef = True
                Exit Do
            End If
            RsD.MoveNext
        Loop

        If indExisteDocRef = False Then
            RsD.AddNew
            RsD.Fields("Item") = "" & RsD.RecordCount
            RsD.Fields("idDocumento") = strTipoDocImportado
            RsD.Fields("GlsDocumento") = traerCampo("documentos", "GlsDocumento", "idDocumento", strTipoDocImportado, False)
            RsD.Fields("idSerie") = "" & rsdd.Fields("idSerie")
            RsD.Fields("idNumDOc") = "" & rsdd.Fields("idDocVentas")
            'strSerieDocImportado = strIdSerie
            'strNumDocImportado = strIdDocVentas
        End If
        rsddtemp.MoveNext
    Loop

    rsdd.MoveFirst
    txtCod_Moneda.Text = traerCampo("docventas", "idMoneda", "iddocumento", strTipoDocImportado, True, " idSerie = '" & rsdd.Fields("idSerie") & "' and idDocVentas = '" & rsdd.Fields("idDocVentas") & "' ")
    dtp_Emision.Value = Format(traerCampo("docventas", "FecEmision", "iddocumento", strTipoDocImportado, True, "idSerie = '" & rsddtemp.Fields("idSerie") & "' and iddocventas = '" & rsddtemp.Fields("idDocVentas") & "' "), "DD/MM/YYYY")
    txtCod_CentroCosto.Text = Trim("" & traerCampo("docventas", "idCentroCosto", "iddocumento", strTipoDocImportado, True, " idSerie = '" & rsdd.Fields("idSerie") & "' and idDocVentas = '" & rsdd.Fields("idDocVentas") & "' "))
    
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    '--- Documentos de Referencia
    If RsD.RecordCount = 0 Then
        RsD.AddNew
        RsD.Fields("Item") = 1
        RsD.Fields("idDocumento") = ""
        RsD.Fields("GlsDocumento") = ""
        RsD.Fields("idSerie") = ""
        RsD.Fields("idNumDOc") = ""
    End If

    mostrarDatosGridSQL gDocReferencia, RsD, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    calcularTotales
    indCargando = False
    Me.Refresh

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Function traerCostoUnit(ByVal codproducto As String, ByVal codalmacen As String, ByVal PFecha As String, ByVal CodMoneda As String, ByRef StrMsgError As String) As Double
On Error GoTo Err
Dim CosUni  As ADODB.Recordset
Dim CanUni  As ADODB.Recordset
Dim bolCan As Boolean
    
    bolCan = False

    csql = "SELECT CAST(ISNULL(sum(CASE WHEN valescab.tipoVale = 'I' THEN valesdet.Cantidad ELSE valesdet.Cantidad * -1 END),0) AS NUMERIC(12,2)) AS Cantidad "
    csql = csql & "FROM valescab " & _
     "INNER JOIN valesdet  " & _
        "ON valescab.idValesCab = valesdet.idValesCab  " & _
        "AND valescab.idEmpresa = valesdet.idEmpresa  " & _
        "AND valescab.idSucursal = valesdet.idSucursal  " & _
        "AND valescab.tipoVale = valesdet.tipoVale " & _
      "INNER JOIN conceptos  " & _
        "ON valescab.idConcepto = conceptos.idConcepto  " & _
      "LEFT JOIN tiposdecambio t " & _
        "ON valescab.fechaEmision = t.fecha "
    csql = csql & "WHERE "
    csql = csql & "valescab.idEmpresa = '" & glsEmpresa & "' "
    csql = csql & "AND (valescab.idPeriodoInv) IN " & _
                    "(" & _
                        "SELECT pi.idPeriodoInv " & _
                        "FROM periodosinv pi " & _
                        "WHERE pi.idEmpresa = valescab.idEmpresa AND pi.idSucursal = valescab.idSucursal and CAST(pi.FecInicio AS DATE) <= CAST('" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "' AS DATE) " & _
                        "and (CAST(pi.FecFin AS DATE) >= CAST('" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "' AS DATE) or pi.FecFin is null)" & _
                    ") "
    csql = csql & " AND valescab.fechaEmision <= CAST('" & PFecha & "' AS DATE) And valesdet.idProducto = '" & codproducto & "' "
    csql = csql & "AND valescab.idAlmacen = '" & codalmacen & "' "
    csql = csql & "AND valescab.estValeCab <> 'ANU' "
    
    Set CanUni = New ADODB.Recordset
    CanUni.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not CanUni.EOF Then
        If CanUni("Cantidad") = 0 Then
            bolCan = True
        End If
    End If
    If CanUni.State = 1 Then CanUni.Close: Set CanUni = Nothing
    

    If Not bolCan Then
        csql = "Select(SUM(CASE WHEN valescab.tipoVale = 'I' THEN valesdet.Cantidad ELSE valesdet.Cantidad * -1 END * " & _
        "CASE '" & CodMoneda & "' " & _
        "WHEN 'PEN' THEN CASE WHEN valescab.idMoneda = 'PEN' THEN valesdet.VVUnit ELSE valesdet.VVUnit * ValesCab.TipoCambio END " & _
        "WHEN 'USD' THEN CASE WHEN valescab.idMoneda = 'USD' THEN valesdet.VVUnit ELSE valesdet.VVUnit / ValesCab.TipoCambio END " & _
        "End) / " & _
        "SUM(CASE WHEN valescab.tipoVale = 'I' THEN valesdet.Cantidad ELSE valesdet.Cantidad * -1 END)) AS COSTO_UNITARIO "
        csql = csql & "FROM valescab " & _
         "INNER JOIN valesdet  " & _
            "ON valescab.idValesCab = valesdet.idValesCab  " & _
            "AND valescab.idEmpresa = valesdet.idEmpresa  " & _
            "AND valescab.idSucursal = valesdet.idSucursal  " & _
            "AND valescab.tipoVale = valesdet.tipoVale " & _
          "INNER JOIN conceptos  " & _
            "ON valescab.idConcepto = conceptos.idConcepto  " & _
          "LEFT JOIN tiposdecambio t " & _
            "ON valescab.fechaEmision = t.fecha "
        csql = csql & "WHERE "
        csql = csql & "valescab.idEmpresa = '" & glsEmpresa & "' "
        csql = csql & "AND (valescab.idPeriodoInv) IN " & _
                        "(" & _
                            "SELECT pi.idPeriodoInv " & _
                            "FROM periodosinv pi " & _
                            "WHERE pi.idEmpresa = valescab.idEmpresa AND pi.idSucursal = valescab.idSucursal and CAST(pi.FecInicio AS DATE) <= CAST('" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "' AS DATE) " & _
                            "and (CAST(pi.FecFin AS DATE) >= CAST('" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "' AS DATE) or pi.FecFin is null)" & _
                        ") "
        csql = csql & " AND valescab.fechaEmision <= CAST('" & PFecha & "' AS DATE) And valesdet.idProducto = '" & codproducto & "' "
        csql = csql & "AND valescab.idAlmacen = '" & codalmacen & "' "
        csql = csql & "AND valescab.estValeCab <> 'ANU' "
        
    
        Set CosUni = New ADODB.Recordset
        CosUni.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not CosUni.EOF Then
           traerCostoUnit = IIf(IsNull(CosUni.Fields("COSTO_UNITARIO")), 0, CosUni.Fields("COSTO_UNITARIO"))
        End If
        If CosUni.State = 1 Then CosUni.Close: Set CosUni = Nothing
    Else
        traerCostoUnit = 0
    End If
    Exit Function
    
Err:
    If CosUni.State = 1 Then CosUni.Close: Set CosUni = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Function

Private Sub muestraColumnasDetalle()
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim pCtrl As Object
Dim strSerie As String
Dim intRegistros As Integer
    
    csql = "SELECT GlsObj,etiqueta,numCol,ancho,Tipodato,Decimales FROM objdocvales " & _
            "where idEmpresa = '" & glsEmpresa & "' " & _
            "and indVisible = 'V' " & _
            "ORDER BY NUMCOL"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly

    Do While Not rst.EOF
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Caption = rst.Fields("etiqueta") & ""
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").ColIndex = Val(rst.Fields("numCol") & "")
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Width = Val(rst.Fields("ancho") & "")
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Visible = True
        If (rst.Fields("Tipodato") & "") = "N" Then
            gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").DecimalPlaces = Val(rst.Fields("Decimales") & "")
        End If
        
        If Trim("" & rst.Fields("GlsObj")) = "IdAtributo" Then
            If indVale = "S" Then
                gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Visible = False
            End If
        End If
        
        rst.MoveNext
    Loop
    
    Exit Sub

Err:
    MsgBox Err.Description
End Sub

Private Sub EliminarVale(ByRef StrMsgError As String, ByRef Button As Integer)
On Error GoTo Err
Dim rst                             As New ADODB.Recordset
Dim iniTrans                        As Boolean
Dim IndEvaluacion                   As Integer
Dim strCodUsuarioAutorizacion       As String
Dim motanula                        As String
Dim obs                             As String
Dim CIdValesCabRef                  As String
Dim CSqlC                           As String

    If MsgBox("Seguro de Eliminar el Vale?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        If indVale = "S" Then
            CIdValesCabRef = "" & traerCampo("ValesCab", "IdValesCabRef", "IdSucursal", glsSucursal, True, "TipoVale = 'S' And IdValesCab = '" & txtCod_Vale.Text & "'")
            If Len(Trim(CIdValesCabRef)) > 0 Then
                StrMsgError = "El Vale de Salida Nº " & txtCod_Vale.Text & " no se puede Eliminar porque ha sido generado del Vale de Ingreso Nº " & CIdValesCabRef & ", verifique.": GoTo Err
            End If
        End If
                
        IndEvaluacion = 0
        frmAprobacion.MostrarForm "19", IndEvaluacion, strCodUsuarioAutorizacion, StrMsgError
        If IndEvaluacion = 0 Then Exit Sub
        If StrMsgError <> "" Then GoTo Err
                    
'''''        frmMotivosAnula.Motivos_Anulacion motanula, obs
'''''        If motanula = 0 Then Exit Sub
'''''        If strMsgError <> "" Then GoTo ERR
        
        Cn.BeginTrans
        iniTrans = True
        
        StrMsgError = ""
        Graba_Logico_Vales "3", StrMsgError, glsSucursal, txtCod_Vale.Text, indVale
        If StrMsgError <> "" Then GoTo Err
        
        'actualizaStock txtCod_Vale.Text, 1, StrMsgError, indVale, False
        Actualiza_Stock_Nuevo StrMsgError, "E", glsSucursal, indVale, txtCod_Vale.Text, CIdAlmacenAnt
        If StrMsgError <> "" Then GoTo Err
      
'''        User = traerCampo("usuarios", "varusuario", "idusuario", glsUser, True)
'''        CSqlC = "UPDATE valescab SET estValeCab = 'ANU',codanula='" & motanula & "',obsanulacion = '" & obs & "',fecanulacion=sysdate(),usuanula='" & User & "' " & _
'''               "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & txtCod_Vale.Text & "' " & _
'''               "AND tipoVale = '" & indVale & "' "
'''        Cn.Execute CSqlC
        
        '--- Actualizamos la Cantidad importada
        CSqlC = "SELECT idDocumentoImp,idDocVentasImp,idSerieImp,idProducto,idUM,Cantidad,idsucursalOrigen " & _
                 "FROM valesdet " & _
                 "WHERE idEmpresa = '" & glsEmpresa & "'" & _
                 "AND idSucursal = '" & glsSucursal & "'" & _
                 "AND idValesCab = '" & txtCod_Vale.Text & "'" & _
                 "AND idDocVentasImp <> '' AND tipoVale = '" & indVale & "' "
        If rst.State = 1 Then rst.Close
        rst.Open CSqlC, Cn, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            CSqlC = "UPDATE docventasdet dd  SET dd.CantidadImp = dd.CantidadImp - " & CStr(rst.Fields("Cantidad")) & ", dd.estDocImportado = 'N' " & _
                   "WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idSucursal = '" & rst.Fields("idSucursalOrigen") & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                   "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "'"
            Cn.Execute CSqlC
            rst.MoveNext
        Loop
        
        '--- Marcamos la Cabecera
       CSqlC = "SELECT DISTINCT idDocumentoImp,idDocVentasImp,idSerieImp,idSucursalOrigen  FROM valesdet " & _
                "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & txtCod_Vale.Text & "' AND idDocVentasImp <> '' AND ISNULL(idDocVentasImp) = False " & _
                "AND tipoVale = '" & indVale & "' "
        If rst.State = 1 Then rst.Close
        rst.Open CSqlC, Cn, adOpenForwardOnly, adLockReadOnly
        
        Do While Not rst.EOF
            CSqlC = "UPDATE docVentas c SET c.estDocImportado = 'N', c.estDocventas = 'GEN' " & _
                   "WHERE c.idEmpresa = '" & glsEmpresa & "' AND c.idSucursal = '" & rst.Fields("idSucursalOrigen") & "' AND c.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND c.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND c.idSerie = '" & rst.Fields("idSerieImp") & "'"
            Cn.Execute CSqlC
            rst.MoveNext
        Loop
        
        If indVale = "I" Then
            CIdValesCabRef = "" & traerCampo("ValesCab", "IdValesCabRef", "IdSucursal", glsSucursal, True, "TipoVale = 'I' And IdValesCab = '" & txtCod_Vale.Text & "'")
            If Len(Trim(CIdValesCabRef)) > 0 Then
                CSqlC = "DELETE from valescab " & _
                       "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & CIdValesCabRef & "' " & _
                       "AND tipoVale = 'S'"
                Cn.Execute CSqlC
                
                CSqlC = "DELETE from valesdet " & _
                       "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & CIdValesCabRef & "' " & _
                       "AND tipoVale = 'S'"
                Cn.Execute CSqlC
                
                CSqlC = "DELETE from valesdetLotes " & _
                       "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & CIdValesCabRef & "' " & _
                       "AND tipoVale = 'S'"
                       
                Cn.Execute CSqlC
                
            End If
        End If
        
        '--- ELIMINAMOS LOS VALES
        CSqlC = "DELETE from valescab " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & txtCod_Vale.Text & "' " & _
               "AND tipoVale = '" & indVale & "' "
        Cn.Execute CSqlC
        
        CSqlC = "DELETE from valesdet " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & txtCod_Vale.Text & "' " & _
               "AND tipoVale = '" & indVale & "' "
        Cn.Execute CSqlC
        
        CSqlC = "DELETE from valesdetLotes " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & txtCod_Vale.Text & "' " & _
               "AND tipoVale = '" & indVale & "' "
        Cn.Execute CSqlC
        
        CSqlC = "Delete A " & _
                "From ValesDetAtributos A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.TipoVale = '" & indVale & "' " & _
                "And A.IdValesCab = '" & txtCod_Vale.Text & "'"
                
        Cn.Execute CSqlC
        
        Cn.CommitTrans
        
        fraGeneral.Enabled = False
        habilitarColumas False
        habilitaBotones 9
        
    End If
    
    Exit Sub
Err:
    If iniTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub EjecutaSQLFormVales(F As Form, tipoOperacion As Integer, ByRef StrMsgError As String, strTipoVale As String, g As dxDBGrid, D As dxDBGrid, strFecha As String)
On Error GoTo Err
Dim rst                                 As New ADODB.Recordset
Dim C                                   As Object
Dim CSqlC                               As String
Dim strCampo                            As String
Dim strTipoDato                         As String
Dim strCampos                           As String
Dim strValores                          As String
Dim strValCod                           As String
Dim strCod                              As String
Dim indTrans                            As Boolean
Dim RsTempVale                          As New ADODB.Recordset
Dim CadMysqlTemp                        As String
Dim strSigno                            As String
Dim strcant                             As Integer
Dim cantReg                             As Integer
Dim codigoVale                          As String
Dim cComproPro                          As String
Dim ncantidad                           As Double
Dim ncantidadTotImp                     As Double
Dim strNewPeriodoinv                    As String
Dim i                                   As Long
Dim strCodAlmacen                       As String
Dim NCantImp                            As Double
Dim swProducto                          As String
Dim RsC                                 As New ADODB.Recordset
Dim rsl                                 As New ADODB.Recordset
Dim NCantidadLote                       As Double
Dim NCantidadProducto                   As Double
Dim NCantidadTotal                      As Double
Dim CIdLote                             As String
Dim CFechaSalidaLote                    As String
Dim valorGrid                           As String

    strNewPeriodoinv = ""
    
    If strTipoVale = "I" Then
        codigoVale = "88"
    Else
        codigoVale = "99"
    End If
    
    indTrans = False
    CSqlC = ""
    
    eliminaNulosGrilla
            
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then 'And C.Visible = True Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        If UCase(strCampo) = UCase("idValesCab") Then
                            If Trim(C.Value) = "" Then
                                strCod = generaCorrelativoAnoMes_ValeFecha("ValesCab", "idValesCab", strTipoVale, strFecha, True)
                                C.Text = strCod
                            Else
                                If traerCampo("ValesCab", "TidValesCab", "TidValesCab", Trim(C.Value), True) <> "" Then
                                    StrMsgError = "El Vale numero " & C.Value & " ya existe"
                                    GoTo Err
                                End If
                            End If
                            strValCod = Trim(C.Value)
                        End If
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & Val(C.Value) & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    Case 1
                        If UCase(strCampo) <> UCase("idValesCab") Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = Val(C.Value)
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
    
    indTrans = True
    Cn.BeginTrans
    
    'ACTUALIZAMOS STOCK EN LINEA
    'If tipoOperacion = 1 Then 'si es modificacion
    '    actualizaStock strValCod, 1, strMsgError, False
    '    If strMsgError <> "" Then GoTo ERR
    'End If
    
    
    If Len(Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", glsSucursal, True, " year(FecInicio) = " & Year(F.dtp_Emision.Value) & " "))) = 0 Then
        strNewPeriodoinv = glsCodPeriodoINV
    Else
        strNewPeriodoinv = Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", glsSucursal, True, " year(FecInicio) = " & Year(F.dtp_Emision.Value) & " "))
    End If
    
    
    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            CSqlC = "INSERT INTO ValesCab(" & strCampos & ",tipoVale,idEmpresa,idSucursal,idPeriodoInv, FechaRegistro) " & _
                    "VALUES(" & strValores & ",'" & strTipoVale & "','" & glsEmpresa & "','" & glsSucursal & "','" & strNewPeriodoinv & "',GETDATE())"
        Case 1
            
            Graba_Logico_Vales "1", StrMsgError, glsSucursal, strValCod, strTipoVale
            If StrMsgError <> "" Then GoTo Err
            
            CSqlC = "UPDATE ValesCab SET " & strCampos & ",FechaMod = GETDATE(),IdUsuarioMod = '" & glsUser & _
                    "' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & strValCod & "' AND tipoVale = '" & strTipoVale & "' "
    End Select
    
    '--- Graba controles
    Cn.Execute CSqlC
    
    CadMysqlTemp = "Select * From valesdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & strValCod & "' and tipoVale = '" & strTipoVale & "' "
    If RsTempVale.State = 1 Then RsTempVale.Close
    RsTempVale.Open CadMysqlTemp, Cn, adOpenStatic, adLockReadOnly
    
    '--- Grabando Grilla detalle
    If TypeName(g) <> "Nothing" Then
        
        If tipoOperacion = "1" Then
        
            CSqlC = "Update B Set B.CantidadImp = B.CantidadImp - A.Cantidad,B.EstDocImportado = 'N' FROM ValesDet A " & _
                   "Inner Join DocVentasDet B " & _
                       "On A.IdEmpresa = B.IdEmpresa And A.IdSucursalOrigen = B.IdSucursal And A.IdDocumentoImp = B.IdDocumento " & _
                       "And A.IdDocVentasImp = B.IdDocVentas And A.IdSerieImp = B.IdSerie And A.IdProducto = B.IdProducto And A.IdUM = B.IdUM " & _
                   " " & _
                   "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdValesCab = '" & strValCod & "' " & _
                   "And A.IdDocVentasImp <> '' And A.TipoVale = '" & strTipoVale & "'"
            Cn.Execute CSqlC
        
            CSqlC = "Update B Set B.EstDocImportado = 'N',B.EstDocventas = 'GEN' FROM ValesDet A " & _
                   "Inner Join DocVentas B " & _
                       "On A.IdEmpresa = B.IdEmpresa And A.IdSucursalOrigen = B.IdSucursal And A.IdDocumentoImp = B.IdDocumento " & _
                       "And A.IdDocVentasImp = B.IdDocVentas And A.IdSerieImp = B.IdSerie " & _
                   " " & _
                   "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdValesCab = '" & strValCod & "' " & _
                   "And A.IdDocVentasImp <> '' And A.TipoVale = '" & strTipoVale & "'"
            Cn.Execute CSqlC
        
        End If
        
        Actualiza_Stock_Nuevo StrMsgError, "E", glsSucursal, strTipoVale, strValCod, CIdAlmacenAnt
        If StrMsgError <> "" Then GoTo Err
    
        Cn.Execute "DELETE FROM valesdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & strValCod & "' and tipoVale = '" & strTipoVale & "' "
        
        g.Dataset.First
        Do While Not g.Dataset.EOF
        
            If glsValidaStock = True Then
                
                If indVale = "S" Then
                    'ValidaStock StrMsgError, Trim("" & txtCod_Almacen.Text), Trim("" & G.Columns.ColumnByFieldName("idProducto").Value), Val(Format("" & G.Columns.ColumnByFieldName("Cantidad").Value, "0.00")), Trim("" & txtCod_Vale.Text)
                    'If StrMsgError <> "" Then
                        
                        If tipoOperacion = 0 Then
                            txtCod_Vale.Text = ""
                        End If
                        
                        GoTo Err
                    'End If
                End If
            
            End If
                
            strCampos = ""
            strValores = ""
            For i = 0 To g.Columns.Count - 1
                    If UCase(left(g.Columns(i).ObjectName, 1)) = "W" Then
                        strTipoDato = Mid(g.Columns(i).ObjectName, 2, 1)
                        strCampo = Mid(g.Columns(i).ObjectName, 3)
                        
                        strCampos = strCampos & strCampo & ","
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & Val(g.Columns(i).Value) & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                        End Select
    
                    End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            CSqlC = "INSERT INTO valesdet(" & strCampos & ",idValesCab,idEmpresa,idSucursal,tipoVale) VALUES(" & strValores & ",'" & strValCod & "','" & glsEmpresa & "','" & glsSucursal & "','" & strTipoVale & "')"
                    
            Cn.Execute CSqlC
            
            g.Dataset.Next
        Loop
    End If
    
    Actualiza_Stock_Nuevo StrMsgError, "I", glsSucursal, strTipoVale, strValCod, txtCod_Almacen.Text
    If StrMsgError <> "" Then GoTo Err
    
    CSqlC = "Delete A " & _
            "From ValesDetAtributos A " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.TipoVale = '" & strTipoVale & "' " & _
            "And A.IdValesCab = '" & strValCod & "'"
    
    Cn.Execute CSqlC
    
    If RsDetAtributos.RecordCount > 0 Then
        
        RsDetAtributos.MoveFirst
        
        Do While Not RsDetAtributos.EOF
            
            CSqlC = "Insert Into ValesDetAtributos(IdEmpresa,IdSucursal,TipoVale,IdValesCab,Item,IdAtributo,Valor)Values(" & _
                    "'" & glsEmpresa & "','" & glsSucursal & "','" & strTipoVale & "','" & strValCod & "'," & Val("" & RsDetAtributos.Fields("Item")) & "," & _
                    "'" & Trim("" & RsDetAtributos.Fields("IdAtributo")) & "'," & Val("" & RsDetAtributos.Fields("Valor")) & ")"
            
            Cn.Execute CSqlC
            
            RsDetAtributos.MoveNext
            
        Loop
        
    End If
        
    '--- Grabando Grilla DocRef
    If TypeName(D) <> "Nothing" Then
        Cn.Execute "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = '" & IIf(strTipoVale = "I", "88", "99") & "' AND numDocOrigen = '" & strValCod & "' AND serieDocOrigen = '000'"
        If D.Count > 0 Then
            D.Dataset.First
            Do While Not D.Dataset.EOF
                If Trim(D.Columns.ColumnByFieldName("idDocumento").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idSerie").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idNumDoc").Value) <> "" Then
                    strCampos = ""
                    strValores = ""
                    For i = 0 To D.Columns.Count - 1
                            If UCase(left(D.Columns(i).ObjectName, 1)) = "W" Then
                                strTipoDato = Mid(D.Columns(i).ObjectName, 2, 1)
                                strCampo = Mid(D.Columns(i).ObjectName, 3)
                                
                                strCampos = strCampos & strCampo & ","
                                
                                Select Case strTipoDato
                                    Case "N"
                                        strValores = strValores & D.Columns(i).Value & ","
                                    Case "T"
                                        strValores = strValores & "'" & Trim(D.Columns(i).Value) & "',"
                                    Case "F"
                                        strValores = strValores & "'" & Format(D.Columns(i).Value, "yyyy-mm-dd") & "',"
                                End Select
            
                            End If
                    Next
                    
                    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
                    If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
                    
                    CSqlC = "INSERT INTO docreferencia(" & strCampos & ",tipoDocOrigen,numDocOrigen,serieDocOrigen,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & codigoVale & "','" & strValCod & "','000','" & glsEmpresa & "','" & glsSucursal & "')"
                    Cn.Execute CSqlC
                End If
                D.Dataset.Next
            Loop
        End If
    End If
    
'    If tipoOperacion = "1" Then
'         If Not RsTempVale.EOF Then
'            RsTempVale.MoveFirst
'            Do While Not RsTempVale.EOF
'
'                strSigno = "-"
'                If strTipoVale = "S" Then strSigno = "+"
'                strCodAlmacen = traerCampo("valescab", "idalmacen", "idvalescab", Trim("" & RsTempVale.Fields("idvalescab")), True, " tipoVale = '" & strTipoVale & "' ")
'
'                CSqlC = "UPDATE productosalmacen " & _
'                        "SET CantidadStock = CantidadStock " & strSigno & " " & RsTempVale.Fields("Cantidad") & " " & _
'                        "WHERE idEmpresa = '" & RsTempVale.Fields("idEmpresa") & "' " & _
'                          "AND idSucursal = '" & RsTempVale.Fields("idSucursal") & "' " & _
'                          "AND idAlmacen = '" & strCodAlmacen & "' " & _
'                          "AND idProducto = '" & RsTempVale.Fields("idProducto") & "' " & _
'                          "AND idUMCompra = '" & RsTempVale.Fields("idUM") & "' "
'                Cn.Execute CSqlC
'
'                CSqlC = "UPDATE productosalmacenporlote " & _
'                        "SET CantidadStock = CantidadStock " & strSigno & " " & RsTempVale.Fields("Cantidad") & " " & _
'                        "WHERE idEmpresa = '" & RsTempVale.Fields("idEmpresa") & "' " & _
'                          "AND idSucursal = '" & RsTempVale.Fields("idSucursal") & "' " & _
'                          "AND idAlmacen = '" & strCodAlmacen & "' " & _
'                          "AND idProducto = '" & RsTempVale.Fields("idProducto") & "' " & _
'                          "AND idUMCompra = '" & RsTempVale.Fields("idUM") & "' " & _
'                          "AND idLote = '" & RsTempVale.Fields("idLote") & "' "
'                Cn.Execute CSqlC
'
'                RsTempVale.MoveNext
'            Loop
'         End If
'
'    End If
        
    '--- ACTUALIZAMOS STOCK EN LINEA
    'actualizaStock strValCod, 0, strMsgError, strTipoVale, False
    'If strMsgError <> "" Then GoTo Err
    
    'actualizaStock_Lote strValCod, 0, StrMsgError, strTipoVale, False
    'If StrMsgError <> "" Then GoTo Err
        
    
    '--- Actualizamos cantidad importada
    CSqlC = "SELECT idDocumentoImp,idDocVentasImp,idSerieImp,idProducto,idUM,Cantidad,idsucursalOrigen " & _
            "FROM valesdet " & _
            "WHERE idEmpresa = '" & glsEmpresa & "'" & _
            "AND idSucursal = '" & glsSucursal & "'" & _
            "AND idValesCab = '" & strValCod & "' " & _
            "AND tipoVale = '" & strTipoVale & "' " & _
            "AND idDocVentasImp <> '' "
    
    If rst.State = 1 Then rst.Close
    rst.Open CSqlC, Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        'PQS 140616 ya no filtra sucursal AND dd.idSucursal = '" & glsSucursal & "'
        'ACTUALIZAMOS CANTIDAD IMPORTADA
        CSqlC = "UPDATE dd  SET dd.CantidadImp = dd.CantidadImp + " & CStr(rst.Fields("Cantidad")) & " " & _
               "FROM docventasdet dd WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                 "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "'"
        Cn.Execute CSqlC
    
        CSqlC = "UPDATE dd  SET dd.estDocImportado = 'S' " & _
               "FROM docventasdet dd WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                 "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "' AND dd.CantidadImp >= dd.Cantidad"
        Cn.Execute CSqlC
    
        rst.MoveNext
    Loop
    
    '--- Marcamos la Cabecera
    CSqlC = "SELECT DISTINCT idDocumentoImp,idDocVentasImp,idSerieImp  FROM valesdet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & strValCod & "' " & _
           " AND tipoVale = '" & strTipoVale & "' " & _
           " AND idDocVentasImp <> '' AND idDocVentasImp IS NOT NULL "
    
    If rst.State = 1 Then rst.Close
    rst.Open CSqlC, Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        CSqlC = "UPDATE c SET c.estDocImportado = 'S' from docVentas c inner join docVentasDet d   " & _
               "ON c.idEmpresa = d.idEmpresa AND c.idSucursal = d.idSucursal AND c.idDocumento = d.idDocumento AND c.idDocVentas = d.idDocVentas AND c.idSerie = d.idSerie " & _
               "WHERE d.idEmpresa = '" & glsEmpresa & "' AND d.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND d.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND d.idSerie = '" & rst.Fields("idSerieImp") & "'" & _
               "AND (SELECT count(x.idDocVentas) FROM docVentasDet x WHERE x.idEmpresa = d.idEmpresa AND x.idSucursal = d.idSucursal AND x.idDocumento = d.idDocumento AND x.idDocVentas = d.idDocVentas AND x.idSerie = d.idSerie AND x.estDocImportado <> 'S') = 0"
        Cn.Execute CSqlC
    
        rst.MoveNext
    Loop
    
    '--- Actualizamos el estado de  la compra PAR(Parcial),CER(Cerrado)
    If D.Columns.ColumnByFieldName("idDocumento").Value = "94" Or D.Columns.ColumnByFieldName("idDocumento").Value = "86" Then
    strcant = 0
    cantReg = 0
        If TypeName(D) <> "Nothing" Then
            D.Dataset.First
            Do While Not D.Dataset.EOF
             strcant = 0
             cantReg = 0
                If Trim(D.Columns.ColumnByFieldName("idDocumento").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idSerie").Value) <> "" And Trim(D.Columns.ColumnByFieldName("idNumDoc").Value) <> "" Then
           
                    CSqlC = "SELECT  d.idDocumentoImp,d.idDocVentasImp,d.idSerieImp,d.idProducto,d.idUM,d.Cantidad,d.estDocImportado,d.CantidadImp " & _
                            "FROM  docVentas c Inner Join docVentasDet d " & _
                            "On c.idempresa = d.idempresa AND c.idSucursal = d.idSucursal AND c.idDocumento = d.idDocumento AND c.idDocVentas = d.idDocVentas " & _
                            "WHERE d.idEmpresa = '" & glsEmpresa & "' AND d.idDocumento = '" & D.Columns.ColumnByFieldName("idDocumento").Value & "'" & _
                            "AND d.idDocVentas = '" & D.Columns.ColumnByFieldName("idDocventas").Value & "'" & _
                            "AND d.idSerie = '" & D.Columns.ColumnByFieldName("idSerie").Value & "'  "
                            
                    If rst.State = 1 Then rst.Close
                    rst.Open CSqlC, Cn, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rst.EOF Then
                        rst.MoveFirst
                        rst.Sort = "idProducto"
                        Do While Not rst.EOF
                        
                            cComproPro = rst.Fields("idProducto").Value
                            Do While cComproPro = rst.Fields("idProducto").Value & ""
                                ncantidad = ncantidad + rst.Fields("Cantidad").Value
                                NCantImp = rst.Fields("CantidadImp").Value
                                  
                                rst.MoveNext
                                If rst.EOF Then Exit Do
                                If cComproPro <> rst.Fields("idProducto").Value & "" Then Exit Do
                            Loop
                            ncantidadTotImp = ncantidadTotImp + NCantImp
                        Loop
                        
                        If D.Columns.ColumnByFieldName("idDocumento").Value = "94" Then
                            If ncantidad = ncantidadTotImp Then
                                CSqlC = "UPDATE  d Set estDocVentas='ING' FROM docVentas d WHERE d.idEmpresa = '" & glsEmpresa & "' " & _
                                       "AND d.idDocumento = '" & D.Columns.ColumnByFieldName("idDocumento").Value & "' AND d.idDocVentas = '" & D.Columns.ColumnByFieldName("idDocventas").Value & "' AND d.idSerie = '" & D.Columns.ColumnByFieldName("idSerie").Value & "' "
                            Else
                                CSqlC = "UPDATE d Set estDocVentas='PAR' FROM docVentas d WHERE d.idEmpresa = '" & glsEmpresa & "' AND d.idDocumento = '" & D.Columns.ColumnByFieldName("idDocumento").Value & "'" & _
                                       "AND d.idDocVentas = '" & D.Columns.ColumnByFieldName("idDocventas").Value & "' AND d.idSerie = '" & D.Columns.ColumnByFieldName("idSerie").Value & "' "
                            End If
                        End If
                        If D.Columns.ColumnByFieldName("idDocumento").Value = "86" Then
                            CSqlC = "UPDATE d Set estDocVentas='CER', estDocImportado='S' FROM docVentas d WHERE d.idEmpresa = '" & glsEmpresa & "' AND d.idDocumento = '" & D.Columns.ColumnByFieldName("idDocumento").Value & "' " & _
                                   "AND d.idDocVentas = '" & D.Columns.ColumnByFieldName("idDocventas").Value & "' AND d.idSerie = '" & D.Columns.ColumnByFieldName("idSerie").Value & "' "
                        End If
                        Cn.Execute CSqlC
                        ncantidad = 0
                        'strcantImp = 0
                    End If
                 End If
                 D.Dataset.Next
              Loop
        End If
     End If
    
    gDetalle.Dataset.First
    If Not gDetalle.Dataset.EOF Then
        Do While Not gDetalle.Dataset.EOF
            swProducto = ""
            swProducto = traerCampo("productosalmacen", "idProducto", "idProducto", Trim("" & gDetalle.Columns.ColumnByFieldName("idProducto").Value), True, "idAlmacen = '" & txtCod_Almacen.Text & "' and idSucursal = '" & glsSucursal & "' ")
            If swProducto = "" Then
                CSqlC = "insert into productosalmacen(idAlmacen, idProducto, item, idEmpresa, " & _
                        "idSucursal, idUMCompra, CantidadStock, CostoUnit) " & _
                        "values('" & txtCod_Almacen.Text & "', '" & Trim("" & gDetalle.Columns.ColumnByFieldName("idProducto").Value) & "',0, '" & glsEmpresa & "', " & _
                        "'" & glsSucursal & "','" & Trim("" & gDetalle.Columns.ColumnByFieldName("idUM").Value) & "', 0, 0)"
                Cn.Execute CSqlC
            End If
            gDetalle.Dataset.Next
        Loop
    End If
    
    If leeParametro("STOCK_POR_LOTE") = "S" Then
    
        CSqlC = "Delete From ValesDetLotes " & _
                "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And TipoVale = '" & strTipoVale & "' " & _
                "And IdValesCab = '" & strValCod & "'"
        
        Cn.Execute CSqlC
        
        If strTipoVale = "S" Then
        
            CSqlC = "Select A.* " & _
                    "From ValesDet A " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.TipoVale = '" & strTipoVale & "' " & _
                    "And A.IdValesCab = '" & strValCod & "' " & _
                    "Order By A.Item"
            RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
            Do While Not RsC.EOF
                
                    NCantidadLote = 0
                    NCantidadTotal = 0
                    NCantidadProducto = Val("" & RsC.Fields("Cantidad"))
                            
                    CSqlC = "Select B.IdLote,Sum(B.Cantidad * CASE WHEN B.TipoVale = 'I' THEN 1 ELSE -1 END) CantidadLote " & _
                            "From ValesCab A " & _
                            "Inner Join ValesDetLotes B " & _
                                "On A.IdEmpresa = B.IdEmpresa And A.IdSucursal = B.IdSucursal And A.TipoVale = B.TipoVale And A.IdValesCab = B.IdValesCab " & _
                            "Inner Join Lotes C " & _
                                "On B.IdEmpresa = C.IdEmpresa And B.IdLote = C.IdLote " & _
                            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdPeriodoInv = '" & strNewPeriodoinv & "' " & _
                            "And A.IdAlmacen = '" & txtCod_Almacen.Text & "' And A.EstValeCab <> 'ANU' " & _
                            "And CAST(A.FechaEmision AS DATE) <= CAST('" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "' AS DATE) And B.IdProducto = '" & Trim("" & RsC.Fields("IdProducto")) & "' " & _
                            "And C.IdLote = '" & Val("" & RsC.Fields("idLote")) & "' " & _
                            "Group By B.IdLote " & _
                            ""
                            
                    rsl.Open CSqlC, Cn, adOpenKeyset, adLockReadOnly
                    If Not rsl.EOF Then
                    
                        If Val("" & rsl.Fields("CantidadLote")) < (NCantidadProducto) Then
                            StrMsgError = "El producto " & Trim("" & RsC.Fields("IdProducto")) & " no cuenta con stock para La Talla " & Trim("" & RsC.Fields("NumLote")) & ",Verifique."
                            txtCod_Vale.Text = ""
                            GoTo Err
                        End If

                        CSqlC = "Insert Into ValesDetLotes(IdEmpresa,IdSucursal,TipoVale,IdValesCab,Item,IdLote,IdProducto,IdUM,Cantidad,CantidadAnt)Values(" & _
                                "'" & glsEmpresa & "','" & glsSucursal & "','" & strTipoVale & "','" & strValCod & "'," & Val("" & RsC.Fields("Item")) & "," & _
                                "'" & Trim("" & rsl.Fields("IdLote")) & "','" & Trim("" & RsC.Fields("IdProducto")) & "','" & Trim("" & RsC.Fields("IdUM")) & "'," & _
                                "" & NCantidadProducto & ",0)"
                        
                        Cn.Execute CSqlC
                        rsl.Close: Set rsl = Nothing
                    
                    End If

                
                RsC.MoveNext
            Loop
        
            RsC.Close: Set RsC = Nothing
        
        Else
        
            CSqlC = "Select A.* " & _
                    "From ValesDet A " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.TipoVale = '" & strTipoVale & "' " & _
                    "And A.IdValesCab = '" & strValCod & "' " & _
                    "Order By A.Item"
            RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
            If Not RsC.EOF Then
                
                   'CIdLote = ("" & traerCampo("Lotes", "IdLote", "TipoVale", strTipoVale, True, "idLote = '" & strValCod & "'"))
                
'''                If CIdLote = "" Then
'''
'''                    CIdLote = GeneraCorrelativoAnoMes("Lotes", "IdLote", True)
'''
'''                    CSqlC = "Insert Into Lotes(IdLote,GlsLote,Estado,IdEmpresa,IdSucursal,TipoVale,IdValesCab,FechaLote)Values(" & _
'''                            "'" & CIdLote & "','" & "Lote " & CIdLote & "','ACT','" & glsEmpresa & "','" & glsSucursal & "','" & strTipoVale & "'," & _
'''                            "'" & strValCod & "','" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "')"
'''
'''                Else
'''
'''                    CFechaSalidaLote = traerCampo("ValesCab A Inner Join ValesDetLotes B " & _
'''                    "On A.IdEmpresa = B.IdEmpresa And A.IdSucursal = B.IdSucursal And A.TipoVale = B.TipoVale And A.IdValesCab = B.IdValesCab", "A.FechaEmision", "A.TipoVale", "S", False, "A.EstValeCab <> 'ANU' And A.IdEmpresa = '" & glsEmpresa & "' And B.IdLote = '" & CIdLote & "' Order By A.FechaEmision Limit 1")
'''
'''                    If CFechaSalidaLote <> "" Then
'''
'''                        If CVDate(dtp_Emision.Value) > CVDate(CFechaSalidaLote) Then
'''
'''                            StrMsgError = "La primera salida del Lote " & CIdLote & " con fecha " & CFechaSalidaLote & " no puede ser menor a la de Ingreso.": GoTo Err
'''
'''                        End If
'''
'''                    End If
'''
'''                    CSqlC = "Update Lotes " & _
'''                            "Set FechaLote = '" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "' " & _
'''                            "Where IdEmpresa = '" & glsEmpresa & "' And IdLote = '" & CIdLote & "'"
'''
'''                End If
'''
'''                Cn.Execute CSqlC
                
                Do While Not RsC.EOF
                
                    CSqlC = "Insert Into ValesDetLotes(IdEmpresa,IdSucursal,TipoVale,IdValesCab,Item,IdLote,IdProducto,IdUM,Cantidad,CantidadAnt)Values(" & _
                            "'" & glsEmpresa & "','" & glsSucursal & "','" & strTipoVale & "','" & strValCod & "'," & Val("" & RsC.Fields("Item")) & "," & _
                            "'" & Val("" & RsC.Fields("idLote")) & "','" & Trim("" & RsC.Fields("IdProducto")) & "','" & Trim("" & RsC.Fields("IdUM")) & "'," & _
                            "" & Val("" & RsC.Fields("Cantidad")) & ",0)"
                    
                    Cn.Execute CSqlC
                        
                    RsC.MoveNext
                    
                Loop
            
            End If
            
            RsC.Close: Set RsC = Nothing
        
        End If
    
    End If
    
    Cn.CommitTrans
    
    CIdAlmacenAnt = txtCod_Almacen.Text
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
    Exit Sub
End Sub

Private Sub Inicia_RecordSet(PRs As ADODB.Recordset, PRsIndex As Integer, StrMsgError As String)
On Error GoTo Err
    
    With PRs
        Select Case PRsIndex
            Case 0
                If .State = 1 Then .Close
                .Fields.Append "Item", adInteger, 10, adFldIsNullable
                .Fields.Append "IdAtributo", adVarChar, 8, adFldIsNullable
                .Fields.Append "GlsAtributo", adVarChar, 250, adFldIsNullable
                .Fields.Append "Valor", adDouble, 14, adFldIsNullable
                .Open
        End Select
    End With
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
End Sub

Private Sub txtCod_Usuario_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_usuario.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_usuario.Text, False)

Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ValidaStock(StrMsgError As String, PIdAlmacen As String, PIdProducto As String, PCantidad As Double, PValeCab As String)
On Error GoTo Err
Dim RsC                     As New ADODB.Recordset
Dim CSqlC                   As String
    
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    
    CSqlC = " SELECT Format(ifnull(XZ.sc_stock,0) + ifnull(s.Stock,0),2) as Stock " & _
    "FROM productos p INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
    "INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM " & _
    "INNER JOIN monedas o ON p.idMoneda = o.idMoneda " & _
    "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
    "Left Join (Select P.idEmpresa,IfNull(vd.idSucursal,'') Idsucursal,P.idProducto,vc.idAlmacen, " & _
    "sum(If(vd.idempresa is null,0,if(vd.tipovale = 'I',Cantidad,Cantidad * -1))) as Stock " & _
    "From Productos P Inner Join ValesDet vd On P.IdEmpresa = vd.IdEmpresa And P.IdProducto = vd.IdProducto " & _
    "Inner Join Valescab vc On vd.idEmpresa = vc.idEmpresa And vd.idSucursal = vc.idSucursal " & _
    "And vd.tipoVale = vc.tipoVale And vd.idValesCab = vc.idValesCab " & _
    "Where P.idEmpresa = '" & glsEmpresa & "' AND estProducto = 'A' and vc.Idvalescab <> '" & PValeCab & "' " & _
    "AND vc.idSucursal = '" & glsSucursal & "' And vc.estValeCab <> 'ANU' " & _
    "AND vc.idAlmacen = '" & PIdAlmacen & "' AND DATE_FORMAT(vc.fechaemision, '%Y%m%d')  = DATE_FORMAT(sysdate(), '%Y%m%d') AND (p.idProducto = '" & PIdProducto & "') " & _
    "Group bY P.idEmpresa,P.idProducto,vc.idAlmacen) S " & _
    "On P.idEmpresa = S.idEmpresa And P.idProducto = S.idProducto " & _
    "Left Join (SELECT sc_periodo,sc_codalm,sc_codart,sc_stock,idempresa FROM tbsaldo_costo_kardex z " & _
    "where sc_codalm = '" & PIdAlmacen & "' and sc_periodo = DATE_FORMAT(sysdate(), '%Y%m') and sc_stock <> 0) XZ " & _
    "On P.idEmpresa  = xz.idempresa And P.idProducto = xz.sc_codart " & _
    "Where p.idEmpresa = '" & glsEmpresa & "' AND (p.idProducto = '" & PIdProducto & "') AND estProducto = 'A' "
    
    RsC.Open CSqlC, Cn, adOpenKeyset, adLockReadOnly
    If Not RsC.EOF Then
    
        If Val(Format("" & RsC.Fields("Stock"), "0.00")) < Val(Format(PCantidad, "0.00")) Then
            StrMsgError = StrMsgError & " La Cantidad Ingresada para el Producto " & " " & PIdProducto & " Exede el stock Verifique" & "  " & Chr(13) & Chr(10)
        End If
    
    End If
    
    RsC.Close: Set RsC = Nothing
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

