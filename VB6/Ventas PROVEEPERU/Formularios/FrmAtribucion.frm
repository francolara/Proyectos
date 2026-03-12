VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmAtribucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atribuciones"
   ClientHeight    =   7575
   ClientLeft      =   3255
   ClientTop       =   2220
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11280
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   7155
      Top             =   30
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
            Picture         =   "FrmAtribucion.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtribucion.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
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
   Begin VB.Frame FraRegistro 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6900
      Left            =   0
      TabIndex        =   7
      Top             =   645
      Width           =   11265
      Begin VB.Frame Frame2 
         Height          =   3120
         Index           =   0
         Left            =   2340
         TabIndex        =   40
         Top             =   2085
         Visible         =   0   'False
         Width           =   6150
         Begin VB.TextBox CATTextBox1 
            Height          =   2250
            Index           =   0
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   180
            Width           =   5970
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Aceptar"
            Height          =   390
            Index           =   0
            Left            =   1725
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   2535
            Width           =   1140
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancelar"
            Height          =   390
            Index           =   0
            Left            =   3165
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   2520
            Width           =   1140
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   120
         TabIndex        =   29
         Top             =   5970
         Width           =   11025
         Begin CATControls.CATTextBox Txt_ImpSubTotal 
            Height          =   315
            Left            =   3825
            TabIndex        =   30
            Tag             =   "NImpSubTotal"
            Top             =   285
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontBold        =   -1  'True
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmAtribucion.frx":3518
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox Txt_ImpIGV 
            Height          =   315
            Left            =   6390
            TabIndex        =   31
            Tag             =   "NImpIGV"
            Top             =   285
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontBold        =   -1  'True
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmAtribucion.frx":3534
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox Txt_ImpTotal 
            Height          =   315
            Left            =   9150
            TabIndex        =   32
            Tag             =   "NImpTotal"
            Top             =   285
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontBold        =   -1  'True
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmAtribucion.frx":3550
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label lbl_TotalBruto 
            Appearance      =   0  'Flat
            Caption         =   "Sub Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   2415
            TabIndex        =   38
            Top             =   330
            Width           =   915
         End
         Begin VB.Label lbl_TotalIGV 
            Appearance      =   0  'Flat
            Caption         =   "IGV:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   5625
            TabIndex        =   37
            Top             =   330
            Width           =   330
         End
         Begin VB.Label lbl_TotalNeto 
            Appearance      =   0  'Flat
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   8160
            TabIndex        =   36
            Top             =   330
            Width           =   540
         End
         Begin VB.Label lbl_SimbMonNeto 
            Appearance      =   0  'Flat
            Caption         =   "S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   8760
            TabIndex        =   35
            Top             =   330
            Width           =   330
         End
         Begin VB.Label lbl_SimbMonIGV 
            Appearance      =   0  'Flat
            Caption         =   "S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   6030
            TabIndex        =   34
            Top             =   330
            Width           =   330
         End
         Begin VB.Label lbl_SimbMonBruto 
            Appearance      =   0  'Flat
            Caption         =   "S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   3480
            TabIndex        =   33
            Top             =   330
            Width           =   285
         End
      End
      Begin VB.Frame FraDetalle 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3765
         Left            =   120
         TabIndex        =   9
         Top             =   2205
         Width           =   11025
         Begin DXDBGRIDLibCtl.dxDBGrid GDetalle 
            Height          =   3450
            Left            =   90
            OleObjectBlob   =   "FrmAtribucion.frx":356C
            TabIndex        =   6
            Top             =   165
            Width           =   10845
         End
      End
      Begin VB.Frame FraGeneral 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2085
         Left            =   105
         TabIndex        =   10
         Top             =   120
         Width           =   11070
         Begin VB.CommandButton cmbAyudaMoneda 
            Height          =   315
            Left            =   6105
            Picture         =   "FrmAtribucion.frx":638F
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1110
            Width           =   390
         End
         Begin VB.TextBox Txt_IdProveedor 
            Appearance      =   0  'Flat
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
            Left            =   1065
            TabIndex        =   2
            Tag             =   "TIdProveedor"
            Top             =   675
            Width           =   825
         End
         Begin VB.TextBox Txt_GlsProveedor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   1965
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   675
            Width           =   8415
         End
         Begin VB.CommandButton Cmd_Proveedor 
            Height          =   315
            Left            =   10425
            Picture         =   "FrmAtribucion.frx":6719
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   675
            Width           =   390
         End
         Begin VB.TextBox Txt_IdSerieAtri 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1065
            TabIndex        =   0
            Tag             =   "TIdSerieAtri"
            Top             =   240
            Width           =   825
         End
         Begin VB.TextBox Txt_IdDocAtri 
            Appearance      =   0  'Flat
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
            Left            =   3090
            TabIndex        =   1
            Tag             =   "TIdDocAtri"
            Top             =   240
            Width           =   960
         End
         Begin MSComCtl2.DTPicker Dtp_FechaEmision 
            Height          =   330
            Left            =   1065
            TabIndex        =   4
            Tag             =   "FFechaEmision"
            Top             =   1545
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
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
            Format          =   42795009
            CurrentDate     =   40718
         End
         Begin CATControls.CATTextBox txtCod_Moneda 
            Height          =   315
            Left            =   1065
            TabIndex        =   3
            Tag             =   "TidMoneda"
            Top             =   1110
            Width           =   825
            _ExtentX        =   1455
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
            Container       =   "FrmAtribucion.frx":6AA3
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Moneda 
            Height          =   315
            Left            =   1965
            TabIndex        =   18
            Top             =   1110
            Width           =   4095
            _ExtentX        =   7223
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
            Container       =   "FrmAtribucion.frx":6ABF
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox Txt_TipoCambio 
            Height          =   315
            Left            =   3765
            TabIndex        =   5
            Tag             =   "NTipoCambio"
            Top             =   1560
            Width           =   825
            _ExtentX        =   1455
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
            Container       =   "FrmAtribucion.frx":6ADB
            Text            =   "0.000"
            Decimales       =   3
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "T.C."
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   2865
            TabIndex        =   39
            Top             =   1635
            Width           =   300
         End
         Begin VB.Label lbl_Moneda 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   165
            TabIndex        =   19
            Top             =   1185
            Width           =   600
         End
         Begin VB.Label Label16 
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
            Height          =   210
            Left            =   165
            TabIndex        =   16
            Top             =   1590
            Width           =   450
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   15
            Top             =   765
            Width           =   480
         End
         Begin VB.Label Label15 
            Caption         =   "Número"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2250
            TabIndex        =   12
            Top             =   285
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Serie"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   11
            Top             =   285
            Width           =   645
         End
      End
   End
   Begin VB.Frame FraLista 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6900
      Left            =   0
      TabIndex        =   20
      Top             =   645
      Width           =   11265
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   75
         TabIndex        =   21
         Top             =   120
         Width           =   11055
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
            ItemData        =   "FrmAtribucion.frx":6AF7
            Left            =   7605
            List            =   "FrmAtribucion.frx":6B1F
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   240
            Width           =   1620
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1185
            TabIndex        =   23
            Top             =   240
            Width           =   5745
            _ExtentX        =   10134
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
            Container       =   "FrmAtribucion.frx":6B88
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Ano 
            Height          =   315
            Left            =   10035
            TabIndex        =   24
            Top             =   255
            Width           =   870
            _ExtentX        =   1535
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
            Container       =   "FrmAtribucion.frx":6BA4
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label25 
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
            Top             =   280
            Width           =   735
         End
         Begin VB.Label Label24 
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
            Left            =   7110
            TabIndex        =   26
            Top             =   280
            Width           =   300
         End
         Begin VB.Label Label26 
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
            Left            =   9495
            TabIndex        =   25
            Top             =   280
            Width           =   300
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid GLista 
         Height          =   5805
         Left            =   75
         OleObjectBlob   =   "FrmAtribucion.frx":6BC0
         TabIndex        =   28
         Top             =   960
         Width           =   11070
      End
   End
End
Attribute VB_Name = "FrmAtribucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim indBoton                                As Integer
Dim RsDet                                   As New ADODB.Recordset
Dim SwF2                                    As Boolean

Private Sub cbx_Mes_Click()
On Error GoTo Err
Dim StrMsgError As String
        
    ListaAtribuciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaMoneda_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Cmd_Proveedor_Click()
On Error GoTo Err
Dim StrMsgError As String
    
    mostrarAyuda "PROVEEDOR", Txt_IdProveedor, Txt_GlsProveedor
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Command1_Click(Index As Integer)
Dim StrMsgError As String
On Error GoTo Err

    
    GDetalle.Dataset.Edit
    GDetalle.Columns.ColumnByFieldName("glsProducto").Value = CATTextBox1(0).Text
    GDetalle.Dataset.Post
    Frame2(0).Visible = False
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, "Ventas"
End Sub

Private Sub Command2_Click(Index As Integer)
    
    Frame2(0).Visible = False
    
End Sub

Private Sub Dtp_FechaEmision_Change()
Dim StrMsgError                             As String
On Error GoTo Err
    
    Txt_TipoCambio.Text = Val("" & traerCampo("TiposDeCambio", "TcVenta", "Fecha", Format(Dtp_FechaEmision.Value, "yyyy-mm-dd"), False))
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Dtp_FechaEmision_KeyDown(KeyCode As Integer, Shift As Integer)
Dim StrMsgError                             As String
On Error GoTo Err
    
    If KeyCode = 13 Then
        
        Txt_TipoCambio.SetFocus
        
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim StrMsgError                             As String
On Error GoTo Err
    
    Me.Width = 11520
    Me.Height = 8145
    Me.top = 0
    Me.left = 0
    
    txt_Ano.Text = Year(getFechaSistema)
    cbx_Mes.ListIndex = Month(getFechaSistema) - 1
    
    SwF2 = True
    
    habilitaBotones 7, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    ConfGrid GDetalle, True, True, False, False
    ConfGrid GDetalle, False, True, False, False
    
    ListaAtribuciones StrMsgError
    If StrMsgError <> "" Then GoTo Err

    FraLista.Visible = True
    FraRegistro.Visible = False
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ListaAtribuciones(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond                                 As String
    
    strCond = ""
    
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = "%" & Trim(txt_TextoBuscar.Text) & "%"
    End If
    
    With GLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = "Call Spu_ListaAtribuciones('" & glsEmpresa & "','" & Val(txt_Ano.Text) & "'," & cbx_Mes.ListIndex + 1 & ",'" & strCond & "')"
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "Item"
    End With
    
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    
End Sub

Private Sub habilitaBotones(indexBoton As Integer, StrMsgError As String)
Dim indHabilitar As Boolean
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
            Toolbar1.Buttons(6).Visible = True 'Imprimir
            Toolbar1.Buttons(7).Visible = True 'Lista
        Case 3 'Modificar
            Toolbar1.Buttons(1).Visible = False 'Nuevo
            Toolbar1.Buttons(2).Visible = True 'Grabar
            Toolbar1.Buttons(3).Visible = False 'Modificar
            Toolbar1.Buttons(4).Visible = True 'Cancelar
            Toolbar1.Buttons(5).Visible = False 'Eliminar
            Toolbar1.Buttons(6).Visible = False 'Imprimir
            Toolbar1.Buttons(7).Visible = False 'Lista
        Case 4 'Cancelar
            
            If indBoton = 1 Then
            
                Toolbar1.Buttons(1).Visible = True 'Nuevo
                Toolbar1.Buttons(2).Visible = False 'Grabar
                Toolbar1.Buttons(3).Visible = True 'Modificar
                Toolbar1.Buttons(4).Visible = False 'Cancelar
                Toolbar1.Buttons(5).Visible = True 'Eliminar
                Toolbar1.Buttons(6).Visible = True 'Imprimir
                Toolbar1.Buttons(7).Visible = True 'Lista
            
            Else
            
                Toolbar1.Buttons(1).Visible = True
                Toolbar1.Buttons(2).Visible = False
                Toolbar1.Buttons(3).Visible = False
                Toolbar1.Buttons(4).Visible = False
                Toolbar1.Buttons(5).Visible = False
                Toolbar1.Buttons(6).Visible = False
                Toolbar1.Buttons(7).Visible = False
            
            End If
            
        Case 7 'Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
    End Select
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    
End Sub

Private Sub gdetalle_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim StrMsgError                     As String
Dim X                               As Double
On Error GoTo Err

    If Action = daInsert Then
        With GDetalle
            
            '.Dataset.Edit
            .Columns.ColumnByFieldName("Item").Value = .Count
            .Columns.ColumnByFieldName("IdProducto").Value = " "
            .Columns.ColumnByFieldName("GlsProducto").Value = ""
            .Columns.ColumnByFieldName("ImpConteo").Value = 0
            .Columns.ColumnByFieldName("PorcOP").Value = 0
            .Columns.ColumnByFieldName("PorcOV").Value = 0
            .Columns.ColumnByFieldName("ImpAtribucion").Value = 0
            .Dataset.Post
            
            .Columns.FocusedIndex = GDetalle.Columns.ColumnByFieldName("IdProducto").Index
            
        End With
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gdetalle_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
Dim StrMsgError                     As String
On Error GoTo Err

    If Action = daInsert Then
        With GDetalle
            If .Columns.ColumnByFieldName("GlsProducto").Value = "" Then
                Allow = False
            End If
        End With
    End If
    
    Exit Sub
Err:
    Allow = False
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gDetalle_OnDblClick()
Dim StrMsgError As String
Dim strGlsProOrigen As String
On Error GoTo Err

    Frame2(0).Visible = True
    CATTextBox1(0).Visible = True
    CATTextBox1(0).SetFocus
    
    CATTextBox1(0).Text = GDetalle.Columns.ColumnByFieldName("glsProducto").Value
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, "Ventas"
End Sub

Private Sub gdetalle_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrMsgError                 As String
Dim CCod                        As String
Dim CDes                        As String
On Error GoTo Err

    With GDetalle

        Select Case .Columns.FocusedColumn.Index
            Case .Columns.ColumnByFieldName("IdProducto").Index
                
                mostrarAyudaTexto "PRODUCTOS", CCod, CDes, "And IdTipoProducto = '06002'"
                If StrMsgError <> "" Then GoTo Err
                
                If Len(Trim(CCod)) > 0 Then
                    
                    .Dataset.Edit
                    .Columns.ColumnByFieldName("IdProducto").Value = CCod
                    .Columns.ColumnByFieldName("GlsProducto").Value = CDes
                    .Dataset.Post
                    
                End If
                
                .SetFocus
                
        End Select
        
    End With
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gDetalle_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrMsgError                 As String
Dim NImpAtribucion              As String
On Error GoTo Err
    
    With GDetalle
        If .Dataset.Modified = False Then Exit Sub

        Select Case .Columns.FocusedColumn.Index
            Case .Columns.ColumnByFieldName("IdProducto").Index
                
                If Len(Trim(.Columns.ColumnByFieldName("IdProducto").Value)) > 0 Then
                        
                    .Columns.ColumnByFieldName("IdProducto").Value = traerCampo("Productos", "GlsProducto", "IdProducto", .Columns.ColumnByFieldName("IdProducto").Value, True, "IdTipoProducto = '06002'")
                    
                Else
                    
                    .Columns.ColumnByFieldName("IdProducto").Value = ""
                
                End If
            
            Case .Columns.ColumnByFieldName("ImpAtribucion").Index
                
                
                
            Case Else
            
                .Columns.ColumnByFieldName("ImpAtribucion").Value = Val("" & .Columns.ColumnByFieldName("ImpConteo").Value) * (Val("" & .Columns.ColumnByFieldName("PorcOP").Value) / 100)
                .Dataset.Post
                
                If Val("" & .Columns.ColumnByFieldName("ImpAtribucion").Value) > 0 Then
                
                    NImpAtribucion = Val("" & .Columns.ColumnByFieldName("ImpConteo").Value)
                    
                    '.Dataset.Post
                    .Dataset.Insert
                    
                    .Dataset.Edit
                    .Columns.ColumnByFieldName("ImpAtribucion").Value = (((NImpAtribucion - (NImpAtribucion * 0.02)) * 0.12) / 2) * -1
                    .Dataset.Post
                
                End If
                
                CalculaTotales StrMsgError
                If StrMsgError <> "" Then GoTo Err
                
        End Select
        
    End With
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gdetalle_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim StrMsgError As String
Dim intFila As Integer
Dim i       As Integer
On Error GoTo Err

    intFila = GDetalle.Dataset.RecNo
    intFila = GDetalle.Dataset.RecNo
    intFila = GDetalle.Dataset.RecNo

    If KeyCode = 46 Then
        If GDetalle.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                
                If GDetalle.Count = 1 Then
                    
                    GDetalle.Dataset.Edit
                    
                    GDetalle.Columns.ColumnByFieldName("Item").Value = 1
                    GDetalle.Columns.ColumnByFieldName("IdProducto").Value = " "
                    GDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
                    GDetalle.Columns.ColumnByFieldName("ImpConteo").Value = 0
                    GDetalle.Columns.ColumnByFieldName("PorcOP").Value = 0
                    GDetalle.Columns.ColumnByFieldName("PorcOV").Value = 0
                    GDetalle.Columns.ColumnByFieldName("ImpAtribucion").Value = 0
                    
                    GDetalle.Dataset.Post
                
                Else
                
                    GDetalle.Dataset.Delete
                    GDetalle.Dataset.First
                    
                    Do While Not GDetalle.Dataset.EOF
                        i = i + 1
                        
                        GDetalle.Dataset.Edit
                        GDetalle.Columns.ColumnByFieldName("Item").Value = i
                        GDetalle.Dataset.Post
                        
                        GDetalle.Dataset.Next
                    Loop
                    
                    If GDetalle.Dataset.State = dsEdit Or GDetalle.Dataset.State = dsInsert Then
                        GDetalle.Dataset.Post
                    End If
                
                End If
                
                CalculaTotales StrMsgError
                If StrMsgError <> "" Then GoTo Err
                
                GDetalle.SetFocus
                GDetalle.Dataset.RecNo = intFila
                
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If GDetalle.Dataset.State = dsEdit Or GDetalle.Dataset.State = dsInsert Then
              GDetalle.Dataset.Post
        End If
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub GDetalle_OnKeyUp(KeyCode As Integer, ByVal Shift As Long)
On Error GoTo Err
Dim StrMsgError As String

    If KeyCode = 113 And SwF2 = True Then
        gdetalle_OnEditButtonClick GDetalle.Columns.FocusedColumn, Nothing
        SwF2 = False
    Else
        SwF2 = True
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    MostrarAtribuciones GLista.Columns.ColumnByName("IdSerieAtri").Value, GLista.Columns.ColumnByName("IdDocAtri").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    FraLista.Visible = False
    FraRegistro.Visible = True
    FraRegistro.Enabled = False
    
    habilitaBotones 2, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub MostrarAtribuciones(PIdSerieAtri As String, PIdDocAtri As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst                         As New ADODB.Recordset
Dim CSqlC                       As String
    
    CSqlC = "Select A.IdSerieAtri,A.IdDocAtri,A.IdProveedor,A.IdMoneda,A.FechaEmision,A.TipoCambio,A.ImpSubTotal,A.ImpIGV,A.ImpTotal " & _
            "From AtribucionesCab A " & _
            "Where IdEmpresa = '" & glsEmpresa & "' And A.IdSerieAtri = '" & PIdSerieAtri & "' And A.IdDocAtri = '" & PIdDocAtri & "'"
    
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If RsDet.State = 1 Then RsDet.Close: Set RsDet = Nothing
    
    RsDet.Fields.Append "Item", adInteger, , adFldIsNullable
    RsDet.Fields.Append "IdProducto", adVarChar, 8, adFldRowID
    RsDet.Fields.Append "GlsProducto", adVarChar, 185, adFldIsNullable
    RsDet.Fields.Append "ImpConteo", adDouble, , adFldIsNullable
    RsDet.Fields.Append "PorcOP", adDouble, , adFldIsNullable
    RsDet.Fields.Append "PorcOV", adDouble, , adFldIsNullable
    RsDet.Fields.Append "ImpAtribucion", adDouble, , adFldIsNullable
    RsDet.Open
        
    CSqlC = "Select A.Item,A.IdProducto,A.GlsProducto,A.ImpConteo,A.PorcOP,A.PorcOV,A.ImpAtribucion " & _
            "From AtribucionesDet A " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSerieAtri = '" & PIdSerieAtri & "' And A.IdDocAtri = '" & PIdDocAtri & "'"
            
    With rst
        .Open CSqlC, Cn, adOpenStatic, adLockReadOnly
        If .EOF Then
            
            RsDet.AddNew
            RsDet.Fields("Item") = 1
            RsDet.Fields("IdProducto") = ""
            RsDet.Fields("GlsProducto") = ""
            RsDet.Fields("ImpConteo") = 0
            RsDet.Fields("PorcOP") = 0
            RsDet.Fields("PorcOV") = 0
            RsDet.Fields("ImpAtribucion") = 0
        
        Else
            
            Do While Not .EOF
                
                RsDet.AddNew
                RsDet.Fields("Item") = Val("" & .Fields("Item"))
                RsDet.Fields("IdProducto") = "" & .Fields("IdProducto")
                RsDet.Fields("GlsProducto") = "" & .Fields("GlsProducto")
                RsDet.Fields("ImpConteo") = Val("" & .Fields("ImpConteo"))
                RsDet.Fields("PorcOP") = Val("" & .Fields("PorcOP"))
                RsDet.Fields("PorcOV") = Val("" & .Fields("PorcOV"))
                RsDet.Fields("ImpAtribucion") = Val("" & .Fields("ImpAtribucion"))
                
                .MoveNext
                
            Loop
            
        End If
        
        .Close: Set rst = Nothing
        
    End With
        
    Set GDetalle.DataSource = Nothing
             
    mostrarDatosGridSQL2 GDetalle, RsDet, "Item", StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    GDetalle.Columns.FocusedIndex = GDetalle.Columns.ColumnByFieldName("IdProducto").Index
             
    Me.Refresh
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            indBoton = 0
            
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            FraLista.Visible = False
            FraRegistro.Visible = True
            FraRegistro.Enabled = True
            
        Case 2 'Grabar
            Validaciones StrMsgError, "Grabar"
            If StrMsgError <> "" Then GoTo Err
            
            Grabar StrMsgError, indBoton
            If StrMsgError <> "" Then GoTo Err
            
        Case 3 'Modificar
            indBoton = 1
            FraRegistro.Enabled = True
            Txt_IdSerieAtri.Locked = True
            Txt_IdDocAtri.Locked = True
    
        Case 4 'Cancelar
            
            If indBoton = 0 Then
            
                FraLista.Visible = True
                FraRegistro.Visible = False
                FraRegistro.Enabled = False
            
            Else
            
                FraRegistro.Enabled = False
            
            End If
            
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        Case 6 'Imprimir
            
            imprimir StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        Case 7 'Lista
            
            FraLista.Visible = True
            FraRegistro.Visible = False
            
        Case 8 'Salir
            Unload Me
            
    End Select
    
    habilitaBotones Button.Index, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String, IndGraba As Integer)
On Error GoTo Err
Dim CIdDocAtri                          As String
Dim strMsg                              As String
Dim CSqlC                               As String
Dim RsDetClone                          As New ADODB.Recordset
Dim indTrans                            As Boolean
Dim CFecSystem                          As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    CFecSystem = getFechaHoraSistema
    
    Cn.BeginTrans
    indTrans = True
    
    If IndGraba = 0 Then
        
        If Len(Trim(Txt_IdDocAtri.Text)) = 0 Then
            
            CIdDocAtri = generaCorrelativo("AtribucionesCab", "IdDocAtri", 8, , True, "IdSerieAtri = '" & Txt_IdSerieAtri.Text & "'")
        
        Else
        
            CIdDocAtri = Txt_IdDocAtri.Text
            
        End If
        
        CSqlC = "Insert Into AtribucionesCab(IdEmpresa,IdSerieAtri,IdDocAtri,IdProveedor,IdMoneda,FechaEmision,TipoCambio,ImpSubTotal,ImpIGV," & _
                "ImpTotal,FechaRegistro,HoraRegistro,IdUsuarioRegistro)Values" & _
                "('" & glsEmpresa & "','" & Txt_IdSerieAtri.Text & "','" & CIdDocAtri & "','" & Txt_IdProveedor.Text & "'," & _
                "'" & txtCod_Moneda.Text & "','" & Format(Dtp_FechaEmision.Value, "yyyy-mm-dd") & "'," & Val(Format(Txt_TipoCambio.Text, "0.000")) & "," & _
                "" & Val(Format(Txt_ImpSubTotal.Text, "0.00")) & "," & Val(Format(Txt_ImpIGV.Text, "0.00")) & "," & _
                "" & Val(Format(Txt_ImpTotal.Text, "0.00")) & ",'" & Format(CFecSystem, "yyyy-mm-dd") & "','" & Format(CFecSystem, "h:mm:ss") & "'," & _
                "'" & glsUser & "')"
        
        Cn.Execute CSqlC
        
        strMsg = "Grabo"
        
    Else
    
        CIdDocAtri = Txt_IdDocAtri.Text
        
        CSqlC = "Update AtribucionesCab " & _
                "Set IdProveedor = '" & Txt_IdProveedor.Text & "',IdMoneda = '" & txtCod_Moneda.Text & "'," & _
                "FechaEmision = '" & Format(Dtp_FechaEmision.Value, "yyyy-mm-dd") & "',TipoCambio = " & Val(Format(Txt_TipoCambio.Text, "0.000")) & "," & _
                "ImpSubTotal = " & Val(Format(Txt_ImpSubTotal.Text, "0.00")) & ",ImpIGV = " & Val(Format(Txt_ImpIGV.Text, "0.00")) & "," & _
                "ImpTotal = " & Val(Format(Txt_ImpTotal.Text, "0.00")) & ",FechaModificacion = '" & Format(CFecSystem, "yyyy-mm-dd") & "'," & _
                "HoraModificacion = '" & Format(CFecSystem, "h:mm:ss") & "',IdUsuarioModificacion = '" & glsUser & "' " & _
                "Where IdEmpresa = '" & glsEmpresa & "' And IdSerieAtri = '" & Txt_IdSerieAtri.Text & "' And IdDocAtri = '" & CIdDocAtri & "'"
        
        Cn.Execute CSqlC
        
        strMsg = "Modifico"
        
        CSqlC = "Delete A From AtribucionesDet A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSerieAtri = '" & Txt_IdSerieAtri.Text & "' " & _
                "And A.IdDocAtri = '" & Txt_IdDocAtri.Text & "'"
    
        Cn.Execute CSqlC
        
    End If
    
    If GDetalle.Dataset.State = dsEdit Then GDetalle.Dataset.Post
    
'    RsDet.Filter = "GlsProducto = ''"
'
'    Do While Not RsDet.EOF
'
'        RsDet.Delete
'
'        RsDet.MoveNext
'
'    Loop
'
'    GDetalle.Dataset.Refresh
                
    Set RsDetClone = RsDet.Clone(adLockReadOnly)
    
    With RsDetClone
    
        If Not .EOF Then
            
            .MoveFirst
            
            CSqlC = ""
            
            Do While Not .EOF
                    
                CSqlC = CSqlC & "('" & glsEmpresa & "','" & Txt_IdSerieAtri.Text & "','" & CIdDocAtri & "'," & .Fields("Item") & "," & _
                        "'" & .Fields("IdProducto") & "','" & .Fields("GlsProducto") & "'," & Val("" & .Fields("ImpConteo")) & "," & Val("" & .Fields("PorcOP")) & "," & _
                        "" & Val("" & .Fields("PorcOV")) & "," & Val("" & .Fields("ImpAtribucion")) & "),"
                
                .MoveNext
                
            Loop
            
            If Len(Trim(CSqlC)) > 0 Then
                
                CSqlC = "Insert Into AtribucionesDet(IdEmpresa,IdSerieAtri,IdDocAtri,Item,IdProducto,GlsProducto,ImpConteo,PorcOP,PorcOV,ImpAtribucion)Values" & left(CSqlC, Len(CSqlC) - 1)
                
                Cn.Execute CSqlC
            
            End If
            
        End If
        
        .Close: Set RsDetClone = Nothing
    
    End With
    
    Cn.CommitTrans
    indTrans = False
    
    Txt_IdDocAtri.Text = CIdDocAtri
    
    Txt_IdSerieAtri.Locked = True
    Txt_IdDocAtri.Locked = True
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    FraRegistro.Enabled = False
    
    ListaAtribuciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans        As Boolean
Dim rsValida        As New ADODB.Recordset
Dim CSqlC           As String

    If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
     
    Cn.BeginTrans
    indTrans = True
 
    'Eliminando el registro
    CSqlC = "Delete A,B " & _
            "From AtribucionesCab A " & _
            "Inner Join AtribucionesDet B " & _
                "On A.IdEmpresa = B.IdEmpresa And A.IdSerieAtri = B.IdSerieAtri And A.IdDocAtri = B.IdDocAtri " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSerieAtri = '" & Txt_IdSerieAtri.Text & "' " & _
            "And A.IdDocAtri = '" & Txt_IdDocAtri.Text & "'"
    
    Cn.Execute CSqlC

    Cn.CommitTrans
    
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    ListaAtribuciones StrMsgError
    
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub
    
Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_Ano_Change()
On Error GoTo Err
Dim StrMsgError As String
        
    ListaAtribuciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdDocAtri_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    If KeyAscii = 13 Then
        
        Txt_IdProveedor.SetFocus
    
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdDocAtri_LostFocus()
On Error GoTo Err
Dim StrMsgError As String
    
    Txt_IdDocAtri.Text = Format(Txt_IdDocAtri.Text, "00000000")
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProveedor_Change()
On Error GoTo Err
Dim StrMsgError As String

    If Len(Trim(Txt_IdProveedor.Text)) = 0 Then
        
        Txt_GlsProveedor.Text = ""
    
    Else
    
        Txt_GlsProveedor.Text = traerCampo("Personas", "GlsPersona", "IdPersona", Txt_IdProveedor.Text, False)
    
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyCode = 113 Then
        
        Cmd_Proveedor_Click
        
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdProveedor_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    If KeyAscii = 13 Then
        
        txtCod_Moneda.SetFocus
    
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdSerieAtri_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String
    
    KeyAscii = ControlaKey(StrMsgError, KeyAscii, "T")
    If StrMsgError <> "" Then GoTo Err
    
    If KeyAscii = 13 Then
        
        Txt_IdDocAtri.SetFocus
    
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Txt_IdSerieAtri_LostFocus()
On Error GoTo Err
Dim StrMsgError As String

    Txt_IdSerieAtri.Text = Format(Txt_IdSerieAtri.Text, "000")
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    ListaAtribuciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Configura_Recordset(StrMsgError As String, PIndNuevo As Boolean)
On Error GoTo Err
        
    If RsDet.State = 1 Then RsDet.Close: Set RsDet = Nothing
     
    RsDet.Fields.Append "Item", adInteger, , adFldIsNullable
    RsDet.Fields.Append "IdProducto", adVarChar, 8, adFldRowID
    RsDet.Fields.Append "GlsProducto", adVarChar, 185, adFldIsNullable
    RsDet.Fields.Append "ImpConteo", adDouble, , adFldIsNullable
    RsDet.Fields.Append "PorcOP", adDouble, , adFldIsNullable
    RsDet.Fields.Append "PorcOV", adDouble, , adFldIsNullable
    RsDet.Fields.Append "ImpAtribucion", adDouble, , adFldIsNullable
    RsDet.Open
    
    If PIndNuevo Then
    
        RsDet.AddNew
        RsDet.Fields("Item") = 1
        RsDet.Fields("IdProducto") = ""
        RsDet.Fields("GlsProducto") = ""
        RsDet.Fields("ImpConteo") = 0
        RsDet.Fields("PorcOP") = 0
        RsDet.Fields("PorcOV") = 0
        RsDet.Fields("ImpAtribucion") = 0
        
        mostrarDatosGridSQL2 GDetalle, RsDet, "Item", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        GDetalle.Columns.FocusedIndex = GDetalle.Columns.ColumnByFieldName("IdProducto").Index
    
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo(StrMsgError As String)
Dim strAno                  As String
On Error GoTo Err
    
    strAno = txt_Ano.Text
    
    limpiaForm Me
    
    txt_Ano.Text = strAno

    Txt_IdSerieAtri.Text = traerCampo("AtribucionesCab", "Max(IdSerieAtri)", "IdEmpresa", glsEmpresa, False)
    
    Txt_TipoCambio.Text = Val("" & traerCampo("TiposDeCambio", "TcVenta", "Fecha", Format(Dtp_FechaEmision.Value, "yyyy-mm-dd"), False))
    
    If Txt_IdSerieAtri.Text = "" Then
        
        Txt_IdSerieAtri.Text = "001"
        
    End If
    
    Configura_Recordset StrMsgError, True
    If StrMsgError <> "" Then GoTo Err
    
    Txt_IdSerieAtri.Locked = False
    Txt_IdDocAtri.Locked = False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub CalculaTotales(StrMsgError As String)
Dim RsDetClone                  As New ADODB.Recordset
On Error GoTo Err
    
    Set RsDetClone = RsDet.Clone(adLockReadOnly)
    
    With RsDetClone
        
        If Not .EOF Then
            
            Txt_ImpSubTotal.Text = "0"
            Txt_ImpTotal.Text = "0"
            
            .MoveFirst
            
            Do While Not .EOF
                
                Txt_ImpSubTotal.Text = Val(Format(Txt_ImpSubTotal.Text, "0.00")) + Val("" & .Fields("ImpAtribucion"))
                
                .MoveNext
                
            Loop
            
            Txt_ImpTotal.Text = Val(Format(Txt_ImpSubTotal.Text, "0.00"))
            
        End If
        
        .Close: Set RsDetClone = Nothing
        
    End With
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Txt_TipoCambio_KeyPress(KeyAscii As Integer)
Dim StrMsgError                             As String
On Error GoTo Err
    
    If KeyAscii = 13 Then
        
        GDetalle.SetFocus
        GDetalle.Columns.FocusedIndex = GDetalle.Columns.ColumnByFieldName("IdProducto").Index
        
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TipoCambio_LostFocus()
Dim StrMsgError                             As String
On Error GoTo Err

    Txt_TipoCambio.Text = Val(Format(Txt_TipoCambio.Text, "0.000"))
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Moneda_Change()
Dim StrMsgError                             As String
On Error GoTo Err

    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
    
    lbl_SimbMonBruto.Caption = traerCampo("monedas", "Simbolo", "idMoneda", txtCod_Moneda.Text, False)
    lbl_SimbMonIGV.Caption = lbl_SimbMonBruto.Caption
    lbl_SimbMonNeto.Caption = lbl_SimbMonBruto.Caption
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Validaciones(ByRef StrMsgError As String, PAccion As String)
On Error GoTo Err
Dim CSqlC           As String
     
    If PAccion = "Grabar" Then
        
        If Len(Trim(Txt_IdSerieAtri.Text)) = 0 Then
            
            StrMsgError = "Ingresar Serie de Atribución": Txt_IdSerieAtri.SetFocus: GoTo Err
            
        End If
         
        'If Len(Trim(Txt_IdDocAtri.Text)) = 0 Then
            
        '    strMsgError = "Ingresar Número de Atribución": Txt_IdDocAtri.SetFocus: GoTo ERR
            
        'End If
        
        If Len(Trim(Txt_GlsProveedor.Text)) = 0 Then
            
            StrMsgError = "Ingresar Proveedor": Txt_IdProveedor.SetFocus: GoTo Err
            
        End If
        
        If Len(Trim(txtGls_Moneda.Text)) = 0 Then
            
            StrMsgError = "Ingresar Moneda": txtCod_Moneda.SetFocus: GoTo Err
            
        End If
        
        If Val("" & Txt_TipoCambio.Text) = 0 Then
            
            StrMsgError = "Ingresar T.C.": Txt_TipoCambio.SetFocus: GoTo Err
            
        End If
        
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Function ControlaKey(StrMsgError As String, ByVal KeyAscii As Integer, PTipo As String) As Integer
On Error GoTo Err
    
    If InStr("0123456789/-", Chr(KeyAscii)) = 0 Then
        ControlaKey = 0
    Else
        ControlaKey = KeyAscii
    End If
    If KeyAscii = 8 Then ControlaKey = KeyAscii ' borrado atras
    If KeyAscii = 13 Then ControlaKey = KeyAscii 'Enter
    If PTipo = "N" Then If KeyAscii = 46 Then ControlaKey = KeyAscii ' PARA EL PUNTO
    
    Exit Function

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Function

Private Sub txtCod_Moneda_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyCode = 113 Then
        
        cmbAyudaMoneda_Click
        
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub imprimir(StrMsgError As String)
Dim p                           As Object
Dim CSqlC                       As String
Dim rst                         As New ADODB.Recordset
Dim RstDet                      As New ADODB.Recordset
Dim NVueltas                    As Long
Dim CImpLetras                  As String
Dim CMes                        As String
On Error GoTo Err
    
    For Each p In Printers
        'If InStr(UCase(p.DeviceName), "ATRDDDDIBUCIONESSSSSS") > 0 Then
        If InStr(UCase(p.DeviceName), "ATRIBUCIONES") > 0 Then
            Set Printer = p
            Exit For
        End If
    Next p
    
    Printer.ScaleMode = 6
    'Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    
    CSqlC = "Select B.GlsPersona,B.Ruc,B.Direccion,A.FechaEmision,A.IdMoneda,A.ImpSubTotal,A.ImpIGV,A.ImpTotal " & _
            "From AtribucionesCab A " & _
            "Inner Join Personas B " & _
                "On A.IdProveedor = B.IdPersona " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSerieAtri = '" & Txt_IdSerieAtri.Text & "' And A.IdDocAtri = '" & Txt_IdDocAtri.Text & "'"
    
    rst.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
    
        ImprimeXY rst.Fields("GlsPersona") & "", "T", 100, 57, 41, 0, 0, StrMsgError
        ImprimeXY rst.Fields("Ruc") & "", "T", 11, 57, 130, 0, 0, StrMsgError
        ImprimeXY rst.Fields("Direccion") & "", "T", 100, 64, 39, 0, 0, StrMsgError
        
        ImprimeXY right(Format(Year(rst.Fields("FechaEmision")), "0000") & "", 2), "T", 2, 64, 132 + 60, 0, 0, StrMsgError
        ImprimeXY strArregloMes(Val(Format(Month(rst.Fields("FechaEmision")), "00") & "")), "T", Len(strArregloMes(Val(Format(Month(rst.Fields("FechaEmision")), "00") & ""))), 64, 132 + 13, 0, 0, StrMsgError
        ImprimeXY Format(Day(rst.Fields("FechaEmision")), "00") & "", "T", 2, 64, 132, 0, 0, StrMsgError
        
        NVueltas = 0
        
        CSqlC = "Select A.GlsProducto,A.ImpConteo,A.PorcOP,A.PorcOV,A.ImpAtribucion " & _
                "From AtribucionesDet A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSerieAtri = '" & Txt_IdSerieAtri.Text & "' And A.IdDocAtri = '" & Txt_IdDocAtri.Text & "'"
            
        RstDet.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
        
        Do While Not RstDet.EOF
            
            ImprimeXY RstDet.Fields("GlsProducto") & "", "T", 100, 77 + (4 * NVueltas), 16, 0, 0, StrMsgError
            ImprimeXY Val(RstDet.Fields("ImpConteo") & ""), "N", 18, 77 + (4 * NVueltas), 104, 2, 0, StrMsgError
            ImprimeXY Val(RstDet.Fields("PorcOP") & ""), "N", 18, 77 + (4 * NVueltas), 128, 2, 0, StrMsgError
            ImprimeXY Val(RstDet.Fields("PorcOV") & ""), "N", 18, 77 + (4 * NVueltas), 150, 2, 0, StrMsgError
            ImprimeXY Val(RstDet.Fields("ImpAtribucion") & ""), "N", 18, 77 + (4 * NVueltas), 178, 2, 0, StrMsgError
            
            RstDet.MoveNext
            
            NVueltas = NVueltas + 1
            
        Loop
        
        RstDet.Close: Set RstDet = Nothing
        
        CImpLetras = UCase(MonedaTexto(Format(rst.Fields("ImpTotal"), "0.00"), IIf(rst.Fields("IdMoneda") = "PEN", "0", "1")))
        
        ImprimeXY CImpLetras, "T", Len(CImpLetras), 131, 24, 0, 0, StrMsgError
        
        ImprimeXY Val(rst.Fields("ImpSubTotal") & ""), "N", 18, 130, 178, 2, 0, StrMsgError
        ImprimeXY Val(rst.Fields("ImpIGV") & ""), "N", 18, 139, 178, 2, 0, StrMsgError
        ImprimeXY Val(rst.Fields("ImpTotal") & ""), "N", 18, 146, 178, 2, 0, StrMsgError
        
        Printer.Print Chr$(149)
        Printer.Print ""
        Printer.Print ""
        Printer.EndDoc
        
    End If
    
    rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
'    Resume
End Sub

Private Sub ImprimeXY(varData As Variant, strTipoDato As String, intTamanoCampo As Integer, intFila As Integer, intColu As Integer, intDecimales As Integer, intFilas As Integer, ByRef StrMsgError As String)
    Dim i As Integer
    Dim strDec  As String
    Dim indFinWhile As Boolean
    Dim intFilaImp As Integer
    Dim intIndiceInicio As Integer
    
    On Error GoTo Err
    Select Case strTipoDato
        Case "T"   'texto
             
             If (intFilas = 0 Or intFilas = 1) Or Len(varData) <= intTamanoCampo Then
                
                Printer.CurrentY = intFila
                Printer.CurrentX = intColu
                
                Printer.Print left(varData, intTamanoCampo)
             Else
                indFinWhile = True
                intFilaImp = 0
                intIndiceInicio = 1
                
                Do While (indFinWhile = True)
                    If intFilaImp < intFilas Then
                        intFilaImp = intFilaImp + 1
                        
                        Printer.CurrentY = intFila
                        Printer.CurrentX = intColu
                        Printer.Print Mid(varData, intIndiceInicio, intTamanoCampo)
                        
                        intFila = intFila + 5
                        
                        intIndiceInicio = intIndiceInicio + intTamanoCampo
                    Else
                        indFinWhile = False
                    End If
                Loop
             End If
        Case "F"   'Fecha
             Printer.CurrentY = intFila
             Printer.CurrentX = intColu
             Printer.Print left(Format(varData, "dd/mm/yyyy"), intTamanoCampo)
        Case "H"   'Hora
             Printer.CurrentY = intFila
             Printer.CurrentX = intColu
             Printer.Print left(Format(varData, "hh:MM"), intTamanoCampo)
        Case "Y"   'Fecha y Hora
             Printer.CurrentY = intFila
             Printer.CurrentX = intColu
             Printer.Print left(Format(varData, "dd/mm/yyyy hh:MM"), intTamanoCampo)
        Case "N"     'numerico
            Printer.CurrentY = intFila
            Printer.CurrentX = intColu
                    
            'asig. la cantidad de decimales
            For i = 1 To intDecimales
                strDec = strDec & "0"
            Next
            
            If Val(varData) >= 0 Then
                If intDecimales > 0 Then
                    Printer.Print right((Space(intTamanoCampo) & Format(varData, "#,###,##0." & strDec)), intTamanoCampo)
                Else
                    Printer.Print right((Space(intTamanoCampo) & Format(varData, "#,###,##0" & strDec)), intTamanoCampo)
                End If
            Else
                Printer.CurrentX = intColu - 2
                Printer.Print "(" & right((Space(intTamanoCampo) & Format(Val(varData) * -1, "#,###,##0." & strDec)), intTamanoCampo - 2) & ")"
            End If
        End Select
    Exit Sub
Err:
     If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
