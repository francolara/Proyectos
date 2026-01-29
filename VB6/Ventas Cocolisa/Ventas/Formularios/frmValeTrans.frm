VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.Ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmValeTrans 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia entre Almacenes"
   ClientHeight    =   9075
   ClientLeft      =   1740
   ClientTop       =   2700
   ClientWidth     =   13500
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
   ScaleHeight     =   9075
   ScaleWidth      =   13500
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   14010
      Top             =   3660
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   14010
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   8250
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
            Picture         =   "frmValeTrans.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmValeTrans.frx":317E
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
      Height          =   8400
      Left            =   60
      TabIndex        =   9
      Top             =   600
      Width           =   13365
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3750
         Left            =   135
         OleObjectBlob   =   "frmValeTrans.frx":3518
         TabIndex        =   13
         Top             =   900
         Width           =   13110
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   3465
         Left            =   135
         OleObjectBlob   =   "frmValeTrans.frx":560C
         TabIndex        =   14
         Top             =   4755
         Width           =   13110
      End
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
         Left            =   120
         TabIndex        =   10
         Top             =   150
         Width           =   13140
         Begin VB.ComboBox cbx_Mes 
            Height          =   330
            ItemData        =   "frmValeTrans.frx":7FAC
            Left            =   2625
            List            =   "frmValeTrans.frx":7FD4
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   225
            Width           =   1755
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   9825
            TabIndex        =   11
            Top             =   270
            Visible         =   0   'False
            Width           =   1065
            _ExtentX        =   1879
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
            Container       =   "frmValeTrans.frx":803D
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Ano 
            Height          =   315
            Left            =   750
            TabIndex        =   1
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
            Container       =   "frmValeTrans.frx":8059
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2175
            TabIndex        =   16
            Top             =   270
            Width           =   300
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Año"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   345
            TabIndex        =   15
            Top             =   270
            Width           =   300
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Búsqueda"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   9030
            TabIndex        =   12
            Top             =   315
            Visible         =   0   'False
            Width           =   735
         End
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
      Height          =   8415
      Left            =   60
      TabIndex        =   8
      Top             =   600
      Width           =   13365
      Begin VB.Frame Frame3 
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
         Height          =   5190
         Left            =   120
         TabIndex        =   18
         Top             =   3075
         Width           =   13125
         Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
            Height          =   4875
            Left            =   90
            OleObjectBlob   =   "frmValeTrans.frx":8075
            TabIndex        =   7
            Top             =   195
            Width           =   12915
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   2955
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   13125
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
            Left            =   7375
            Picture         =   "frmValeTrans.frx":E5EB
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1570
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaAlmacenOrigen 
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
            Left            =   7375
            Picture         =   "frmValeTrans.frx":E975
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   770
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaAlmacenDestino 
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
            Left            =   7375
            Picture         =   "frmValeTrans.frx":ECFF
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1175
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Vale 
            Height          =   315
            Left            =   11805
            TabIndex        =   21
            Tag             =   "TidValesTrans"
            Top             =   180
            Width           =   1215
            _ExtentX        =   2143
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
            Container       =   "frmValeTrans.frx":F089
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_AlmacenOrigen 
            Height          =   315
            Left            =   1410
            TabIndex        =   2
            Tag             =   "TidAlmacenOrigen"
            Top             =   780
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
            Container       =   "frmValeTrans.frx":F0A5
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_AlmacenOrigen 
            Height          =   315
            Left            =   2340
            TabIndex        =   22
            Top             =   780
            Width           =   5010
            _ExtentX        =   8837
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
            Container       =   "frmValeTrans.frx":F0C1
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_AlmacenDestino 
            Height          =   315
            Left            =   1410
            TabIndex        =   3
            Tag             =   "TidAlmacenDestino"
            Top             =   1170
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
            Container       =   "frmValeTrans.frx":F0DD
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_AlmacenDestino 
            Height          =   315
            Left            =   2340
            TabIndex        =   23
            Top             =   1170
            Width           =   5010
            _ExtentX        =   8837
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
            Container       =   "frmValeTrans.frx":F0F9
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtObs 
            Height          =   900
            Left            =   1410
            TabIndex        =   6
            Tag             =   "TglsObs"
            Top             =   1950
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   1588
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
            Container       =   "frmValeTrans.frx":F115
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp_Emision 
            Height          =   315
            Left            =   11865
            TabIndex        =   5
            Tag             =   "FFecRegistro"
            Top             =   780
            Width           =   1200
            _ExtentX        =   2117
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
            Format          =   142082049
            CurrentDate     =   38955
         End
         Begin CATControls.CATTextBox txtNum_ValeIngreso 
            Height          =   315
            Left            =   9165
            TabIndex        =   24
            Top             =   1185
            Width           =   1170
            _ExtentX        =   2064
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
            Container       =   "frmValeTrans.frx":F131
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtNum_ValeSalida 
            Height          =   315
            Left            =   9165
            TabIndex        =   25
            Top             =   810
            Width           =   1170
            _ExtentX        =   2064
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
            Container       =   "frmValeTrans.frx":F14D
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Moneda 
            Height          =   315
            Left            =   1410
            TabIndex        =   4
            Tag             =   "TidMoneda"
            Top             =   1560
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
            Container       =   "frmValeTrans.frx":F169
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Moneda 
            Height          =   315
            Left            =   2340
            TabIndex        =   36
            Top             =   1560
            Width           =   5010
            _ExtentX        =   8837
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
            Container       =   "frmValeTrans.frx":F185
            Vacio           =   -1  'True
         End
         Begin DXDBGRIDLibCtl.dxDBGrid gDocReferencia 
            Height          =   1275
            Left            =   7920
            OleObjectBlob   =   "frmValeTrans.frx":F1A1
            TabIndex        =   38
            Top             =   1575
            Width           =   5085
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nº Transferencia"
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
            Left            =   10380
            TabIndex        =   31
            Top             =   225
            Width           =   1365
         End
         Begin VB.Label lbl_FechaEmision 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   10695
            TabIndex        =   29
            Top             =   825
            Width           =   450
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nº Vale Ingreso"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   7950
            TabIndex        =   27
            Top             =   1260
            Width           =   1140
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nº Vale Salida"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   7950
            TabIndex        =   26
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label lbl_Moneda 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   120
            TabIndex        =   37
            Top             =   1650
            Width           =   570
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
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   270
            Width           =   3480
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Almacén Origen"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   32
            Top             =   825
            Width           =   1155
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Almacén Destino"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   30
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   120
            TabIndex        =   28
            Top             =   1965
            Width           =   1110
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   2170
      ButtonWidth     =   3254
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "             Nuevo              "
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
            Caption         =   "Importar P.M. "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar Vale"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar Insumo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmValeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StrEstValeTrans                         As String
Dim dblIgvNEw                                   As Double
Dim strParamCR                                  As String
Dim RstClon                                     As New ADODB.Recordset
Dim CIdAlmacenOriAnt                            As String
Dim CIdAlmacenDesAnt                            As String

Private Sub cbx_Mes_Click()
On Error GoTo Err
Dim StrMsgError As String

    If indNuevoDoc = False Then
        listaValesTrans StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaAlmacenOrigen_Click()

    mostrarAyuda "ALMACEN", txtCod_AlmacenOrigen, txtGls_AlmacenOrigen
'    If txtCod_AlmacenOrigen.Text <> "" Then SendKeys "{TAB}"
    
End Sub

Private Sub cmbAyudaAlmacenDestino_Click()

    mostrarAyuda "ALMACENVTA", txtCod_AlmacenDestino, txtGls_AlmacenDestino

End Sub

Private Sub cmbAyudaMoneda_Click()

    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    
End Sub

Private Sub dtp_Emision_Change()
Dim strPeriodo      As String

    strPeriodo = Format(dtp_Emision.Value, "yyyymm")
    If Trim("" & traerCampo("parametros", "valparametro", "glsparametro", "PERIODO_CAMBIO_IGV", True)) > strPeriodo Then
        dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV_ANT", True), "0.00")) / 100, "0.00")
    Else
        dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV", True), "0.00")) / 100, "0.00")
    End If

End Sub

Private Sub dtp_Emision_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then SendKeys "{tab}"
    
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError     As String

    Me.top = 0
    Me.left = 0
    
    ConfGrid_Inv gDocReferencia, True, False, False, True, False
    
    txt_Ano.Text = Year(getFechaSistema)
    cbx_Mes.ListIndex = Month(getFechaSistema) - 1
    
    ConfGrid_Inv gLista, False, False, False, False, True
    ConfGrid_Inv gListaDetalle, False, False, False, False
    ConfGrid_Inv gDetalle, True, False, False, False
    
    If Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "STOCK_POR_LOTE", True) & "") = "S" Then
        gDetalle.Columns.ColumnByFieldName("idlote").Visible = True
        gDetalle.Columns.ColumnByFieldName("NumLote").Visible = True
    Else
        gDetalle.Columns.ColumnByFieldName("idlote").Visible = False
        gDetalle.Columns.ColumnByFieldName("NumLote").Visible = False
    End If
    
    If Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "VIZUALIZA_CODIGO_RAPIDO", True) & "") = "S" Then
        gDetalle.Columns.ColumnByFieldName("CodigoRapido").Visible = True
        gDetalle.Columns.ColumnByFieldName("IdProducto").Visible = False
    Else
        gDetalle.Columns.ColumnByFieldName("CodigoRapido").Visible = False
        gDetalle.Columns.ColumnByFieldName("IdProducto").Visible = True
    End If
    
    strParamCR = Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "BUSQUEDA_POR_CODIGO_RAPIDO", True) & "")
    
    listaValesTrans StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    
    If leeParametro("NO_VISUALIZAR_PESO") = "S" Then
        gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Visible = False
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo        As String
Dim strMsg           As String
Dim strCodValeIng    As String
Dim strCodValeSal    As String
Dim visCodRapido     As String
Dim strCodPro        As String
Dim dblSaldo         As Double
Dim strRUC           As String

    getEstadoCierreMes CVDate(dtp_Emision.Value), StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
'    If Val("" & traerCampo("TiposDeCambio", "TcVenta", "Fecha", Format(dtp_Emision.Value, "yyyy-mm-dd"), False)) = 0 Then
'
'        StrMsgError = "Ingrese el Tipo de Cambio del " & dtp_Emision.Value & " para proceder a Grabar.": GoTo Err
'
'    End If
    
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
    
    strRUC = traerCampo("Empresas", "RUC", "idEmpresa", glsEmpresa, False)
    
    If gDetalle.Count >= 1 Then
        If gDetalle.Count = 1 And (gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" Or gDetalle.Columns.ColumnByFieldName("Cantidad").Value <= 0) Then
            StrMsgError = "Falta Ingresar Detalle"
            GoTo Err
        End If
    End If
    
    visCodRapido = Trim("" & traerCampo("parametros", "ValParametro", "GlsParametro", "VIZUALIZA_CODIGO_RAPIDO", True))
        
    If glsValidaStock = True Then
        With gDetalle
            .Dataset.First
             Do While Not .Dataset.EOF
             
                If visCodRapido = "S" Then
                    strCodPro = traerCampo("Productos", "CodigoRapido", "idProducto", Trim("" & .Columns.ColumnByFieldName("idProducto").Value), True)
                Else
                    strCodPro = Trim("" & .Columns.ColumnByFieldName("idProducto").Value)
                End If
                
                dblSaldo = traerCantSaldo(Trim(.Columns.ColumnByFieldName("IdProducto").Value), Trim(txtCod_AlmacenOrigen.Text), Format(dtp_Emision.Value, "yyyy-mm-dd"), Trim("" & txtNum_ValeSalida.Text), StrMsgError)
                If dblSaldo < Val(.Columns.ColumnByFieldName("Cantidad").Value) Then
                    StrMsgError = StrMsgError & "En el item " & .Columns.ColumnByFieldName("item").Value & " La Cantidad Ingresada para el Producto " & " " & strCodPro & " Exede el stock Verifique" & "  " & Chr(13) & Chr(10)
                End If
                .Dataset.Next
             Loop
        End With
    End If
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_Vale.Text = "" Then  '--- Graba controlar si es nuevo o mod con una variable
        EjecutaSQLFormValesTrans Me, 0, StrMsgError, gDetalle, strCodValeIng, strCodValeSal, dtp_Emision.Value
        If StrMsgError <> "" Then GoTo Err
        
        'Tomasini 07/07/13
        If strRUC = "20513250445" Then 'Solo INMAC
           GrabaReferenciaVI StrMsgError
           If StrMsgError <> "" Then GoTo Err
           
           Enviar_Correo strCodValeIng, strCodValeSal, StrMsgError
           If StrMsgError <> "" Then GoTo Err
        End If
        strMsg = "Grabo"
    Else '--- Modifica
        EjecutaSQLFormValesTrans Me, 1, StrMsgError, gDetalle, strCodValeIng, strCodValeSal, dtp_Emision.Value
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modifico"
    End If
    
    txtNum_ValeIngreso.Text = strCodValeIng
    txtNum_ValeSalida.Text = strCodValeSal
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    fraGeneral.Enabled = False
    habilitaBotones 2
    listaValesTrans StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    '************** Luis 09/04/2019 *******************
    If InStr(1, StrMsgError, "Duplicate", vbTextCompare) > 0 Then
        txtCod_Vale.Text = ""
        txtNum_ValeIngreso.Text = ""
        txtNum_ValeSalida.Text = ""
    End If
    '***************************************************
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

Private Sub gDocReferencia_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError                         As String
    
    If gDocReferencia.Columns.ColumnByFieldName("IndImportado").Value = "1" Then
        gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").DisableEditor = True
        gDocReferencia.Columns.ColumnByFieldName("IdSerie").ReadOnly = True
        gDocReferencia.Columns.ColumnByFieldName("IdNumDoc").ReadOnly = True
    Else
        gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").DisableEditor = False
        gDocReferencia.Columns.ColumnByFieldName("IdSerie").ReadOnly = False
        gDocReferencia.Columns.ColumnByFieldName("IdNumDoc").ReadOnly = False
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gDocReferencia_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim StrDes As String
    
    Select Case Column.Index
        Case gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Index
            strCod = gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value
            StrDes = gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value
            mostrarAyudaTexto IIf(indVale = "I", "DOCUMENTOSI", "DOCUMENTOS"), strCod, StrDes
            gDocReferencia.Dataset.Edit
            gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value = strCod
            gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value = StrDes
            gDocReferencia.Dataset.Post
    End Select

End Sub

Private Sub gDocReferencia_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String
    
    If gDocReferencia.Dataset.Modified = False Then Exit Sub
    
    Select Case gDocReferencia.Columns.FocusedColumn.Index
        Case gDocReferencia.Columns.ColumnByFieldName("idSerie").Index
            gDocReferencia.Dataset.Edit
            'gDocReferencia.Columns.ColumnByFieldName("idSerie").Value = Format(gDocReferencia.Columns.ColumnByFieldName("idSerie").Value, "0000000000")
            gDocReferencia.Dataset.Post
        
        Case gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Index
            gDocReferencia.Dataset.Edit
            'gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value = Format(gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value, "0000000000000000")
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
            
            If gDocReferencia.Columns.ColumnByFieldName("IndImportado").Value <> "1" Then
            
                If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                               
                    If gDocReferencia.Count = 1 Then
                        gDocReferencia.Dataset.Edit
                        gDocReferencia.Columns.ColumnByFieldName("Item").Value = 1
                        gDocReferencia.Columns.ColumnByFieldName("idAlmacen").Value = ""
                        gDocReferencia.Columns.ColumnByFieldName("GlsAlmacen").Value = ""
                        gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value = ""
                        gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value = ""
                        gDocReferencia.Columns.ColumnByFieldName("idSerie").Value = ""
                        gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value = ""
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
    End If
    If KeyCode = 13 Then
        If gDocReferencia.Dataset.State = dsEdit Or gDocReferencia.Dataset.State = dsInsert Then
              gDocReferencia.Dataset.Post
        End If
    End If

End Sub

Private Sub gDocReferencia_OnKeyPress(Key As Integer)
Dim strCod As String
Dim StrDes As String
    
    If Key <> 9 And Key <> 13 And Key <> 27 Then
        Select Case gDocReferencia.Columns.FocusedColumn.Index
            Case gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Index
                strCod = gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value
                StrDes = gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value
                
                mostrarAyudaKeyasciiTexto Key, IIf(indVale = "I", "DOCUMENTOSI", "DOCUMENTOS"), strCod, StrDes
                Key = 0
                
                gDocReferencia.Dataset.Edit
                gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value = strCod
                gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value = StrDes
                gDocReferencia.Dataset.Post
        End Select
    End If

End Sub

Private Sub nuevo(ByRef StrMsgError As String)
On Error GoTo Err
Dim strAno As String
Dim RsD                             As New ADODB.Recordset

    strAno = txt_Ano.Text
    limpiaForm Me
    
    CIdAlmacenOriAnt = ""
    CIdAlmacenDesAnt = ""
    
    StrEstValeTrans = "GEN"
    txt_Ano.Text = strAno
    txtCod_Moneda.Text = "PEN"
    lbl_Anulado.Caption = ""
    fraGeneral.Enabled = True
    dtp_Emision_Change
    
    FormatoGrillaDetalle StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adVarChar, 16, adFldIsNullable
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
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gdetalle_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)

    If Action = daInsert Then
        gDetalle.Columns.ColumnByFieldName("item").Value = gDetalle.Count
        gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
        gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = ""
        gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
        gDetalle.Columns.ColumnByFieldName("idUM").Value = ""
        gDetalle.Columns.ColumnByFieldName("GlsUM").Value = ""
        gDetalle.Columns.ColumnByFieldName("idLote").Value = ""
        gDetalle.Columns.ColumnByFieldName("NumLote").Value = ""
        gDetalle.Columns.ColumnByFieldName("Factor").Value = 1
        gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
        gDetalle.Columns.ColumnByFieldName("Afecto").Value = 1
        gDetalle.Columns.ColumnByFieldName("VVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("PVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Value = "0"
        gDetalle.Dataset.Post
    End If

End Sub

Private Sub gdetalle_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If (gDetalle.Columns.ColumnByFieldName("idProducto").Value = "") And indInserta = False Then
            Allow = False
        Else
            gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("idProducto").ColIndex
        End If
    End If

End Sub

Private Sub gdetalle_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim rscd As New ADODB.Recordset
Dim StrMsgError As String
Dim strCod As String
Dim StrDes As String
Dim strCodUM   As String
Dim strDesUM   As String
Dim dblFactor  As Double
Dim indPedido As Boolean
Dim codigo      As String
Dim Descripcion As String
Dim codproducto As String
Dim codalmacen As String

    Select Case Column.Index
        Case gDetalle.Columns.ColumnByFieldName("idProducto").Index
            strCod = gDetalle.Columns.ColumnByFieldName("idProducto").Value
            StrDes = gDetalle.Columns.ColumnByFieldName("GlsProducto").Value
            
            If txtCod_AlmacenOrigen.Text = "" Then
                StrMsgError = "Ingrese Almacen Origen"
                txtCod_AlmacenOrigen.OnError = True
                GoTo Err
            End If
            
            strCod = ""
            StrDes = ""
            strCodUM = ""
            indPedido = False
            FrmAyudaProdOC.ExecuteReturnTextAlm txtCod_AlmacenOrigen.Text, rscd, strCod, StrDes, strCodUM, glsValidaStock, "", True, True, indPedido, False, StrMsgError
            If rscd.RecordCount <> 0 Then
                mostrarDocImportado rscd, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
        
        Case gDetalle.Columns.ColumnByFieldName("CodigoRapido").Index
        
            If txtCod_AlmacenOrigen.Text = "" Then
                StrMsgError = "Ingrese Almacen Origen"
                txtCod_AlmacenOrigen.OnError = True
                GoTo Err
            End If
            
            'FrmAyudaProductosCalcula.MostrarForm strMsgError, rscd, "", txtCod_AlmacenOrigen.Text, dtp_Emision.Value
            FrmAyudaProductosCalcula.MostrarForm StrMsgError, rscd, "", txtCod_AlmacenOrigen.Text, dtp_Emision.Value, "", "", "", "", "", ""
            If StrMsgError <> "" Then GoTo Err
            
            gDetalle.SetFocus
            If rscd.State = 1 Then
                If rscd.RecordCount <> 0 Then
                    mostrarDocImportado_AyudaMM rscd, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                End If
            End If
            
        Case gDetalle.Columns.ColumnByFieldName("idUM").Index
            strCod = gDetalle.Columns.ColumnByFieldName("idUM").Value
            StrDes = gDetalle.Columns.ColumnByFieldName("GlsUM").Value

            mostrarAyudaTexto "PRESENTACIONES", strCod, StrDes, " AND idProducto = '" & gDetalle.Columns.ColumnByFieldName("idProducto").Value & "'"
            gDetalle.SetFocus
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idUM").Value = strCod
            gDetalle.Columns.ColumnByFieldName("GlsUM").Value = StrDes
            dblFactor = traerCampo("presentaciones", "Factor", "idProducto", gDetalle.Columns.ColumnByFieldName("idProducto").Value, True, " idUM = '" & strCod & "'")
            gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
            gDetalle.Dataset.Post
            If strCod <> "" Then
                gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").ColIndex
            End If
            
        Case gDetalle.Columns.ColumnByFieldName("idlote").Index
            codalmacen = Trim("" & txtCod_AlmacenOrigen.Text)
            codproducto = Trim("" & gDetalle.Columns.ColumnByFieldName("idproducto").Value)
            FrmAyudaLotes_Vales.mostrar_from "S", Descripcion, codigo, codproducto, codalmacen
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idLote").Value = Trim(codigo)
            gDetalle.Columns.ColumnByFieldName("NumLote").Value = Trim(Descripcion)
            gDetalle.Dataset.Post
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gDetalle_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim dblVVUnit  As Double
Dim dblIGVUnit As Double
Dim dblPVUnit  As Double
Dim strCod As String
Dim StrDes As String
Dim strCodFabri As String
Dim strCodMar As String
Dim strDesMar As String
Dim intAfecto As Integer
Dim strTipoProd As String
Dim strMoneda As String
Dim strCodUM   As String
Dim strDesUM   As String
Dim dblFactor  As Double
Dim StrMsgError   As String
Dim RsP             As New ADODB.Recordset
'Actual
    If gDetalle.Dataset.Modified = False Then Exit Sub
     Select Case gDetalle.Columns.FocusedColumn.Index
      
      Case gDetalle.Columns.ColumnByFieldName("VVUnit").Index
            procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 0, Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value), dblVVUnit, dblIGVUnit, dblPVUnit
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = traerCostoUnit(Trim("" & rsdd.Fields("idProducto")), Trim("" & txtCod_AlmacenOrigen.Text), strFecIni, txtCod_Moneda.Text, StrMsgError)
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), dblVVUnit, dblIGVUnit, dblPVUnit, Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
            gDetalle.Dataset.Post
            
        Case gDetalle.Columns.ColumnByFieldName("Cantidad").Index
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Value = Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value) * Val("" & traerCampo("Productos", "IdTallaPeso", "IdProducto", gDetalle.Columns.ColumnByFieldName("IdProducto").Value, True))
            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("IGVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
            gDetalle.Dataset.Post
            
            procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 0, Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value), dblVVUnit, dblIGVUnit, dblPVUnit
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
       
        Case gDetalle.Columns.ColumnByFieldName("IdProducto").Index
            If Len(Trim(gDetalle.Columns.ColumnByFieldName("idProducto").Value)) > 0 Then
                strCod = Trim(gDetalle.Columns.ColumnByFieldName("idProducto").Value)
                StrDes = gDetalle.Columns.ColumnByFieldName("GlsProducto").Value

                csql = "SELECT idProducto,GlsProducto,idUMVenta FROM Productos " & _
                        "WHERE idempresa = '" & glsEmpresa & _
                        "' AND (idProducto = '" & strCod & "'  OR idFabricante = '" & strCod & "' OR CodigoRapido = '" & strCod & "')"
                RsP.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
                If RsP.EOF Or RsP.BOF Then
                    StrMsgError = "No se encuentra registrado el producto"
                    gDetalle.Dataset.Edit
                    gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
                    gDetalle.Dataset.Post
                    GoTo Err
                Else
                    mostrarDocImportado_Ayuda RsP, StrMsgError
                End If
                gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").ColIndex
            End If
    End Select
    If RsP.State = 1 Then RsP.Close: Set RsP = Nothing
    
    Exit Sub
    
Err:
    If RsP.State = 1 Then RsP.Close: Set RsP = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

 
Private Sub gdetalle_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gDetalle.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gDetalle.Count = 1 Then
                    gDetalle.Dataset.Edit
                    gDetalle.Columns.ColumnByFieldName("item").Value = 1
                    gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idUM").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsUM").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idLote").Value = ""
                    gDetalle.Columns.ColumnByFieldName("NumLote").Value = ""
                    gDetalle.Columns.ColumnByFieldName("Factor").Value = 1
                    gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
                    gDetalle.Columns.ColumnByFieldName("Afecto").Value = 1
                    gDetalle.Columns.ColumnByFieldName("VVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("PVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = 0
                    gDetalle.Dataset.Post
                
                Else
                    gDetalle.Dataset.Delete
                    gDetalle.Dataset.First
                    Do While Not gDetalle.Dataset.EOF
                        i = i + 1
                        gDetalle.Dataset.Edit
                        gDetalle.Columns.ColumnByFieldName("Item").Value = i
                        gDetalle.Dataset.Post
                        gDetalle.Dataset.Next
                    Loop
                    If gDetalle.Dataset.State = dsEdit Or gDetalle.Dataset.State = dsInsert Then
                        gDetalle.Dataset.Post
                    End If
                End If
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
Dim StrMsgError As String
Dim strCod As String
Dim StrDes As String
Dim strCodUM   As String
Dim strDesUM   As String
Dim dblFactor  As Double
Dim rscd As New ADODB.Recordset

    If Key <> 9 And Key <> 13 And Key <> 27 Then
        Select Case gDetalle.Columns.FocusedColumn.Index
            Case gDetalle.Columns.ColumnByFieldName("idProducto").Index
                strCod = gDetalle.Columns.ColumnByFieldName("idProducto").Value
                StrDes = gDetalle.Columns.ColumnByFieldName("GlsProducto").Value
                
                If txtCod_AlmacenOrigen.Text = "" Then
                    StrMsgError = "Ingrese Almacen Origen"
                    txtCod_AlmacenOrigen.OnError = True
                    Key = 0
                    GoTo Err
                End If
                
                If strParamCR <> "S" Then
                    
'                    mostrarAyudaKeyasciiTextoProdAlm Key, txtCod_AlmacenOrigen.Text, strCod, strDes, strCodUM, True, glsListaVentas, False, False, True, strMsgError
'                    Key = 0
'                    If strMsgError <> "" Then GoTo ERR
'
'                    gDetalle.SetFocus
'                    gDetalle.Dataset.Edit
'                    gDetalle.Columns.ColumnByFieldName("idProducto").Value = strCod
'
'                    'trae datos producto
'                    If DatosProducto(strCod, strCodUM, strDesUM, dblFactor) = False Then
'                    End If
'                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = strDes & " " & strDesMar
'                    gDetalle.Columns.ColumnByFieldName("idUM").Value = strCodUM
'                    gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
'                    gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
'                    gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
'                    gDetalle.Dataset.Post
'                    If strCod <> "" Then
'                        gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").ColIndex
'                    End If
                    
                    strCod = gDetalle.Columns.ColumnByFieldName("idProducto").Value
                    StrDes = gDetalle.Columns.ColumnByFieldName("GlsProducto").Value
                    
                    If txtCod_AlmacenOrigen.Text = "" Then
                        StrMsgError = "Ingrese Almacen Origen"
                        txtCod_AlmacenOrigen.OnError = True
                        GoTo Err
                    End If
                    
                    strCod = ""
                    StrDes = ""
                    strCodUM = ""
                    indPedido = False
                    FrmAyudaProdOC.ExecuteReturnTextAlm txtCod_AlmacenOrigen.Text, rscd, strCod, StrDes, strCodUM, glsValidaStock, "", True, True, indPedido, False, StrMsgError
                    If rscd.RecordCount <> 0 Then
                        mostrarDocImportado rscd, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
            
                End If
                
            Case gDetalle.Columns.ColumnByFieldName("idUM").Index
                strCod = gDetalle.Columns.ColumnByFieldName("idUM").Value
                StrDes = gDetalle.Columns.ColumnByFieldName("GlsUM").Value
                mostrarAyudaKeyasciiTexto Key, "PRESENTACIONES", strCod, StrDes, " AND idProducto = '" & gDetalle.Columns.ColumnByFieldName("idProducto").Value & "'"
                Key = 0
                gDetalle.SetFocus
                gDetalle.Dataset.Edit
                gDetalle.Columns.ColumnByFieldName("idUM").Value = strCod
                gDetalle.Columns.ColumnByFieldName("GlsUM").Value = StrDes
                dblFactor = traerCampo("presentaciones", "Factor", "idProducto", gDetalle.Columns.ColumnByFieldName("idProducto").Value, True, " idUM = '" & strCod & "'")
                gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
                gDetalle.Dataset.Post
                If strCod <> "" Then
                    gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").ColIndex
                End If
        End Select
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    listaDetalle

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarValeTrans gLista.Columns.ColumnByName("idValesTrans").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError                                 As String
Dim reporte                                     As String
Dim indVale                                     As String
Dim CTipoVale                                   As String
Dim RsC                                         As New ADODB.Recordset
Dim RsD                                         As New ADODB.Recordset
Dim CTipoDocImportado                           As String

    Select Case Button.Index
        Case 1 'Nuevo
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
            habilitaBotones Button.Index
            fraListado.Visible = False
            fraGeneral.Visible = True
        Case 2 'Grabar
            If txtCod_AlmacenOrigen.Text = txtCod_AlmacenDestino.Text Then
                StrMsgError = "El almacén origen no puede ser el mismo que el almacén destino"
                GoTo Err
            End If
            
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        Case 3 'Modificar
            
            StrMsgError = "La opción se encuentra des habilitada.": GoTo Err
            
            getEstadoCierreMes CVDate(dtp_Emision.Value), StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            If StrEstValeTrans = "GEN" Then
                fraGeneral.Enabled = True
                habilitaBotones Button.Index
            Else
                StrMsgError = "No se puede Modificar la Transferencia."
                GoTo Err
            End If
        Case 4 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
            habilitaBotones Button.Index
            
        Case 5 'Anular
            getEstadoCierreMes CVDate(dtp_Emision.Value), StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            
            anularDoc StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 6 'Imprimir
            If Trim(txtNum_ValeSalida.Text) <> "" And Trim(txtNum_ValeIngreso.Text <> "") Then
                
'''''                If Trim(Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "FORMATO_VALE_ALMACEN", True))) = "S" Then
'''''                    reporte = "rptImpVale" & indVale & "Trans" & "2.rpt"
'''''                Else
'''''                    reporte = "rptImpVale" & indVale & "Trans" & ".rpt"
'''''                End If
                
                CTipoVale = Trim(Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "FORMATO_VALE_ALMACEN", True)))
                indVale = "S"
                Select Case CTipoVale
                    Case "S":
                        reporte = "rptImpVale" & indVale & "Trans" & "2.rpt"
                    Case "3":
                        reporte = "rptImpVale" & indVale & "Trans" & "3.rpt"
                    Case "4":
                        reporte = "rptImpVale" & indVale & "Trans" & "4.rpt"
                    Case Else
                        reporte = "rptImpVale" & indVale & "Trans" & ".rpt"
                End Select
                
                mostrarReporte reporte, "varEmpresa|varSucursal|varNumvale|varTipovale|varNumvaleTrans", glsEmpresa & "|" & glsSucursal & "|" & txtNum_ValeSalida.Text & "|" & indVale & "|" & txtCod_Vale.Text, "vale", StrMsgError
                If StrMsgError <> "" Then GoTo Err
                                
'''''                If Trim(Trim(traerCampo("Parametros", "ValParametro", "GlsParametro", "FORMATO_VALE_ALMACEN", True))) = "S" Then
'''''                    reporte = "rptImpVale" & indVale & "2.rpt"
'''''                Else
'''''                    reporte = "rptImpVale" & indVale & ".rpt"
'''''                End If
                
                indVale = "I"
                Select Case CTipoVale
                    Case "S":
                        reporte = "rptImpVale" & indVale & "2.rpt"
                    Case "3":
                        reporte = "rptImpVale" & indVale & "3.rpt"
                    Case "4":
                        reporte = "rptImpVale" & indVale & "4.rpt"
                    Case Else
                        reporte = "rptImpVale" & indVale & ".rpt"
                End Select
                
                mostrarReporte reporte, "varEmpresa|varSucursal|varNumvale|varTipovale", glsEmpresa & "|" & traerCampo("Almacenes", "idSucursal", "idAlmacen", txtCod_AlmacenDestino.Text, True) & "|" & txtNum_ValeIngreso.Text & "|" & indVale, "vale", StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
        
        Case 7 'Lista
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
            habilitaBotones Button.Index
        
        Case 8 'Importar P.M.
            
            frmListaDocExportar.MostrarForm "TE", glsPersonaEmpresa, RsC, RsD, CTipoDocImportado, StrMsgError
            If StrMsgError <> "" Then GoTo Err
    
            If CTipoDocImportado <> "" Then
                 
                mostrarDocImportado2 RsC, RsD, CTipoDocImportado, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                 
            End If
             
            Unload frmListaDocExportar
            
        Case 9 'Importar Vale
            ImportarVales StrMsgError
            If StrMsgError <> "" Then GoTo Err
                
            
        Case 10 'Importar Receta
            ImportarReceta StrMsgError
            If StrMsgError <> "" Then GoTo Err
                
        Case 11 'Salir
            Unload Me
    End Select
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarDocImportado2(ByVal rscd As ADODB.Recordset, ByVal rsdd As ADODB.Recordset, ByVal strTipoDocImportado As String, ByRef StrMsgError As String)
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
Dim stridum As String, strGlsUM As String, strafecto As String, stridTipoProducto As String, stridMoneda As String
Dim strNumLote As String, strFecVencProd As String, stridSucursal As String
Dim nfactor As Double, ncantidad As Double, ncantidad2 As Double, NVVUnit As Double, NIGVUnit As Double, NPVUnit As Double
Dim nTotalVVBruto As Double, nTotalPVBruto As Double, nPorDcto As String, nDctoVV As Double, nDctoPV  As Double
Dim nTotalVVNeto As Double, nTotalIGVNeto As Double, nTotalPVNeto As Double
Dim nVVUnitLista As Double, nPVUnitLista As Double, nVVUnitNeto As Double, nPVUnitNeto As Double
Dim strRepetirProductosGrid As String
Dim strSerieDocImportado    As String
Dim strNumDocImportado      As String
Dim CCodigoRapido           As String

    indCargando = True
    Set rsddtemp = rsdd
    i = 0

    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adChar, 8, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsg.Fields.Append "Stock", adDouble, 14, adFldIsNullable
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
    rsg.Fields.Append "idSerieImp", adVarChar, 3, adFldIsNullable
    rsg.Fields.Append "idLote", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 50, adFldIsNullable
    rsg.Fields.Append "IdDocumentoR", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "IdSerieR", adVarChar, 3, adFldIsNullable
    rsg.Fields.Append "IdDocVentasR", adVarChar, 8, adFldIsNullable
    rsg.Open

    '--- Formato Documento de refrencia
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adVarChar, 16, adFldIsNullable
    RsD.Open , , adOpenKeyset, adLockOptimistic

    If rsdd.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("CodigoRapido") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("idUM") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("Factor") = 1
        rsg.Fields("Afecto") = 1
        rsg.Fields("Stock") = 0
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
        rsg.Fields("NumLote") = ""
        rsg.Fields("IdDocumentoR") = ""
        rsg.Fields("IdSerieR") = ""
        rsg.Fields("IdDocVentasR") = ""
    
    Else
    
        rscd.MoveFirst
        If Not rscd.EOF Then
            
            If Len(Trim("" & rscd.Fields("IdAlmacen"))) > 0 Then txtCod_Almacen.Text = "" & rscd.Fields("IdAlmacen")
            'txtCod_Cliente.Text = "" & rscd.Fields("IdPersona")
            If Len(Trim("" & rscd.Fields("IdMoneda"))) > 0 Then txtCod_Moneda.Text = "" & rscd.Fields("IdMoneda")
            txtObs.Text = "" & rscd.Fields("ObsDocVentas")
            
            If strTipoDocImportado = "PM" Then
        
                'TxtIdArea.Text = "" & rscd.Fields("IdUPP")
                'CIdDocumento = "" & rscd.Fields("TipoDocReferencia")
                'TxtIdSerie.Text = "" & rscd.Fields("SerieDocReferencia")
                'TxtIdDocVentas.Text = "" & rscd.Fields("NumDocReferencia")
                
            End If
            
            'txtCod_CentroCosto.Text = "" & rscd.Fields("IdCentroCosto")
            'CIdSucursalPres = "" & rscd.Fields("IdSucursal")
            
        End If
        
        rsdd.MoveFirst
        rsdd.Sort = "idProducto"
        
        Do While Not rsdd.EOF
            strIdDocventas = "" & rsdd.Fields("idDocVentas")
            strIdSerie = "" & rsdd.Fields("idSerie")
            strcodprod_aux = rsdd.Fields("idProducto")
            CCodigoRapido = rsdd.Fields("CodigoRapido")
            stridCodFabricante = "" & rsdd.Fields("idCodFabricante")
            strGlsproducto = "" & rsdd.Fields("GlsProducto")
            stridmarca = "" & rsdd.Fields("idMarca")
            strGlsMarca = "" & rsdd.Fields("GlsMarca")
            stridum = "" & rsdd.Fields("idUM")
            strGlsUM = "" & rsdd.Fields("GlsUM")
            nfactor = "" & rsdd.Fields("Factor")
            strafecto = "" & rsdd.Fields("Afecto")
            stridTipoProducto = "" & rsdd.Fields("idTipoProducto")
            stridMoneda = "" & rsdd.Fields("idMoneda")
            strNumLote = "" & rsdd.Fields("NumLote")
            strFecVencProd = "" & rsdd.Fields("FecVencProd")
            stridSucursal = "" & rsdd.Fields("idSucursal")
            ncantidad2 = "" & rsdd.Fields("Cantidad2")
            ncantidad = "" & rsdd.Fields("Cantidad")
            NVVUnit = "" & rsdd.Fields("VVUnit")
            NIGVUnit = "" & rsdd.Fields("IGVUnit")
            NPVUnit = "" & rsdd.Fields("PVUnit")
            nTotalVVBruto = "" & rsdd.Fields("TotalVVBruto")
            nTotalPVBruto = "" & rsdd.Fields("TotalPVBruto")
            nPorDcto = "" & rsdd.Fields("PorDcto")
            nDctoVV = "" & rsdd.Fields("DctoVV")
            nDctoPV = "" & rsdd.Fields("DctoPV")
            'nTotalVVNeto = "" & rsdd.Fields("TotalVVNeto")
            'nTotalIGVNeto = "" & rsdd.Fields("TotalIGVNeto")
            'nTotalPVNeto = "" & rsdd.Fields("TotalPVNeto")
            nVVUnitLista = "" & rsdd.Fields("VVUnitLista")
            nPVUnitLista = "" & rsdd.Fields("PVUnitLista")
            nVVUnitNeto = "" & rsdd.Fields("VVUnitNeto")
            nPVUnitNeto = "" & rsdd.Fields("PVUnitNeto")
            nTotalVVNeto = rsdd.Fields("Cantidad") * NVVUnit
            If strafecto = 1 Then
                nTotalIGVNeto = nTotalVVNeto * dblIgvNEw
            Else
                nTotalIGVNeto = 0
            End If
            nTotalPVNeto = nTotalVVNeto + nTotalIGVNeto
            
            rsg.AddNew
            i = i + 1
            rsg.Fields("Item") = i
            rsg.Fields("idProducto") = strcodprod_aux
            rsg.Fields("CodigoRapido") = CCodigoRapido
            rsg.Fields("GlsProducto") = strGlsproducto
            rsg.Fields("idUM") = stridum
            rsg.Fields("GlsUM") = strGlsUM
            rsg.Fields("Factor") = nfactor
            rsg.Fields("Afecto") = strafecto
            rsg.Fields("Stock") = 0
            rsg.Fields("Cantidad") = ncantidad
            rsg.Fields("Cantidad2") = ncantidad2
                        
            rsg.Fields("VVUnit") = traerCostoUnit(strcodprod_aux, Trim("" & txtCod_AlmacenOrigen.Text), Format(dtp_Emision.Value), txtCod_Moneda.Text, StrMsgError) 'nvvunit
            If strafecto = "1" Then
                rsg.Fields("IGVUnit") = Val("" & rsg.Fields("VVUnit")) * dblIgvNEw 'nIGVUnit
            Else
                rsg.Fields("IGVUnit") = 0
            End If
            
            rsg.Fields("PVUnit") = Val("" & rsg.Fields("VVUnit")) + Val("" & rsg.Fields("IGVUnit")) 'nPVUnit
            
            rsg.Fields("TotalVVNeto") = Val("" & rsg.Fields("VVUnit")) * ncantidad 'nTotalVVNeto
            If strafecto = "1" Then
                rsg.Fields("TotalIGVNeto") = Val("" & rsg.Fields("TotalVVNeto")) * dblIgvNEw 'nTotalIGVNeto
            Else
                rsg.Fields("TotalIGVNeto") = 0
            End If
            
            rsg.Fields("TotalPVNeto") = Val("" & rsg.Fields("TotalVVNeto")) + Val("" & rsg.Fields("TotalIGVNeto")) 'nTotalPVNeto
            
            rsg.Fields("NumLote") = strNumLote
            rsg.Fields("FecVencProd") = strFecVencProd
            rsg.Fields("idDocumentoImp") = "94"
            rsg.Fields("idDocVentasImp") = strIdDocventas
            rsg.Fields("idSerieImp") = strIdSerie
            rsg.Fields("IdSucursalOrigen") = stridSucursal
            rsg.Fields("idLote") = ""
            rsg.Fields("NumLote") = ""
            rsg.Fields("IdDocumentoR") = "" & rsdd.Fields("IdDocumentoR")
            rsg.Fields("IdSerieR") = "" & rsdd.Fields("IdSerieR")
            rsg.Fields("IdDocVentasR") = "" & rsdd.Fields("IdDocVentasR")
            
            rsdd.MoveNext
        Loop
    End If

    rsddtemp.MoveFirst
    Do While Not rsddtemp.EOF
        If RsD.RecordCount > 0 Then RsD.MoveFirst
        indExisteDocRef = False
        Do While Not RsD.EOF
            If Trim(RsD.Fields("idDocumento")) = strTipoDocImportado And RsD.Fields("idSerie") = "" & rsddtemp.Fields("idSerie") And RsD.Fields("idNumDOc") = "" & rsddtemp.Fields("idDocVentas") Then
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
            strSerieDocImportado = strIdSerie
            strNumDocImportado = strIdDocventas
        End If
        rsddtemp.MoveNext
    Loop

    rsdd.MoveFirst
    txtCod_Moneda.Text = traerCampo("docventaspEDIDOS", "idMoneda", "iddocumento", strTipoDocImportado, True, " idSerie = '" & rsdd.Fields("idSerie") & "' and idDocVentas = '" & rsdd.Fields("idDocVentas") & "' ")

    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
 
    '--- Documentos de referencia
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

'    calcularTotales
    indCargando = False
    
'    txtCod_CentroCosto_Change
    
    Me.Refresh

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
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
            Toolbar1.Buttons(8).Visible = Not indHabilitar 'Importar P.M.
            Toolbar1.Buttons(9).Visible = Not indHabilitar 'Importar Vale
            Toolbar1.Buttons(10).Visible = Not indHabilitar 'Importar Vale
            
        Case 4, 7 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
            Toolbar1.Buttons(8).Visible = False
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(10).Visible = False
            
    End Select
    
    ocultarColumnasEstado
    If Toolbar1.Buttons(8).Visible Then
        If leeParametro("IMPORTAR_PM_TRANSFERENCIAS") = "1" Then
            Toolbar1.Buttons(8).Visible = True
        Else
            Toolbar1.Buttons(8).Visible = False
        End If
    End If
    
    If Toolbar1.Buttons(9).Visible Then
        If Len(Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "CONCEPTO_FILTRO_AYUDATRANS", True))) > 0 Then
            Toolbar1.Buttons(9).Visible = True
        Else
            Toolbar1.Buttons(9).Visible = False
        End If
    End If
    
End Sub

Private Sub txt_Ano_Change()
On Error GoTo Err
Dim StrMsgError As String

    If indNuevoDoc = False Then
        listaValesTrans StrMsgError
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

    listaValesTrans StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus
    
End Sub

Private Sub listaValesTrans(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsdatos    As New ADODB.Recordset

    csql = "SELECT v.idValesTrans,v.FecRegistro,v.idAlmacenOrigen,o.GlsAlmacen as GlsAlmacenOrigen,v.idAlmacenDestino,d.GlsAlmacen as GlsAlmacenDestino " & _
            "FROM valestrans v,almacenes o,almacenes d " & _
            "WHERE v.idAlmacenOrigen = o.idAlmacen " & _
            "AND v.idAlmacenDestino = d.idAlmacen " & _
            "AND v.idEmpresa = '" & glsEmpresa & "' AND v.idSucursal = '" & glsSucursal & "' " & _
            "AND o.idEmpresa = '" & glsEmpresa & "' " & _
            "AND d.idEmpresa = '" & glsEmpresa & "' " & _
            "AND  year(FecRegistro) = " & Val(txt_Ano.Text) & " AND Month(FecRegistro) = " & cbx_Mes.ListIndex + 1
    csql = csql & " ORDER BY idValesTrans"
    
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
'        .KeyField = "idValesTrans"
'    End With
    
    listaDetalle
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarValeTrans(strnum As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst                         As New ADODB.Recordset
Dim rsg                         As New ADODB.Recordset
Dim RsD                         As New ADODB.Recordset
Dim dblVVUnit                   As Double
Dim dblIGVUnit                  As Double
Dim dblPVUnit                   As Double

    indCargando = True
    csql = "SELECT idValesTrans, idEmpresa, idSucursal, FecRegistro, IdAlmacenOrigen, IdAlmacenDestino, idValeIngreso, idValeSalida, glsObs, estValeTrans, idMoneda " & _
           "FROM valestrans d " & _
           "WHERE d.idValesTrans = '" & strnum & "' AND d.idEmpresa = '" & glsEmpresa & "' AND d.idSucursal = '" & glsSucursal & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        StrEstValeTrans = Trim("" & rst.Fields("estValeTrans"))
        If StrEstValeTrans = "ANU" Then
            lbl_Anulado.Caption = "ANULADO"
            fraGeneral.Enabled = False
        Else
            lbl_Anulado.Caption = ""
        End If
        
        CIdAlmacenOriAnt = Trim("" & rst.Fields("IdAlmacenOrigen"))
        CIdAlmacenDesAnt = Trim("" & rst.Fields("IdAlmacenDestino"))
    
    End If
    
    dtp_Emision_Change
    
    txtNum_ValeIngreso.Tag = "TidValeIngreso"
    txtNum_ValeSalida.Tag = "TidValeSalida"
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    txtNum_ValeIngreso.Tag = ""
    txtNum_ValeSalida.Tag = ""

    csql = "Select vt.idValesTrans, vt.idEmpresa, vt.idSucursal, p.idProducto,  p.glsProducto, Item, idUM, Factor, " & _
          "Cantidad ,0 VVUnit,0 IGVUnit,0 PVUnit,0 TotalVVNeto,0 TotalIGVNeto,0 TotalPVNeto,P.AfectoIgv Afecto,P.CodigoRapido,vd.IdTallaPeso " & _
          "From ValesTrans vt " & _
          "Inner Join ValesTransDet vd   On vt.idValesTrans = vd.idValesTrans  And vt.idEmpresa = vd.idEmpresa  And vt.idSucursal = vd.idSucursal " & _
          "Inner Join Productos P On vd.IdEmpresa = P.IdEmpresa And vd.IdProducto = P.IdProducto " & _
          "Where vt.idValesTrans = '" & strnum & "' " & _
          "And vt.idEmpresa = '" & glsEmpresa & "' And vt.idSucursal = '" & glsSucursal & "' " & _
          "Order By Item "
           
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idLote", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 50, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 40, adFldIsNullable
    rsg.Fields.Append "IdTallaPeso", adVarChar, 30, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("CodigoRapido") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("idUM") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("idLote") = ""
        rsg.Fields("NumLote") = ""
        rsg.Fields("Factor") = 1
        rsg.Fields("Cantidad") = 0
        rsg.Fields("VVUnit") = 0
        rsg.Fields("IGVUnit") = 0
        rsg.Fields("PVUnit") = 0
        rsg.Fields("TotalVVNeto") = 0
        rsg.Fields("TotalIGVNeto") = 0
        rsg.Fields("TotalPVNeto") = 0
        rsg.Fields("Afecto") = 1
        rsg.Fields("IdTallaPeso") = "0"
        
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idProducto") = "" & rst.Fields("idProducto")
            rsg.Fields("CodigoRapido") = "" & rst.Fields("CodigoRapido")
            rsg.Fields("GlsProducto") = "" & rst.Fields("GlsProducto")
            rsg.Fields("idUM") = "" & rst.Fields("idUM")
            rsg.Fields("idLote") = traerCampo("Sucursales", "idLote", "idSucursal", glsSucursal, True)
            rsg.Fields("NumLote") = traerCampo("Sucursales", "idLote", "idSucursal", glsSucursal, True)
            rsg.Fields("GlsUM") = traerCampo("unidadmedida", "abreUM", "idUM", ("" & rst.Fields("idUM")), False)
            rsg.Fields("Factor") = "" & rst.Fields("Factor")
            rsg.Fields("Cantidad") = "" & rst.Fields("Cantidad")
            rsg.Fields("VVUnit") = "" & rst.Fields("VVUnit")
            rsg.Fields("IGVUnit") = "" & rst.Fields("IGVUnit")
            rsg.Fields("PVUnit") = "" & rst.Fields("PVUnit")
            rsg.Fields("TotalVVNeto") = "" & rst.Fields("TotalVVNeto")
            rsg.Fields("TotalIGVNeto") = "" & rst.Fields("TotalIGVNeto")
            rsg.Fields("TotalPVNeto") = "" & rst.Fields("TotalPVNeto")
            rsg.Fields("Afecto") = 1
            rsg.Fields("IdTallaPeso") = "" & rst.Fields("IdTallaPeso")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    gDetalle.Dataset.First
    Do While Not gDetalle.Dataset.EOF
        
        gDetalle.Dataset.Edit
        
        gDetalle.Columns.ColumnByFieldName("VVUnit").Value = traerCostoUnit(Trim("" & gDetalle.Columns.ColumnByFieldName("IdProducto").Value), Trim("" & txtCod_AlmacenOrigen.Text), Format(dtp_Emision.Value, "yyyy-mm-dd"), txtCod_Moneda.Text, StrMsgError)
        If StrMsgError <> "" Then GoTo Err
        
        calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("IGVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
        
        gDetalle.Dataset.Post
        
        procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 0, Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value), dblVVUnit, dblIGVUnit, dblPVUnit
        gDetalle.Dataset.Edit
        gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
        gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
        gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            
        gDetalle.Dataset.Next
    
    Loop
    
    csql = "SELECT r.item, r.tipoDocReferencia idDocumento,d.GlsDocumento, r.numDocReferencia idNumDoc, r.serieDocReferencia idSerie " & _
            "FROM docreferencia r , documentos d " & _
            "WHERE r.tipoDocReferencia = d.idDocumento AND tipoDocOrigen = 'TE' AND numDocOrigen = '" & strnum & "' AND serieDocOrigen = '000' AND r.idEmpresa = '" & glsEmpresa & "' AND r.idSucursal = '" & glsSucursal & "' ORDER BY ITEM"
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adVarChar, 16, adFldIsNullable
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
    Me.Refresh
    
    Exit Sub

Err:
    txtNum_ValeIngreso.Tag = ""
    txtNum_ValeSalida.Tag = ""
    If StrMsgError = "" Then StrMsgError = Err.Description
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

Private Function DatosProducto(strCodProd As String, ByRef strCodUM As String, ByRef strDesUM As String, ByRef dblFactor As Double) As Boolean
Dim rst As New ADODB.Recordset

    csql = "SELECT p.AfectoIGV,p.idMoneda,p.idUMCompra,u.abreUM " & _
            "FROM productos p,unidadmedida u " & _
            "WHERE p.idUMCompra = u.idUM " & _
            "AND p.idEmpresa = '" & glsEmpresa & "' " & _
            "AND p.idProducto = '" & strCodProd & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        DatosProducto = True
        strCodUM = "" & rst.Fields("idUMCompra")
        strDesUM = "" & rst.Fields("abreUM")
        dblFactor = 1
     Else
        DatosProducto = False
        strCodUM = ""
        strDesUM = ""
        dblFactor = 0
    End If
    rst.Close: Set rst = Nothing

End Function

Private Sub listaDetalle()
Dim rsdatos                     As New ADODB.Recordset

    csql = "SELECT  V.Item,V.idProducto,p.GlsProducto,M.GlsMarca,V.idUM,U.GlsUM,V.Cantidad,V.IdTallaPeso " & _
            "FROM valestransdet V, productos P,marcas M,unidadmedida U " & _
            "WHERE V.idProducto = P.idProducto " & _
            "AND P.idMarca    = M.idMarca " & _
            "AND V.idEmpresa = '" & glsEmpresa & "' AND V.idSucursal = '" & glsSucursal & "' " & _
            "AND p.idEmpresa = '" & glsEmpresa & "' " & _
            "AND m.idEmpresa = '" & glsEmpresa & "' " & _
            "AND  V.idUM       = U.idUM AND idValesTrans = '" & gLista.Columns.ColumnByFieldName("idValesTrans").Value & "'" & _
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

Private Sub ocultarColumnasEstado()

    Select Case StrEstValeTrans
        Case "ANU"
            Toolbar1.Buttons(2).Visible = False 'GRABAR
            Toolbar1.Buttons(3).Visible = False 'MODIFICAR
            Toolbar1.Buttons(4).Visible = False 'CANCELAR
            Toolbar1.Buttons(5).Visible = False 'ANULAR
            Toolbar1.Buttons(6).Visible = False 'IMPRIMIR
            Toolbar1.Buttons(8).Visible = False 'IMPORTAR
    End Select

End Sub

Private Sub txtCod_AlmacenOrigen_Change()
    
    txtGls_AlmacenOrigen.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", txtCod_AlmacenOrigen.Text, True)

End Sub

Private Sub txtCod_AlmacenDestino_Change()

    If Len(Trim(txtCod_AlmacenOrigen.Text)) > 0 And Len(Trim(txtCod_AlmacenDestino.Text)) Then
        If txtCod_AlmacenOrigen.Text = txtCod_AlmacenDestino.Text Then
            MsgBox "El almacen origen no puede ser el mismo que el almacen destino"
        Else
            txtGls_AlmacenDestino.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", txtCod_AlmacenDestino.Text, True)
        End If
    End If
    
End Sub

Private Sub txtCod_AlmacenOrigen_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "ALMACEN", txtCod_AlmacenOrigen, txtGls_AlmacenOrigen
        KeyAscii = 0
        If txtCod_AlmacenOrigen.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_AlmacenDestino_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "ALMACEN", txtCod_AlmacenDestino, txtGls_AlmacenDestino
        KeyAscii = 0
    End If

End Sub

Private Sub FormatoGrillaDetalle(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg As New ADODB.Recordset

    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adChar, 12, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idLote", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsg.Fields.Append "IdTallaPeso", adVarChar, 30, adFldIsNullable
    rsg.Open
    
    rsg.AddNew
    rsg.Fields("Item") = 1
    rsg.Fields("idProducto") = ""
    rsg.Fields("CodigoRapido") = ""
    rsg.Fields("GlsProducto") = ""
    rsg.Fields("idUM") = ""
    rsg.Fields("GlsUM") = ""
    rsg.Fields("idLote") = ""
    rsg.Fields("NumLote") = ""
    rsg.Fields("Factor") = 1
    rsg.Fields("Cantidad") = 0
    rsg.Fields("VVUnit") = 0
    rsg.Fields("TotalVVNeto") = 0
    rsg.Fields("IGVUnit") = 0
    rsg.Fields("PVUnit") = 0
    rsg.Fields("TotalIGVNeto") = 0
    rsg.Fields("TotalPVNeto") = 0
    rsg.Fields("Afecto") = 1
    rsg.Fields("IdTallaPeso") = "0"
    
    Set gDetalle.DataSource = Nothing
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("idProducto").ColIndex

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub anularDoc(ByRef StrMsgError As String)
On Error GoTo Err
Dim SucurSalOri      As String
Dim SucurSalDes      As String
Dim indEvaluacion                   As Integer
Dim strCodUsuarioAutorizacion       As String
Dim motanula                        As String
Dim obs                             As String
Dim CIdSucursalOriAnt               As String
Dim CIdSucursalDesAnt               As String

    SucurSalOri = Trim("" & traerCampo("almacenes", "idsucursal", "idalmacen", txtCod_AlmacenOrigen.Text, True))
    SucurSalDes = Trim("" & traerCampo("almacenes", "idsucursal", "idalmacen", txtCod_AlmacenDestino.Text, True))

    If MsgBox("Seguro de Anular el Vale", vbQuestion + vbYesNo, App.Title) = vbYes Then
        
        indEvaluacion = 0
    
        frmAprobacion.MostrarForm "05", indEvaluacion, strCodUsuarioAutorizacion, StrMsgError
        If indEvaluacion = 0 Then Exit Sub
        If StrMsgError <> "" Then GoTo Err
                    
        frmMotivosAnula.Motivos_Anulacion motanula, obs
        If motanula = 0 Then Exit Sub
        If StrMsgError <> "" Then GoTo Err
        
        CIdSucursalOriAnt = traerCampo("Almacenes A", "A.IdSucursal", "A.IdAlmacen", CIdAlmacenOriAnt, True)
        CIdSucursalDesAnt = traerCampo("Almacenes A", "A.IdSucursal", "A.IdAlmacen", CIdAlmacenDesAnt, True)
                
        Actualiza_Stock_Nuevo StrMsgError, "E", CIdSucursalDesAnt, "I", txtNum_ValeIngreso.Text, CIdAlmacenDesAnt
        If StrMsgError <> "" Then GoTo Err
        
        Actualiza_Stock_Nuevo StrMsgError, "E", CIdSucursalOriAnt, "S", txtNum_ValeSalida.Text, CIdAlmacenDesAnt
        If StrMsgError <> "" Then GoTo Err
        
        'Anulamos los Vales
        csql = "UPDATE valescab SET estValeCab = 'ANU' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & SucurSalDes & "' AND idValesCab = '" & txtNum_ValeIngreso.Text & "' AND tipoVale = 'I' "
        Cn.Execute csql
        
        csql = "UPDATE valescab SET estValeCab = 'ANU' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & SucurSalOri & "' AND idValesCab = '" & txtNum_ValeSalida.Text & "' AND tipoVale = 'S' "
        Cn.Execute csql
        
        'Anulamos la Operacion
        csql = "UPDATE valestrans SET estValeTrans = 'ANU' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesTrans = '" & txtCod_Vale.Text & "'"
        Cn.Execute csql
        
        StrEstValeTrans = "ANU"
        lbl_Anulado.Caption = "ANULADO"
        fraGeneral.Enabled = False
        
        'ACTUALIZAMOS STOCK EN LINEA
        'actualizaStock_Trans txtNum_ValeSalida.Text, 1, "S", StrMsgError, False, SucurSalOri
        'If StrMsgError <> "" Then GoTo Err
        
        'actualizaStock_Lote_Trans txtNum_ValeSalida.Text, 1, "S", StrMsgError, False, SucurSalOri
        'If StrMsgError <> "" Then GoTo Err
                  
        'ACTUALIZAMOS STOCK EN LINEA ValeIngreso
        'actualizaStock_Trans txtNum_ValeIngreso.Text, 1, "I", StrMsgError, False, SucurSalDes
        'If StrMsgError <> "" Then GoTo Err
        
        'actualizaStock_Lote_Trans txtNum_ValeIngreso.Text, 1, "I", StrMsgError, False, SucurSalDes
        'If StrMsgError <> "" Then GoTo Err
        
        listaValesTrans StrMsgError
        If StrMsgError <> "" Then GoTo Err
        habilitaBotones 5
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_AlmacenOrigen_LostFocus()
    
    If txtCod_AlmacenOrigen.Text = "" Then Exit Sub
    If txtCod_AlmacenOrigen.Text = txtCod_AlmacenDestino.Text Then
        MsgBox "El almacen origen no puede ser el mismo que el almacen destino"
    End If

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

Private Sub mostrarDocImportado(ByVal rsdd As ADODB.Recordset, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg As New ADODB.Recordset
Dim RsD As New ADODB.Recordset
Dim strSerieDocVentas As String
Dim dblTC  As Double
Dim strCodFabri As String
Dim strCodMar As String
Dim strDesMar As String
Dim intAfecto As Integer
Dim strTipoProd As String
Dim strMoneda As String
Dim strCodUM   As String
Dim strDesUM   As String
Dim dblVVUnit  As Double
Dim dblIGVUnit  As Double
Dim dblPVUnit  As Double
Dim dblFactor  As Double
Dim intFila As Integer
Dim i As Integer
Dim indExisteDocRef As Boolean
Dim primero         As Boolean
Dim strInserta      As Boolean
Dim strFecIni       As String
    
    strFecIni = Format(dtp_Emision.Value, "yyyy-mm-dd")
    primero = True
    rsdd.MoveFirst
    Do While Not rsdd.EOF
        strInserta = True
        If strInserta = True Then
            If primero = True Then
                primero = False
            Else
                gDetalle.Dataset.Insert
            End If
            gDetalle.SetFocus
            gDetalle.Dataset.RecNo = intFila
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" & rsdd.Fields("idProducto")
            gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = "" & rsdd.Fields("GlsProducto")
            strCodUM = traerCampo("productos", "idUMCompra", "idProducto", "" & rsdd.Fields("idProducto"), True)
            If strDesUM = "" And strCodUM <> "" Then strDesUM = traerCampo("unidadMedida", "abreUM", "idUM", strCodUM, False)
            If Trim("" & rsdd.Fields("idProducto")) = "" Then Exit Sub
            
            '--- Trae datos producto
            If DatosProducto("" & rsdd.Fields("idProducto"), strCodUM, strDesUM, dblFactor) = False Then
            End If
            
            gDetalle.Columns.ColumnByFieldName("idUM").Value = strCodUM
            gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
            gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
            gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
            gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Value = "0"
            gDetalle.Columns.ColumnByFieldName("Afecto").Value = 1
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = 0
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = 0
            gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
            gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = 0
            gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = 0

            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = traerCostoUnit(Trim("" & rsdd.Fields("idProducto")), Trim("" & txtCod_AlmacenOrigen.Text), strFecIni, txtCod_Moneda.Text, StrMsgError)
            If StrMsgError <> "" Then GoTo Err

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
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Function traerCostoUnit(ByVal codproducto As String, ByVal codalmacen As String, ByVal PFecha As String, ByVal CodMoneda As String, ByRef StrMsgError As String) As Double
On Error GoTo Err
Dim CosUni  As ADODB.Recordset

   csql = "Select (SUM(CASE WHEN valescab.tipoVale = 'I' THEN valesdet.Cantidad ELSE valesdet.Cantidad * -1 END * " & _
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
    csql = csql & " AND CAST(valescab.fechaEmision AS DATE) <= CAST('" & PFecha & "' AS DATE) And valesdet.idProducto = '" & codproducto & "' "
    csql = csql & "AND valescab.idAlmacen = '" & codalmacen & "' "
    csql = csql & "AND valescab.estValeCab <> 'ANU' "

    Set CosUni = New ADODB.Recordset
    CosUni.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not CosUni.EOF Then
       traerCostoUnit = IIf(IsNull(CosUni.Fields("COSTO_UNITARIO")), 0, CosUni.Fields("COSTO_UNITARIO"))
    End If
    If CosUni.State = 1 Then CosUni.Close: Set CosUni = Nothing
    
    Exit Function
    
Err:
    If CosUni.State = 1 Then CosUni.Close: Set CosUni = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Function

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

Private Sub procesaMoneda(strMonProd As String, strMonDoc As String, intTipoValor As Integer, dblValor As Double, intAfecto As Integer, ByRef dblVVUnit As Double, ByRef dblIGVUnit As Double, ByRef dblPVUnit As Double)
Dim dblIGV As Double

    dblIGV = dblIgvNEw
    dblTC = Val(Format(traerCampo("tiposdecambio", "tcVenta", "day(fecha)", Day(dtp_Emision.Value), False, " month(fecha)= " & Month(dtp_Emision.Value) & " and year(fecha)= " & Year(dtp_Emision.Value) & " "), "0.000"))
    
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
    
    If intTipoValor = 0 Then 'valor venta
        dblVVUnit = dblValor
        dblIGVUnit = dblValor * dblIGV
        dblPVUnit = dblVVUnit + dblIGVUnit
    Else 'precio venta
        dblVVUnit = dblValor / (dblIGV + 1)
        dblIGVUnit = dblValor - dblVVUnit
        dblPVUnit = dblValor
    End If

End Sub

Private Sub MostrarValeImportado(PRsC As ADODB.Recordset, PRsD As ADODB.Recordset, StrMsgError As String)
On Error GoTo Err
Dim strCodFabri As String
Dim strCodMar As String
Dim strDesMar As String
Dim intAfecto As Integer
Dim strTipoProd As String
Dim strMoneda As String
Dim strCodUM   As String
Dim strDesUM   As String
Dim dblVVUnit  As Double
Dim dblIGVUnit  As Double
Dim dblPVUnit  As Double
Dim dblFactor  As Double
Dim intFila As Integer
Dim i As Integer
Dim indExisteDocRef As Boolean
Dim primero         As Boolean
Dim strInserta      As Boolean
Dim strFecIni       As String
Dim RsD             As New ADODB.Recordset

    '--- Formato Documento de refrencia
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adVarChar, 16, adFldIsNullable
    RsD.Open , , adOpenKeyset, adLockOptimistic
    
    If PRsC.RecordCount > 0 Then
        PRsC.MoveFirst
        If Not PRsC.EOF Then
            txtCod_AlmacenOrigen.Text = "" & PRsC.Fields("IdAlmacen")
            
            RsD.AddNew
            RsD.Fields("Item") = "" & RsD.RecordCount
            RsD.Fields("idDocumento") = "88"
            RsD.Fields("GlsDocumento") = traerCampo("documentos", "GlsDocumento", "idDocumento", "88", False)
            RsD.Fields("idSerie") = "999"
            RsD.Fields("idNumDOc") = "" & "" & PRsC.Fields("IdValesCab")
            
        End If
    End If
    
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
    
    strFecIni = Format(dtp_Emision.Value, "yyyy-mm-dd")
    primero = True
    PRsD.MoveFirst
    Do While Not PRsD.EOF
        If primero = True Then
            primero = False
        Else
            gDetalle.Dataset.Insert
        End If
        gDetalle.SetFocus
        gDetalle.Dataset.RecNo = intFila
        gDetalle.Dataset.Edit
        gDetalle.Columns.ColumnByFieldName("IdProducto").Value = "" & PRsD.Fields("IdProducto")
        gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = "" & PRsD.Fields("CodigoRapido")
        gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = "" & PRsD.Fields("GlsProducto")
        strCodUM = traerCampo("Productos", "IdUMCompra", "IdProducto", "" & PRsD.Fields("IdProducto"), True)
        If strDesUM = "" And strCodUM <> "" Then strDesUM = traerCampo("UnidadMedida", "AbreUM", "IdUM", strCodUM, False)
        If Trim("" & PRsD.Fields("IdProducto")) = "" Then Exit Sub
        
        'trae datos producto
        If DatosProducto("" & PRsD.Fields("IdProducto"), strCodUM, strDesUM, dblFactor) = False Then
        End If
        
        gDetalle.Columns.ColumnByFieldName("IdUM").Value = strCodUM
        gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
        gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
        gDetalle.Columns.ColumnByFieldName("Cantidad").Value = Val("" & PRsD.Fields("Cantidad"))
        gDetalle.Columns.ColumnByFieldName("IdTallaPeso").Value = Val("" & PRsD.Fields("Cantidad")) * Val("" & traerCampo("Productos", "IdTallaPeso", "IdProducto", "" & PRsD.Fields("IdProducto"), True))
        
        gDetalle.Columns.ColumnByFieldName("Afecto").Value = 1
        gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("PVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = 0

        gDetalle.Columns.ColumnByFieldName("VVUnit").Value = traerCostoUnit(Trim("" & PRsD.Fields("IdProducto")), Trim("" & txtCod_AlmacenOrigen.Text), strFecIni, txtCod_Moneda.Text, StrMsgError)
        If StrMsgError <> "" Then GoTo Err

        gDetalle.Dataset.Post
        gDetalle.Dataset.RecNo = intFila
        gDetalle.Dataset.Edit
        gDetalle.Dataset.Post
        
        If "" & PRsD.Fields("IdProducto") <> "" Then
            gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").ColIndex
        End If
        
        
        
        PRsD.MoveNext
    Loop
    'Tomasini 07/07/13
     'InsertaReferenciaVI PRsC, StrMsgError
     'If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub ImportarVales(StrMsgError As String)
On Error GoTo Err
Dim RsC     As New ADODB.Recordset
Dim RsD     As New ADODB.Recordset
    
    FrmListaValesExportar.MostrarForm txtCod_AlmacenOrigen.Text, RsC, RsD, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If RsC.State = 1 Then
        If RsC.RecordCount <> 0 Then
            MostrarValeImportado RsC, RsD, StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarDocImportado_AyudaMM(ByVal rsdd As ADODB.Recordset, ByRef StrMsgError As String)
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
Dim CWhereProductos                     As String
Dim CSqlC                               As String
Dim RsC                                 As New ADODB.Recordset

    rsdd.MoveFirst
    Do While Not rsdd.EOF
        CWhereProductos = CWhereProductos & "'" & "" & rsdd.Fields("IdProducto") & "',"
        rsdd.MoveNext
    Loop
    rsdd.Close ': Set rsdd = Nothing
    
    If Len(Trim(CWhereProductos)) = 0 Then
        Exit Sub
    Else
        CWhereProductos = left(CWhereProductos, Len(CWhereProductos) - 1)
    End If
    
    CSqlC = "Select A.IdProducto,A.CodigoRapido,A.GlsProducto,A.IdUMCompra " & _
            "From Productos A " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProducto In(" & CWhereProductos & ")"
    
    rsdd.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    strFecIni = Format(dtp_Emision.Value, "yyyy-mm-dd")
    primero = True
    rsdd.MoveFirst
    Do While Not rsdd.EOF
        strInserta = True
        If strInserta = True Then
            If primero = True Then
                primero = False
            Else
                gDetalle.Dataset.Insert
            End If
        
            gDetalle.SetFocus
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" & rsdd.Fields("idProducto")
            gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = "" & rsdd.Fields("CodigoRapido")
            gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = "" & rsdd.Fields("GlsProducto")
            strCodUM = "" & rsdd.Fields("idUMCompra")
            If strDesUM = "" And strCodUM <> "" Then strDesUM = traerCampo("unidadMedida", "abreUM", "idUM", strCodUM, False)
            If Trim("" & rsdd.Fields("idProducto")) = "" Then Exit Sub
            
            If DatosProducto("" & rsdd.Fields("idProducto"), strCodUM, strDesUM, dblFactor) = False Then
            End If
            
            gDetalle.Columns.ColumnByFieldName("idUM").Value = strCodUM
            gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
            gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
            gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 1
            
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = traerCostoUnit(Trim("" & rsdd.Fields("idProducto")), Trim("" & txtCod_AlmacenOrigen.Text), strFecIni, txtCod_Moneda.Text, StrMsgError)
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
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
End Sub

Private Function traerCantSaldo(ByVal codproducto As String, ByVal codalmacen As String, ByVal strFecha As String, ByRef PValeCab As String, ByRef StrMsgError As String) As Double
On Error GoTo Err
Dim rsSaldo As New ADODB.Recordset
Dim strSQL As String

    traerCantSaldo = 0

    strSQL = " SELECT Format(ifnull(XZ.sc_stock,0) + ifnull(s.Stock,0),2) as Stock " & _
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
    "AND vc.idAlmacen = '" & codalmacen & "' AND DATE_FORMAT(vc.fechaemision, '%Y%m%d')  = DATE_FORMAT(sysdate(), '%Y%m%d') AND (p.idProducto = '" & codproducto & "') " & _
    "Group bY P.idEmpresa,P.idProducto,vc.idAlmacen) S " & _
    "On P.idEmpresa = S.idEmpresa And P.idProducto = S.idProducto " & _
    "Left Join (SELECT sc_periodo,sc_codalm,sc_codart,sc_stock,idempresa FROM tbsaldo_costo_kardex z " & _
    "where sc_codalm = '" & codalmacen & "' and sc_periodo = DATE_FORMAT(sysdate(), '%Y%m') and sc_stock <> 0) XZ " & _
    "On P.idEmpresa  = xz.idempresa And P.idProducto = xz.sc_codart " & _
    "Where p.idEmpresa = '" & glsEmpresa & "' AND (p.idProducto = '" & codproducto & "') AND estProducto = 'A' "
            
    rsSaldo.Open strSQL, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsSaldo.EOF Then
        If Not IsNull(rsSaldo.Fields("STOCK")) Then
            traerCantSaldo = rsSaldo.Fields("STOCK")
        End If
    End If
    
    If rsSaldo.State = 1 Then rsSaldo.Close: Set rsSaldo = Nothing
    
    Exit Function

Err:
    If rsSaldo.State = 1 Then rsSaldo.Close: Set rsSaldo = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Function

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
            strCodUM = traerCampo("productos", "idUMCompra", "idProducto", "" & rsdd.Fields("idProducto"), True)
            If strDesUM = "" And strCodUM <> "" Then strDesUM = traerCampo("unidadMedida", "abreUM", "idUM", strCodUM, False)
            If Trim("" & rsdd.Fields("idProducto")) = "" Then Exit Sub
            
            If DatosProducto_Ayuda("" & rsdd.Fields("idProducto"), strCodUM, strDesUM, dblFactor) = False Then
            End If
            
            gDetalle.Columns.ColumnByFieldName("idUM").Value = strCodUM
            gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
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
     
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

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


Private Sub InsertaReferenciaVI(rst As ADODB.Recordset, StrMsgError As String)
On Error GoTo Err
    
    If RstClon.State = 1 Then RstClon.Close: Set RstClon = Nothing
  RstClon.Fields.Append "idValeIngreso", adChar, 8, adFldIsNullable
  RstClon.Fields.Append "idOCompra", adVarChar, 200, adFldIsNullable
  RstClon.Fields.Append "idRCompra", adVarChar, 200, adFldIsNullable
  RstClon.Open , , adOpenKeyset, adLockOptimistic
  
  If rst.RecordCount > 0 Then
  rst.MoveFirst
    If Not rst.EOF Then
      Do While Not rst.EOF
       RstClon.AddNew
       RstClon.Fields("idValeIngreso") = rst.Fields("IdValesCab")
       RstClon.Fields("idOCompra") = Trim("" & traerCampo("docreferencia", "group_concat(numdocreferencia) as idOCompra", "numdocorigen", rst.Fields("IdValesCab"), True, "tipodocorigen = '88' And tipodocreferencia = '94' And idSucursal ='" & glsSucursal & "' "))
       RstClon.Fields("idRCompra") = Trim("" & traerCampo("docreferencia RQ Inner Join(  Select tipoDocReferencia As TipDoc, numDocReferencia As NumOC, serieDocReferencia  AS SerieOC, idEmpresa, idSucursal  From DocReferencia  where idempresa   = '" & glsEmpresa & "'  and numdocorigen = '" & rst.Fields("IdValesCab") & "'  And tipodocorigen = '88'  And tipodocreferencia = '94') OC  On RQ.tipoDocOrigen = OC.TipDoc And RQ.numDocOrigen = OC.NumOC And  RQ.serieDocOrigen = OC.SerieOC  And RQ.idEmpresa = OC.idEmpresa And RQ.idSucursal  = OC.idSucursal", " Group_Concat(numdocreferencia) As idRCompra", "RQ.idSucursal", glsSucursal, False, "RQ.idEmpresa = '" & glsEmpresa & "'"))
       rst.MoveNext
      Loop
    End If
  End If
  
  
  Exit Sub
Err:
If RstClon.State = 1 Then RstClon.Close: Set RstClon = Nothing
If StrMsgError = "" Then StrMsgError = Err.Description
Exit Sub
Resume
End Sub

Private Sub GrabaReferenciaVI(StrMsgError As String)
On Error GoTo Err
 
  If RstClon.RecordCount > 0 Then
  RstClon.MoveFirst
  Do While Not RstClon.EOF
  
    csql = "Insert Into ValesTransRef(idValesTrans, idEmpresa, idSucursal, idValeIngreso, idOCompra, idRCompra)" & _
           "Values('" & txtCod_Vale.Text & "','" & glsEmpresa & "','" & glsSucursal & "','" & RstClon.Fields("idValeIngreso") & "','" & RstClon.Fields("idOCompra") & "','" & RstClon.Fields("idRCompra") & "')"
    Cn.Execute (csql)
    
    RstClon.MoveNext
   Loop
  End If
  
  If RstClon.State = 1 Then RstClon.Close: Set RstClon = Nothing

  Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
 
Private Sub Enviar_Correo(strCodValeI As String, strCodValeS As String, ByRef StrMsgError As String)
Dim StrNomDocProv                As String
Dim CorreoRespRC                 As String
Dim CorreoRespAlDes              As String
On Error GoTo Err

StrNomDocProv = ""


        CorreoRespRC = Trim("" & traerCampo("Personas", "mail", "idPersona", traerCampo("Docventas", "idPervendedor", "idDocventas", traerCampo("ValesTransRef", "idRCompra", "idValesTrans", txtCod_Vale.Text, True), True, "idSucursal = '" & glsSucursal & "' And idSerie = '999' And idDocumento  ='87'"), False))
        CorreoRespAlDes = Trim("" & traerCampo("Personas", "mail", "idPersona", traerCampo("Almacenes", "idResponsable", "idAlmacen", txtCod_AlmacenDestino.Text, True), False))

        StrNomDocProv = "Transferencia N°" & txtCod_Vale.Text


        If Len(Trim("" & CorreoRespRC)) > 0 Then
        
 
            ExportarReporte "rptImpValeSTrans2.rpt", "parEmpresa|parSucursal|parNumvale|parTipovale|parNumvaleTrans", glsEmpresa & "|" & glsSucursal & "|" & strCodValeS & "|" & indVale & "|" & txtCod_Vale.Text, "Vale", StrNomDocProv, StrMsgError
            If StrMsgError <> "" Then GoTo Err
             
            ExportarReporte "rptImpValeI2.rpt", "parEmpresa|parSucursal|parNumvale|parTipovale", glsEmpresa & "|" & traerCampo("Almacenes", "idSucursal", "idAlmacen", txtCod_AlmacenDestino.Text, True) & "|" & strCodValeI & "|" & indVale, "Vale", StrNomDocProv, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        
            With MAPISession1
                .NewSession = False
                .SignOn
            End With

            With MAPIMessages1
                .SessionID = MAPISession1.SessionID
                .Compose ' CREAMOS EL MENSAJE
                .MsgSubject = "Transferencia de Almacén" ' ASUNTO DEL MENSAJE
                .MsgNoteText = "Se Adjunta Transferencia de Almacén N°" & txtCod_Vale.Text ' MENSAJE

                'XXXXXXCORREOSXXXXXX
                .RecipIndex = 0
                .RecipDisplayName = CorreoRespRC  ' Receptor
                .RecipType = mapToList

                If Len(Trim("" & CorreoResp)) > 0 Then
                    .RecipIndex = 1
                    .RecipDisplayName = CorreoRespAlDes ' Copia
                    .RecipType = mapCcList
                End If
                'XXXXXXXXXXXXXXXXXXX


                'XXXXXXADJUNTOSXXXXX
                .AttachmentIndex = 0
                .AttachmentPathName = App.Path & "\Temporales\" & StrNomDocProv & ".pdf" ' ARCHIVO ADJUNTO
                .AttachmentPosition = 0
    
                'XXXXXXXXXXXXXXXXXXX

                .Send False ' ENVIA CORREO
            End With

            MAPISession1.SignOff ' CIERRA SESIÓN ABIERTA
        End If


   Exit Sub
Err:
MAPISession1.SignOff
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub EjecutaSQLFormValesTrans(F As Form, tipoOperacion As Integer, ByRef StrMsgError As String, G As dxDBGrid, ByRef strVarCodValeIng As String, ByRef strVarCodValeSal As String, strFecha As String)
On Error GoTo Err
Dim C                   As Object
Dim rst                 As New ADODB.Recordset
Dim csql                As String
Dim strCampo            As String
Dim strTipoDato         As String
Dim strCampos           As String
Dim strValores          As String
Dim strValCod           As String
Dim strCod              As String
Dim strCodValeSalida    As String
Dim strCodValeIngreso   As String
Dim indTrans            As Boolean
Dim swProducto          As String
Dim csucursal_recep     As String, cperiodo_recep As String, calmacen_destino As String
Dim cIdSucursalOri      As String
Dim cIdSucursalDes      As String
Dim cIdPeriodoInvOri    As String
Dim cIdPeriodoInvDes    As String
Dim CArray()            As String
Dim cAlmacen_Origen     As String
Dim LoteOri             As String
Dim LoteDes             As String
Dim RsTempValeS         As New ADODB.Recordset
Dim RsTempValeI         As New ADODB.Recordset
Dim CadMysqlTemp        As String
Dim strValorTotal       As Double
Dim strIgvTotal         As Double
Dim strPrecioTotal      As Double
Dim CIdSucursalOriAnt   As String
Dim CIdSucursalDesAnt   As String
Dim strCodAlmacen       As String
Dim strSigno            As String
Dim ntc                 As Double
Dim RsC                 As New ADODB.Recordset
Dim rsl                 As New ADODB.Recordset
Dim CSqlC               As String

    strCodValeIngreso = ""
    strCodValeSalida = ""
        
    indTrans = False
    csql = ""
    calmacen_destino = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                
                If strCampo = "idAlmacenDestino" Then
                    calmacen_destino = Trim(C.Value)
                End If
                
                If strCampo = "idAlmacenOrigen" Then
                    cAlmacen_Origen = Trim(C.Value)
                End If
                
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        If UCase(strCampo) = UCase("idValesTrans") Then
                            strValCod = generaCorrelativoAnoMesFecha("valestrans", "idValesTrans", strFecha)
                            C.Text = strValCod
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
                        If UCase(strCampo) <> UCase("idValesTrans") Then
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
    
    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            csql = "INSERT INTO valestrans(" & strCampos & ",idEmpresa,idSucursal,idUsuarioRegistro, FechaRegistro, HoraRegistro) VALUES(" & strValores & ",'" & glsEmpresa & "','" & glsSucursal & "','" & glsUser & "','" & Format(Date, "yyyy-mm-dd") + Space(1) + Format(Time, "h:mm:ss") & "', '" & Format(Time, "h:mm:ss") & "')"
            Cn.Execute csql
        
        Case 1
            csql = "Select IdValeIngreso,IdValeSalida,IdAlmacenOrigen,IdAlmacenDestino " & _
                    "From ValesTrans " & _
                    "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And IdValesTrans = '" & strValCod & "'"
            rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            
            If Not rst.EOF Then
                CIdSucursalOriAnt = traerCampo("Almacenes A", "A.IdSucursal", "A.IdAlmacen", "" & rst.Fields("IdAlmacenOrigen"), True)
                CIdSucursalDesAnt = traerCampo("Almacenes A", "A.IdSucursal", "A.IdAlmacen", "" & rst.Fields("IdAlmacenDestino"), True)
    
                strCodValeIngreso = "" & rst.Fields("IdValeIngreso")
                strCodValeSalida = "" & rst.Fields("IdValeSalida")
                
                '--- Eliminamos el vale de ingreso
                CadMysqlTemp = "Select * From valesdet " & _
                               "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & CIdSucursalDesAnt & "' " & _
                               "And IdValesCab = '" & strCodValeIngreso & "' AND tipoVale = 'I'"
                               
                If RsTempValeI.State = 1 Then RsTempValeI.Close
                RsTempValeI.Open CadMysqlTemp, Cn, adOpenStatic, adLockReadOnly
                
                Graba_Logico_Vales "1", StrMsgError, CIdSucursalDesAnt, strCodValeIngreso, "I"
                If StrMsgError <> "" Then GoTo Err
                
                Actualiza_Stock_Nuevo StrMsgError, "E", CIdSucursalDesAnt, "I", strCodValeIngreso, CIdAlmacenDesAnt
                If StrMsgError <> "" Then GoTo Err
        
                Cn.Execute "Delete From ValesDet " & _
                           "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & CIdSucursalDesAnt & "' " & _
                           "And IdValesCab = '" & strCodValeIngreso & "' AND tipoVale = 'I'"
                
                Cn.Execute "Delete From ValesCab " & _
                           "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & CIdSucursalDesAnt & "' " & _
                           "And IdValesCab = '" & strCodValeIngreso & "' AND tipoVale = 'I'"
                
                '--- Eliminamos el vale de salida
                
                CadMysqlTemp = "Select * From valesdet " & _
                               "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & CIdSucursalOriAnt & "' " & _
                               "And IdValesCab = '" & strCodValeSalida & "' AND tipoVale = 'S'"
                               
                If RsTempValeS.State = 1 Then RsTempValeS.Close
                RsTempValeS.Open CadMysqlTemp, Cn, adOpenStatic, adLockReadOnly
                
                Graba_Logico_Vales "1", StrMsgError, CIdSucursalOriAnt, strCodValeSalida, "S"
                If StrMsgError <> "" Then GoTo Err
                
                Actualiza_Stock_Nuevo StrMsgError, "E", CIdSucursalOriAnt, "S", strCodValeSalida, CIdAlmacenOriAnt
                If StrMsgError <> "" Then GoTo Err
                
                Cn.Execute "Delete From ValesDet " & _
                           "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & CIdSucursalOriAnt & "' " & _
                           "And IdValesCab = '" & strCodValeSalida & "' AND tipoVale = 'S'"
                
                Cn.Execute "Delete From ValesCab " & _
                           "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & CIdSucursalOriAnt & "' " & _
                           "And IdValesCab = '" & strCodValeSalida & "' AND tipoVale = 'S'"
            End If
            csql = "UPDATE valestrans SET " & strCampos & " WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesTrans = '" & strValCod & "'"
            
            Cn.Execute csql
    End Select
    
    'Grabando Grilla detalle
    If TypeName(G) <> "Nothing" Then
        Cn.Execute "DELETE FROM valestransdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesTrans = '" & strValCod & "'"
        
        G.Dataset.First
        Do While Not G.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To G.Columns.Count - 1
                If UCase(left(G.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(G.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(G.Columns(i).ObjectName, 3)
                    
                    strCampos = strCampos & strCampo & ","
                    
                    Select Case strTipoDato
                        Case "N"
                            strValores = strValores & Val(G.Columns(i).Value) & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(G.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(G.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO valestransdet(" & strCampos & ",idValesTrans,idEmpresa,idSucursal) VALUES(" & strValores & ",'" & strValCod & "','" & glsEmpresa & "','" & glsSucursal & "')"
            Cn.Execute csql
            
            G.Dataset.Next
        Loop
    End If
    
    'Grabando Grilla DocRef
    If TypeName(gDocReferencia) <> "Nothing" Then
        Cn.Execute "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = 'TE' AND numDocOrigen = '" & strValCod & "' AND serieDocOrigen = '000'"
        If gDocReferencia.Count > 0 Then
            gDocReferencia.Dataset.First
            Do While Not gDocReferencia.Dataset.EOF
                If Trim(gDocReferencia.Columns.ColumnByFieldName("IdDocumento").Value) <> "" And Trim(gDocReferencia.Columns.ColumnByFieldName("idSerie").Value) <> "" And Trim(gDocReferencia.Columns.ColumnByFieldName("idNumDoc").Value) <> "" Then
                    strCampos = ""
                    strValores = ""
                    For i = 0 To gDocReferencia.Columns.Count - 1
                        If UCase(left(gDocReferencia.Columns(i).ObjectName, 1)) = "W" Then
                            strTipoDato = Mid(gDocReferencia.Columns(i).ObjectName, 2, 1)
                            strCampo = Mid(gDocReferencia.Columns(i).ObjectName, 3)
                            
                            strCampos = strCampos & strCampo & ","
                            
                            If strCampo = "tipoDocReferencia" Then
                                strValores = strValores & "'PM',"
                                
                            Else
                            Select Case strTipoDato
                                Case "N"
                                    strValores = strValores & gDocReferencia.Columns(i).Value & ","
                                Case "T"
                                    strValores = strValores & "'" & Trim(gDocReferencia.Columns(i).Value) & "',"
                                Case "F"
                                    strValores = strValores & "'" & Format(gDocReferencia.Columns(i).Value, "yyyy-mm-dd") & "',"
                            End Select
                            End If
                        End If
                    Next
                    
                    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
                    If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
                    
                    csql = "INSERT INTO docreferencia(" & strCampos & ",tipoDocOrigen,numDocOrigen,serieDocOrigen,idEmpresa,idSucursal) VALUES(" & strValores & ",'TE','" & strValCod & "','000','" & glsEmpresa & "','" & glsSucursal & "')"
                    Cn.Execute csql
                End If
                gDocReferencia.Dataset.Next
            Loop
        End If
    End If
    
    ReDim CArray(2)
    traerCampos "Almacenes A,PeriodosInv P", "A.IdSucursal,P.IdPeriodoInv", "A.IdAlmacen", "" & cAlmacen_Origen, 2, CArray(), False, "A.IdEmpresa = '" & glsEmpresa & "' And A.IdEmpresa = P.IdEmpresa And A.IdSucursal = P.IdSucursal And P.EstPeriodoInv = 'ACT'"
    
    cIdSucursalOri = "" & CArray(0)
    cIdPeriodoInvOri = "" & CArray(1)
    
    
    If Len(Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", cIdSucursalOri, True, " year(FecInicio) = " & Year(F.dtp_Emision.Value) & " "))) = 0 Then
        cIdPeriodoInvOri = cIdPeriodoInvOri
    Else
        cIdPeriodoInvOri = Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", cIdSucursalOri, True, " year(FecInicio) = " & Year(F.dtp_Emision.Value) & " "))
    End If
    
    ReDim CArray(2)
    traerCampos "Almacenes A,PeriodosInv P", "A.IdSucursal,P.IdPeriodoInv", "A.IdAlmacen", "" & calmacen_destino, 2, CArray(), False, "A.IdEmpresa = '" & glsEmpresa & "' And A.IdEmpresa = P.IdEmpresa And A.IdSucursal = P.IdSucursal And P.EstPeriodoInv = 'ACT'"
    
    cIdSucursalDes = "" & CArray(0)
    cIdPeriodoInvDes = "" & CArray(1)
    
    If Len(Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", cIdSucursalDes, True, " year(FecInicio) = " & Year(F.dtp_Emision.Value) & " "))) = 0 Then
        cIdPeriodoInvDes = cIdPeriodoInvDes
    Else
        cIdPeriodoInvDes = Trim("" & traerCampo("periodosinv", "idPeriodoInv", "idSucursal", cIdSucursalDes, True, " year(FecInicio) = " & Year(F.dtp_Emision.Value) & " "))
    End If
    
    '--- GENERAMOS VALE SALIDA
    '--- GENERAMOS CODIGO DEL VALE
    If strCodValeSalida = "" Then
        strCodValeSalida = generaCorrelativoAnoMes_ValeFecha("ValesCab", "idValesCab", "S", strFecha, True)
    End If
    
    '--- GENERAMOS CABECERA
    
    strValorTotal = 0
    strIgvTotal = 0
    strPrecioTotal = 0
    
    '--- Asignamos Valores a los totales de las cabeceras
    strValorTotal = Format(G.Columns.ColumnByFieldName("TotalVVNeto").SummaryFooterValue, "0.00")
    strIgvTotal = Format(G.Columns.ColumnByFieldName("TotalIGVNeto").SummaryFooterValue, "0.00")
    strPrecioTotal = Format(strValorTotal + strIgvTotal, "0.00")
    
    csql = "Insert Into ValesCab (idValesCab,tipoVale,fechaEmision, valorTotal, igvTotal,precioTotal, idProvCliente, " & _
            "idConcepto, idAlmacen, idMoneda,GlsDocReferencia,TipoCambio,idEmpresa,idSucursal,idPeriodoInv,obsValesCab,FechaRegistro,IdUsuarioRegistro) " & _
            "Select '" & strCodValeSalida & "','S',d.FecRegistro,'" & strValorTotal & "','" & strIgvTotal & "','" & strPrecioTotal & "','" & glsSystem & "','25',d.idAlmacenOrigen,d.idMoneda," & _
            "('Trans. - ' + D.IdValesTrans),CAST(ISNULL(B.TcVenta,0) AS NUMERIC(12,2)),idEmpresa,'" & cIdSucursalOri & "','" & cIdPeriodoInvOri & "',d.GlsObs,GETDATE(),'" & glsUser & "' " & _
            "From valestrans d " & _
            "Left Join TiposDeCambio B " & _
                "On Cast(d.FecRegistro As Date) = Cast(B.Fecha As Date) " & _
            "Where d.idValesTrans = '" & strValCod & "'" & _
             " AND d.idEmpresa = '" & glsEmpresa & "'" & _
             " AND d.idSucursal = '" & glsSucursal & "'"
             
    Cn.Execute csql
    
    LoteOri = Trim("" & traerCampo("Sucursales", "idLote", "idSucursal", cIdSucursalOri, True))
    
    '--- GENERAMOS DETALLE DEL VALE
    With G
        .Dataset.First
        If Not .Dataset.EOF Then
            Do While Not .Dataset.EOF
                
                csql = "INSERT INTO valesdet (tipoVale,idLote,idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit," & _
                "TotalVVNeto,TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal,IdTallaPeso,NumLote) " & _
                "VALUES('S','" & .Columns.ColumnByFieldName("IdLote").Value & "','" & strCodValeSalida & "' ,'" & .Columns.ColumnByFieldName("item").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("idProducto").Value & "','" & .Columns.ColumnByFieldName("GlsProducto").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("idUM").Value & "','" & .Columns.ColumnByFieldName("Factor").Value & "',1," & _
                "'" & .Columns.ColumnByFieldName("Cantidad").Value & "','" & .Columns.ColumnByFieldName("VVUnit").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("IGVUnit").Value & "','" & .Columns.ColumnByFieldName("PVUnit").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("TotalVVNeto").Value & "','" & .Columns.ColumnByFieldName("TotalIGVNeto").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("TotalPVNeto").Value & "','','" & glsEmpresa & "','" & cIdSucursalOri & "'," & _
                "'" & .Columns.ColumnByFieldName("IdTallaPeso").Value & "','" & .Columns.ColumnByFieldName("NumLote").Value & "')"
               Cn.Execute csql
               .Dataset.Next
            Loop
        End If
    End With
    
    Actualiza_Stock_Nuevo StrMsgError, "I", cIdSucursalOri, "S", strCodValeSalida, txtCod_AlmacenOrigen.Text
    If StrMsgError <> "" Then GoTo Err
    
    '--- GENERAMOS VALE DE INGRESO
    '--- GENERAMOS CODIGO DEL VALE
    If strCodValeIngreso = "" Then
        strCodValeIngreso = generaCorrelativoAnoMes_ValeFecha("ValesCab", "idValesCab", "I", strFecha, True)
    End If
    
    '--- GENERAMOS CABECERA
    csql = "INSERT INTO ValesCab (idValesCab,tipoVale,fechaEmision, valorTotal, igvTotal," & _
           "precioTotal, idProvCliente,idConcepto, idAlmacen, idMoneda,GlsDocReferencia,TipoCambio,idEmpresa,idSucursal,idPeriodoInv,obsValesCab,FechaRegistro,IdUsuarioRegistro)" & _
           "Select '" & strCodValeIngreso & "','I',d.FecRegistro,'" & strValorTotal & "','" & strIgvTotal & "','" & strPrecioTotal & "','" & glsSystem & "','26',d.idAlmacenDestino,d.idMoneda," & _
           "('Trans. - ' + D.IdValesTrans),CAST(ISNULL(B.TcVenta,0) AS NUMERIC(12,2)),idEmpresa,'" & cIdSucursalDes & "','" & cIdPeriodoInvDes & "',d.GlsObs,GETDATE(),'" & glsUser & "' " & _
           "FROM valestrans d " & _
           "Left Join TiposDeCambio B " & _
                "On Cast(d.FecRegistro As Date) = Cast(B.Fecha As Date) " & _
           "WHERE d.idValesTrans = '" & strValCod & "'" & _
            " AND d.idEmpresa = '" & glsEmpresa & "'" & _
            " AND d.idSucursal = '" & glsSucursal & "'"
            
    Cn.Execute csql
    
    '--- GENERAMOS DETALLE DEL VALE
    LoteDes = Trim("" & traerCampo("Sucursales", "idLote", "idSucursal", cIdSucursalDes, True))
    With G
        .Dataset.First
        If Not .Dataset.EOF Then
            Do While Not .Dataset.EOF
                
                csql = "INSERT INTO valesdet (tipoVale,idLote,idValesCab,item,idProducto,GlsProducto,idUM,Factor,Afecto,Cantidad,VVUnit,IGVUnit,PVUnit," & _
                "TotalVVNeto,TotalIGVNeto,TotalPVNeto,idMoneda,idEmpresa,idSucursal,IdTallaPeso,NumLote) " & _
                "VALUES('I','" & .Columns.ColumnByFieldName("IdLote").Value & "','" & strCodValeIngreso & "','" & .Columns.ColumnByFieldName("item").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("idProducto").Value & "','" & .Columns.ColumnByFieldName("GlsProducto").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("idUM").Value & "','" & .Columns.ColumnByFieldName("Factor").Value & "',1," & _
                "'" & .Columns.ColumnByFieldName("Cantidad").Value & "','" & .Columns.ColumnByFieldName("VVUnit").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("IGVUnit").Value & "','" & .Columns.ColumnByFieldName("PVUnit").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("TotalVVNeto").Value & "','" & .Columns.ColumnByFieldName("TotalIGVNeto").Value & "'," & _
                "'" & .Columns.ColumnByFieldName("TotalPVNeto").Value & "','','" & glsEmpresa & "','" & cIdSucursalDes & "'," & _
                "'" & .Columns.ColumnByFieldName("IdTallaPeso").Value & "','" & .Columns.ColumnByFieldName("NumLote").Value & "')"
                Cn.Execute csql
                .Dataset.Next
            Loop
        End If
    End With
    
    Actualiza_Stock_Nuevo StrMsgError, "I", cIdSucursalDes, "I", strCodValeIngreso, txtCod_AlmacenDestino.Text
    If StrMsgError <> "" Then GoTo Err
    
    '--- GENERAMOS RELACION PRODUCTOS ALMACEN
    csql = "SELECT v.idAlmacenDestino, d.idProducto, 0, v.idEmpresa, v.idSucursal, d.idUM, 0, 0 " & _
           "FROM valestrans v, valestransdet d " & _
           "Where v.idEmpresa = D.idEmpresa " & _
           "and v.idValesTrans = d.idValesTrans " & _
           "and v.idSucursal = d.idSucursal " & _
           "and v.idEmpresa = '" & glsEmpresa & "' " & _
           "and v.idSucursal = '" & glsSucursal & "' " & _
           "and v.idValesTrans = '" & strValCod & "' "
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
           
    If rst.RecordCount <> 0 Then
        Do While Not rst.EOF
            swProducto = ""
            swProducto = traerCampo("productosalmacen", "idProducto", "idProducto", rst.Fields("idProducto"), True, "idAlmacen = '" & rst.Fields("idAlmacenDestino") & "' and idSucursal = '" & cIdSucursalDes & "' ")
            If swProducto = "" Then
                csql = "insert into productosalmacen(idAlmacen, idProducto, item, idEmpresa, " & _
                                                    "idSucursal, idUMCompra, CantidadStock, CostoUnit) " & _
                       "values('" & rst.Fields("idAlmacenDestino") & "', '" & rst.Fields("idProducto") & "', 0, '" & rst.Fields("idEmpresa") & "', " & _
                       "'" & cIdSucursalDes & "', '" & rst.Fields("idUM") & "', 0, 0)"
                Cn.Execute csql
            End If
            rst.MoveNext
        Loop
                        
        rst.MoveFirst
        Do While Not rst.EOF
            swProducto = ""
            swProducto = traerCampo("productosalmacenporlote", "idProducto", "idProducto", rst.Fields("idProducto"), True, "idAlmacen = '" & rst.Fields("idAlmacenDestino") & "' and idSucursal = '" & cIdSucursalDes & "' and idLote = '" & LoteDes & "' ")
            If swProducto = "" Then
                csql = "insert into productosalmacenporlote(idLote,idAlmacen, idProducto, item, idEmpresa, " & _
                                                    "idSucursal, idUMCompra, CantidadStock, CostoUnit) " & _
                       "values('" & LoteDes & "','" & rst.Fields("idAlmacenDestino") & "', '" & rst.Fields("idProducto") & "', 0, '" & rst.Fields("idEmpresa") & "', " & _
                       "'" & cIdSucursalDes & "', '" & rst.Fields("idUM") & "', 0, 0)"
                Cn.Execute csql
            End If
            rst.MoveNext
        Loop
    End If
    
'    If tipoOperacion = "1" Then
'         If Not RsTempValeS.EOF Then
'            RsTempValeS.MoveFirst
'            Do While Not RsTempValeS.EOF
'                strSigno = "+"
'                strCodAlmacen = traerCampo("valescab", "idalmacen", "idvalescab", Trim("" & RsTempValeS.Fields("idvalescab")), True, " tipoVale = 'S' ")
'
'                csql = "UPDATE productosalmacen " & _
'                        "SET CantidadStock = CantidadStock " & strSigno & " " & RsTempValeS.Fields("Cantidad") & " " & _
'                        "WHERE idEmpresa = '" & RsTempValeS.Fields("idEmpresa") & "' " & _
'                        "AND idSucursal = '" & RsTempValeS.Fields("idSucursal") & "' " & _
'                        "AND idAlmacen = '" & strCodAlmacen & "' " & _
'                        "AND idProducto = '" & RsTempValeS.Fields("idProducto") & "' " & _
'                        "AND idUMCompra = '" & RsTempValeS.Fields("idUM") & "' "
'                Cn.Execute csql
'
'                csql = "UPDATE productosalmacenporlote " & _
'                        "SET CantidadStock = CantidadStock " & strSigno & " " & RsTempValeS.Fields("Cantidad") & " " & _
'                        "WHERE idEmpresa = '" & RsTempValeS.Fields("idEmpresa") & "' " & _
'                          "AND idSucursal = '" & RsTempValeS.Fields("idSucursal") & "' " & _
'                          "AND idAlmacen = '" & strCodAlmacen & "' " & _
'                          "AND idProducto = '" & RsTempValeS.Fields("idProducto") & "' " & _
'                          "AND idUMCompra = '" & RsTempValeS.Fields("idUM") & "' " & _
'                          "AND idLote = '" & RsTempValeS.Fields("idLote") & "' "
'                Cn.Execute csql
'
'                RsTempValeS.MoveNext
'            Loop
'         End If
'
'         If Not RsTempValeI.EOF Then
'
'            RsTempValeI.MoveFirst
'            Do While Not RsTempValeI.EOF
'                strSigno = "-"
'                strCodAlmacen = traerCampo("valescab", "idalmacen", "idvalescab", Trim("" & RsTempValeI.Fields("idvalescab")), True, " tipoVale = 'I' ")
'
'                csql = "UPDATE productosalmacen " & _
'                        "SET CantidadStock = CantidadStock " & strSigno & " " & RsTempValeI.Fields("Cantidad") & " " & _
'                        "WHERE idEmpresa = '" & RsTempValeI.Fields("idEmpresa") & "' " & _
'                        "AND idSucursal = '" & RsTempValeI.Fields("idSucursal") & "' " & _
'                        "AND idAlmacen = '" & strCodAlmacen & "' " & _
'                        "AND idProducto = '" & RsTempValeI.Fields("idProducto") & "' " & _
'                        "AND idUMCompra = '" & RsTempValeI.Fields("idUM") & "' "
'                Cn.Execute csql
'
'                csql = "UPDATE productosalmacenporlote " & _
'                        "SET CantidadStock = CantidadStock " & strSigno & " " & RsTempValeI.Fields("Cantidad") & " " & _
'                        "WHERE idEmpresa = '" & RsTempValeI.Fields("idEmpresa") & "' " & _
'                          "AND idSucursal = '" & RsTempValeI.Fields("idSucursal") & "' " & _
'                          "AND idAlmacen = '" & strCodAlmacen & "' " & _
'                          "AND idProducto = '" & RsTempValeI.Fields("idProducto") & "' " & _
'                          "AND idUMCompra = '" & RsTempValeI.Fields("idUM") & "' " & _
'                          "AND idLote = '" & RsTempValeI.Fields("idLote") & "' "
'                Cn.Execute csql
'
'                RsTempValeI.MoveNext
'            Loop
'         End If
'    End If
'
'    '--- ACTUALIZAMOS STOCK EN LINEA
'     actualizaStock_Trans strCodValeSalida, 0, "S", StrMsgError, False, cIdSucursalOri
'     If StrMsgError <> "" Then GoTo Err
'
'     actualizaStock_Lote_Trans strCodValeSalida, 0, "S", StrMsgError, False, cIdSucursalOri
'     If StrMsgError <> "" Then GoTo Err
'
'    '--- ACTUALIZAMOS STOCK EN LINEA ValeIngreso
'     actualizaStock_Trans strCodValeIngreso, 0, "I", StrMsgError, False, cIdSucursalDes
'     If StrMsgError <> "" Then GoTo Err
'
'     actualizaStock_Lote_Trans strCodValeIngreso, 0, "I", StrMsgError, False, cIdSucursalDes
'     If StrMsgError <> "" Then GoTo Err
                 
    '--- Actualizamos los numeros de vales al registro
    csql = "UPDATE valestrans SET idValeIngreso = '" & strCodValeIngreso & "',idValeSalida = '" & strCodValeSalida & _
            "' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesTrans = '" & strValCod & "'"
    Cn.Execute csql
    
    If leeParametro("STOCK_POR_LOTE") = "S" Then
    
        CSqlC = "Delete From ValesDetLotes " & _
                "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & cIdSucursalOri & "' And TipoVale = 'S' And IdValesCab = '" & strCodValeSalida & "'"
        
        Cn.Execute CSqlC
        
        CSqlC = "Delete From ValesDetLotes " & _
                "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & cIdSucursalDes & "' And TipoVale = 'I' And IdValesCab = '" & strCodValeIngreso & "'"
        
        Cn.Execute CSqlC
        
        CSqlC = "Select A.* " & _
                "From ValesDet A " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & cIdSucursalOri & "' And A.TipoVale = 'S' " & _
                "And A.IdValesCab = '" & strCodValeSalida & "' " & _
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
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdPeriodoInv = '" & cIdPeriodoInvOri & "' And B.IdLote = '" & Trim("" & RsC.Fields("IdLote")) & "' " & _
                    "And A.IdAlmacen = '" & txtCod_AlmacenOrigen.Text & "' And A.EstValeCab <> 'ANU' " & _
                    "And CAST(A.FechaEmision AS DATE) <= CAST('" & Format(dtp_Emision.Value, "yyyy-mm-dd") & "' AS DATE) And B.IdProducto = '" & Trim("" & RsC.Fields("IdProducto")) & "' " & _
                    "Group By B.IdLote " & _
                    ""
                    
            rsl.Open CSqlC, Cn, adOpenKeyset, adLockReadOnly
            Do While Not rsl.EOF
                
                If Val("" & rsl.Fields("CantidadLote")) < (NCantidadProducto) Then
                    StrMsgError = "El producto " & Trim("" & RsC.Fields("IdProducto")) & " no cuenta con stock para La Talla " & Trim("" & RsC.Fields("NumLote")) & ",Verifique."
                    txtNum_ValeSalida.Text = ""
                    txtNum_ValeIngreso.Text = ""
                    txtCod_Vale.Text = ""
                    GoTo Err
                End If
                        
                NCantidadTotal = NCantidadTotal + NCantidadLote
                    
                CSqlC = "Insert Into ValesDetLotes(IdEmpresa,IdSucursal,TipoVale,IdValesCab,Item,IdLote,IdProducto,IdUM,Cantidad,CantidadAnt)Values(" & _
                        "'" & glsEmpresa & "','" & cIdSucursalOri & "','S','" & strCodValeSalida & "'," & Val("" & RsC.Fields("Item")) & "," & _
                        "'" & Trim("" & rsl.Fields("IdLote")) & "','" & Trim("" & RsC.Fields("IdProducto")) & "','" & Trim("" & RsC.Fields("IdUM")) & "'," & _
                        "" & NCantidadLote & ",0)"
                
                Cn.Execute CSqlC
                
                CSqlC = "Insert Into ValesDetLotes(IdEmpresa,IdSucursal,TipoVale,IdValesCab,Item,IdLote,IdProducto,IdUM,Cantidad,CantidadAnt)Values(" & _
                        "'" & glsEmpresa & "','" & cIdSucursalDes & "','I','" & strCodValeIngreso & "'," & Val("" & RsC.Fields("Item")) & "," & _
                        "'" & Trim("" & rsl.Fields("IdLote")) & "','" & Trim("" & RsC.Fields("IdProducto")) & "','" & Trim("" & RsC.Fields("IdUM")) & "'," & _
                        "" & NCantidadLote & ",0)"
                
                Cn.Execute CSqlC
                
                If NCantidadTotal >= NCantidadProducto Then
                    
                    Exit Do
                
                End If
                
                rsl.MoveNext
            
            Loop
            
            rsl.Close: Set rsl = Nothing
            
            RsC.MoveNext
            
        Loop
    
        RsC.Close: Set RsC = Nothing
            
    End If
    
    strVarCodValeIng = strCodValeIngreso
    strVarCodValeSal = strCodValeSalida
    
    Cn.CommitTrans
    
    CIdAlmacenOriAnt = txtCod_AlmacenOrigen.Text
    CIdAlmacenDesAnt = txtCod_AlmacenDestino.Text
        
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub

Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
    Exit Sub
    Resume
End Sub

Private Sub ImportarReceta(StrMsgError As String)
On Error GoTo Err
Dim CCod                        As String
Dim CSqlC                       As String
Dim RsC                         As New ADODB.Recordset
Dim RsGOri                      As New ADODB.Recordset
Dim RsGDes                      As New ADODB.Recordset
Dim dblVVUnit                   As Double
Dim dblIGVUnit                  As Double
Dim dblPVUnit                   As Double
Dim NFactorReceta               As Double
Dim IndExiste                   As Boolean

    FrmAyudaReceta.MostrarForm StrMsgError, CCod, NFactorReceta
    If StrMsgError <> "" Then GoTo Err
    
    If CCod <> "" Then
        
        IndExiste = False
        
        CIdReceta = CCod
        nfactor = NFactorReceta
                
        CSqlC = "Select A.IdProducto,B.GlsProducto,C.IdUM,C.GlsUM,B.AfectoIGV,A.Cantidad * " & nfactor & " Cantidad " & _
                "From RecetaDetOrigen A " & _
                "Inner Join Productos B " & _
                    "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                "Inner Join UnidadMedida C " & _
                    "On B.IdUMVenta = C.IdUM " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdReceta = '" & CIdReceta & "' " & _
                "Order By A.Item"
        RsC.Open CSqlC, Cn, adOpenKeyset, adLockReadOnly
        If Not RsC.EOF Then
        
            Do While Not RsC.EOF
                
                IndExiste = False
                
                If gDetalle.Dataset.RecordCount > 0 Then gDetalle.Dataset.First
                Do While Not gDetalle.Dataset.EOF
                    
                    If Trim("" & gDetalle.Columns.ColumnByFieldName("IdProducto").Value) = Trim("" & RsC.Fields("IdProducto")) Then
                        
                        IndExiste = True
                        
                        gDetalle.Dataset.Edit
                        gDetalle.Columns.ColumnByFieldName("Cantidad").Value = Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value) + Val("" & RsC.Fields("Cantidad"))
                        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value) * Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value)
                        gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value) * Val("" & gDetalle.Columns.ColumnByFieldName("IGVUnit").Value)
                        gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value) * Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value)
                        gDetalle.Dataset.Post
                        
                        Exit Do
                        
                    End If
                    
                    gDetalle.Dataset.Next
                    
                Loop
                
                If Not IndExiste Then
                    
                    gDetalle.Dataset.Insert
                    gDetalle.Dataset.Edit
                    
                    gDetalle.Columns.ColumnByFieldName("IdProducto").Value = "" & RsC.Fields("IdProducto")
                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = "" & RsC.Fields("GlsProducto")
                    gDetalle.Columns.ColumnByFieldName("IdUM").Value = "" & RsC.Fields("IdUM")
                    gDetalle.Columns.ColumnByFieldName("GlsUM").Value = "" & RsC.Fields("GlsUM")
                    gDetalle.Columns.ColumnByFieldName("Factor").Value = 1
                    gDetalle.Columns.ColumnByFieldName("Afecto").Value = Val("" & RsC.Fields("AfectoIGV"))
                    gDetalle.Columns.ColumnByFieldName("Cantidad").Value = Val("" & RsC.Fields("Cantidad"))
                    gDetalle.Columns.ColumnByFieldName("Cantidad2").Value = 0
                    
                    dblVVUnit = Val("" & traerCostoUnit(Trim("" & gDetalle.Columns.ColumnByFieldName("IdProducto").Value), Trim("" & txtCod_AlmacenOrigen.Text), Format(dtp_Emision.Value, "yyyy-mm-dd"), txtCod_Moneda.Text, StrMsgError))
                    If StrMsgError <> "" Then GoTo Err
                    
                    procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 0, dblVVUnit, gDetalle.Columns.ColumnByFieldName("Afecto").Value, dblVVUnit, dblIGVUnit, dblPVUnit
                    
                    gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
                    gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
                    gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
                    gDetalle.Columns.ColumnByFieldName("IdMoneda").Value = ""
                    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = Val("" & RsC.Fields("Cantidad")) * dblVVUnit
                    gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = Val("" & RsC.Fields("Cantidad")) * dblIGVUnit
                    gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = Val("" & RsC.Fields("Cantidad")) * dblPVUnit
                    
                    gDetalle.Columns.ColumnByFieldName("IdDocumentoImp").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSerieImp").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdDocVentasImp").Value = ""
                    
                    gDetalle.Dataset.Post
                    
                End If
                
                RsC.MoveNext
            
            Loop
        
        Else
        
            gDetalle.Dataset.Insert
                
        End If
        
        RsC.Close: Set RsC = Nothing
        
        gDetalle.Dataset.First
        Do While Not gDetalle.Dataset.EOF
            If Trim("" & gDetalle.Columns.ColumnByFieldName("IdProducto").Value) = "" Or Trim("" & gDetalle.Columns.ColumnByFieldName("GlsProducto").Value) = "" Then
                gDetalle.Dataset.Delete
            Else
                gDetalle.Dataset.Next
            End If
        Loop
        
        gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("IdProducto").Index
                
    End If
    
Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If RsGOri.State = 1 Then RsGOri.Close: Set RsGOri = Nothing
    If RsGDes.State = 1 Then RsGDes.Close: Set RsGDes = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

