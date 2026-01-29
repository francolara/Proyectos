VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.Ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmDocVentas_OC 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compras"
   ClientHeight    =   9630
   ClientLeft      =   7005
   ClientTop       =   2745
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   8865
      Left            =   90
      TabIndex        =   60
      Top             =   660
      Width           =   12960
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   4185
         Left            =   135
         OleObjectBlob   =   "frmDocVentas_OC.frx":0000
         TabIndex        =   3
         Top             =   990
         Width           =   12690
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   3435
         Left            =   120
         OleObjectBlob   =   "frmDocVentas_OC.frx":4F1C
         TabIndex        =   4
         Top             =   5295
         Width           =   12690
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   61
         Top             =   150
         Width           =   12690
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1050
            TabIndex        =   0
            Top             =   270
            Width           =   5925
            _ExtentX        =   10451
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
            Container       =   "frmDocVentas_OC.frx":869D
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Ano 
            Height          =   315
            Left            =   11490
            TabIndex        =   2
            Top             =   270
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDocVentas_OC.frx":86B9
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
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
            ItemData        =   "frmDocVentas_OC.frx":86D5
            Left            =   8085
            List            =   "frmDocVentas_OC.frx":8700
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   270
            Width           =   2025
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
            Left            =   11085
            TabIndex        =   80
            Top             =   320
            Width           =   300
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
            Left            =   7665
            TabIndex        =   79
            Top             =   320
            Width           =   300
         End
         Begin VB.Label Label1 
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
            Left            =   225
            TabIndex        =   62
            Top             =   320
            Width           =   735
         End
      End
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   13455
      Top             =   4815
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   13590
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   4125
      Index           =   0
      Left            =   2910
      TabIndex        =   143
      Top             =   3060
      Visible         =   0   'False
      Width           =   7110
      Begin VB.TextBox CATTextBox1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Index           =   0
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   146
         Top             =   315
         Width           =   6915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   3555
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   3540
         Width           =   1140
      End
   End
   Begin VB.Frame fraTotales 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   75
      TabIndex        =   24
      Top             =   8850
      Width           =   12915
      Begin CATControls.CATTextBox txt_TotalDsctoVV 
         Height          =   285
         Left            =   1830
         TabIndex        =   119
         Tag             =   "NTotalDsctoVV"
         Top             =   60
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   12640511
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
         Container       =   "frmDocVentas_OC.frx":8770
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalDsctoPV 
         Height          =   285
         Left            =   2760
         TabIndex        =   121
         Tag             =   "NTotalDsctoPV"
         Top             =   60
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   12640511
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
         Container       =   "frmDocVentas_OC.frx":878C
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalBruto 
         Height          =   315
         Left            =   5625
         TabIndex        =   82
         Tag             =   "NTotalValorVenta"
         Top             =   225
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmDocVentas_OC.frx":87A8
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalIGV 
         Height          =   315
         Left            =   8190
         TabIndex        =   83
         Tag             =   "NTotalIGVVenta"
         Top             =   225
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmDocVentas_OC.frx":87C4
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalNeto 
         Height          =   315
         Left            =   10950
         TabIndex        =   84
         Tag             =   "NTotalPrecioVenta"
         Top             =   225
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmDocVentas_OC.frx":87E0
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_MontoLetras 
         Height          =   285
         Left            =   75
         TabIndex        =   100
         Tag             =   "TtotalLetras"
         Top             =   375
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   33023
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
         Container       =   "frmDocVentas_OC.frx":87FC
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_SimboloMonBruto 
         Height          =   285
         Left            =   1050
         TabIndex        =   101
         Tag             =   "TsimboloMonBruto"
         Top             =   375
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   33023
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
         Container       =   "frmDocVentas_OC.frx":8818
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_DocReferencia 
         Height          =   285
         Left            =   4050
         TabIndex        =   108
         Tag             =   "TGlsDocReferencia"
         Top             =   375
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   33023
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
         Container       =   "frmDocVentas_OC.frx":8834
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_SimboloMonIGV 
         Height          =   285
         Left            =   2025
         TabIndex        =   109
         Tag             =   "TsimboloMonIGV"
         Top             =   375
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   33023
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
         Container       =   "frmDocVentas_OC.frx":8850
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_SimboloMonNeto 
         Height          =   285
         Left            =   3000
         TabIndex        =   110
         Tag             =   "TsimboloMonNeto"
         Top             =   375
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   33023
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
         Container       =   "frmDocVentas_OC.frx":886C
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalExonerado 
         Height          =   285
         Left            =   3570
         TabIndex        =   120
         Tag             =   "NTotalExonerado"
         Top             =   60
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         BackColor       =   12640511
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
         Container       =   "frmDocVentas_OC.frx":8888
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TotalBaseImponible 
         Height          =   285
         Left            =   5655
         TabIndex        =   122
         Tag             =   "NTotalBaseImponible"
         Top             =   225
         Visible         =   0   'False
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         BackColor       =   12640511
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
         Container       =   "frmDocVentas_OC.frx":88A4
         Text            =   "0"
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
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
         Left            =   5280
         TabIndex        =   99
         Top             =   270
         Visible         =   0   'False
         Width           =   285
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
         Left            =   7830
         TabIndex        =   98
         Top             =   270
         Visible         =   0   'False
         Width           =   330
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
         Left            =   10560
         TabIndex        =   97
         Top             =   270
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lbl_TotalLetras 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000007&
         Height          =   390
         Left            =   120
         TabIndex        =   92
         Top             =   225
         Visible         =   0   'False
         Width           =   4500
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
         Left            =   10095
         TabIndex        =   87
         Top             =   270
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lbl_TotalIGV 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "IGV"
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
         Left            =   7425
         TabIndex        =   86
         Top             =   270
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lbl_TotalBruto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   4665
         TabIndex        =   85
         Top             =   270
         Visible         =   0   'False
         Width           =   390
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   2170
      ButtonWidth     =   3016
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "            Nuevo            "
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
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anular"
            Object.ToolTipText     =   "Anular"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Excel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Importar"
            Object.ToolTipText     =   "Importar"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Correo"
            Object.ToolTipText     =   "Correo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraDetalle 
      ForeColor       =   &H00000000&
      Height          =   3390
      Left            =   75
      TabIndex        =   7
      Top             =   5520
      Width           =   12960
      Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
         Height          =   3090
         Left            =   60
         OleObjectBlob   =   "frmDocVentas_OC.frx":88C0
         TabIndex        =   71
         Top             =   240
         Width           =   12825
      End
   End
   Begin VB.Frame fraGeneral 
      ForeColor       =   &H00000000&
      Height          =   4815
      Left            =   90
      TabIndex        =   5
      Top             =   675
      Width           =   12960
      Begin VB.ComboBox cboPrioridad 
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
         Height          =   330
         ItemData        =   "frmDocVentas_OC.frx":161D7
         Left            =   3450
         List            =   "frmDocVentas_OC.frx":161E4
         TabIndex        =   153
         Tag             =   "TPrioridad"
         Text            =   "Combo1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton cmbcontactosclientes 
         Height          =   360
         Left            =   5445
         Picture         =   "frmDocVentas_OC.frx":161FB
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   2565
         Visible         =   0   'False
         Width           =   390
      End
      Begin CATControls.CATTextBox txtgls_contacto 
         Height          =   330
         Left            =   1845
         TabIndex        =   149
         Top             =   2565
         Visible         =   0   'False
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   582
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
         Container       =   "frmDocVentas_OC.frx":16585
      End
      Begin CATControls.CATTextBox txtCod_contacto 
         Height          =   330
         Left            =   900
         TabIndex        =   148
         Tag             =   "TIdContacto"
         Top             =   2565
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
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
         Container       =   "frmDocVentas_OC.frx":165A1
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.CommandButton cmbAyudaFormaPago 
         Height          =   315
         Left            =   5430
         Picture         =   "frmDocVentas_OC.frx":165BD
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   585
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton CmdAyudaUnidProduc 
         Height          =   315
         Left            =   11385
         Picture         =   "frmDocVentas_OC.frx":16947
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   1215
         Visible         =   0   'False
         Width           =   390
      End
      Begin CATControls.CATTextBox txtVal_Camision 
         Height          =   285
         Left            =   8550
         TabIndex        =   133
         Tag             =   "NComisionVtas"
         Top             =   4050
         Visible         =   0   'False
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
         MaxLength       =   11
         Container       =   "frmDocVentas_OC.frx":16CD1
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Partida2 
         Height          =   285
         Left            =   900
         TabIndex        =   127
         Tag             =   "TPartida2"
         Top             =   3900
         Visible         =   0   'False
         Width           =   4515
         _ExtentX        =   7964
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
         MaxLength       =   255
         Container       =   "frmDocVentas_OC.frx":16CED
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Llegada2 
         Height          =   285
         Left            =   900
         TabIndex        =   128
         Tag             =   "Tllegada2"
         Top             =   4200
         Visible         =   0   'False
         Width           =   4515
         _ExtentX        =   7964
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
         MaxLength       =   255
         Container       =   "frmDocVentas_OC.frx":16D09
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.CommandButton cmbAyudaTipoTicket 
         Height          =   315
         Left            =   5460
         Picture         =   "frmDocVentas_OC.frx":16D25
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   4170
         Visible         =   0   'False
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_TipoTicket 
         Height          =   285
         Left            =   885
         TabIndex        =   124
         Tag             =   "TidTipoTicket"
         Top             =   4170
         Visible         =   0   'False
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
         Container       =   "frmDocVentas_OC.frx":170AF
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TipoTicket 
         Height          =   285
         Left            =   1860
         TabIndex        =   125
         Top             =   4170
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
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
         Container       =   "frmDocVentas_OC.frx":170CB
         Vacio           =   -1  'True
      End
      Begin VB.CommandButton cmbAyudaVendedorCampo 
         Height          =   315
         Left            =   5460
         Picture         =   "frmDocVentas_OC.frx":170E7
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   4500
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaCentroCosto 
         Height          =   315
         Left            =   11340
         Picture         =   "frmDocVentas_OC.frx":17471
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   3840
         Visible         =   0   'False
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_CentroCosto 
         Height          =   285
         Left            =   6765
         TabIndex        =   112
         Tag             =   "TidCentroCosto"
         Top             =   3840
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":177FB
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_CentroCosto 
         Height          =   285
         Left            =   7740
         TabIndex        =   113
         Tag             =   "TGlsCentroCosto"
         Top             =   3840
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":17817
         Vacio           =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtp_Pago 
         Height          =   315
         Left            =   10125
         TabIndex        =   104
         Tag             =   "FFecPago"
         Top             =   3630
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   132317185
         CurrentDate     =   38955
      End
      Begin MSComCtl2.DTPicker dtp_IniTraslado 
         Height          =   315
         Left            =   6825
         TabIndex        =   102
         Tag             =   "FFecIniTraslado"
         Top             =   3675
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   132317185
         CurrentDate     =   38955
      End
      Begin CATControls.CATTextBox txtObs 
         Height          =   510
         Left            =   6450
         TabIndex        =   81
         Tag             =   "TObsDocVentas"
         Top             =   2790
         Visible         =   0   'False
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   900
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
         Container       =   "frmDocVentas_OC.frx":17833
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.CommandButton cmbAyudaMotivoNCD 
         Height          =   315
         Left            =   11340
         Picture         =   "frmDocVentas_OC.frx":1784F
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   4320
         Visible         =   0   'False
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_MotivoNCD 
         Height          =   285
         Left            =   6750
         TabIndex        =   94
         Tag             =   "TidMotivoNCD"
         Top             =   4200
         Visible         =   0   'False
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
         Locked          =   -1  'True
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":17BD9
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_MotivoNCD 
         Height          =   285
         Left            =   7725
         TabIndex        =   95
         Tag             =   "TGlsMotivoNCD"
         Top             =   4200
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":17BF5
         Vacio           =   -1  'True
      End
      Begin VB.CommandButton cmbAyudaMotivoTraslado 
         Height          =   315
         Left            =   11325
         Picture         =   "frmDocVentas_OC.frx":17C11
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   4425
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaLista 
         Height          =   315
         Left            =   11325
         Picture         =   "frmDocVentas_OC.frx":17F9B
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   4050
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaVendedor 
         Height          =   315
         Left            =   11400
         Picture         =   "frmDocVentas_OC.frx":18325
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1860
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaAlmacen 
         Height          =   315
         Left            =   11400
         Picture         =   "frmDocVentas_OC.frx":186AF
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1515
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   11400
         Picture         =   "frmDocVentas_OC.frx":18A39
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2205
         Visible         =   0   'False
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   285
         Left            =   900
         TabIndex        =   38
         Tag             =   "TidPerCliente"
         Top             =   900
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":18DC3
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_NumDoc 
         Height          =   315
         Left            =   10350
         TabIndex        =   36
         Tag             =   "TidDocVentas"
         Top             =   150
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         FontBold        =   -1  'True
         FontName        =   "Arial"
         FontSize        =   9
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":18DDF
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.CommandButton cmbAyudaVehiculo 
         Height          =   315
         Left            =   5445
         Picture         =   "frmDocVentas_OC.frx":18DFB
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3375
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaEmpTrans 
         Height          =   315
         Left            =   5445
         Picture         =   "frmDocVentas_OC.frx":19185
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3075
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaChofer 
         Height          =   315
         Left            =   5445
         Picture         =   "frmDocVentas_OC.frx":1950F
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2400
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaCliente 
         Height          =   315
         Left            =   5440
         Picture         =   "frmDocVentas_OC.frx":19899
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   870
         Visible         =   0   'False
         Width           =   390
      End
      Begin MSComCtl2.DTPicker dtp_Emision 
         Height          =   315
         Left            =   10275
         TabIndex        =   12
         Tag             =   "FFecEmision"
         Top             =   825
         Visible         =   0   'False
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   132317185
         CurrentDate     =   38955
      End
      Begin MSComctlLib.ImageList imgDocVentas 
         Left            =   6075
         Top             =   270
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
               Picture         =   "frmDocVentas_OC.frx":19C23
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":19FBD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1A40F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1A7A9
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1AB43
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1AEDD
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1B277
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1B611
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1B9AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1BD45
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1C0DF
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1CDA1
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1D13B
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1D58D
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1D927
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDocVentas_OC.frx":1E339
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin CATControls.CATTextBox txt_Serie 
         Height          =   315
         Left            =   8220
         TabIndex        =   37
         Tag             =   "TidSerie"
         Top             =   150
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial"
         FontSize        =   9.75
         ForeColor       =   -2147483640
         MaxLength       =   3
         Container       =   "frmDocVentas_OC.frx":1EA0B
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   285
         Left            =   1875
         TabIndex        =   39
         Tag             =   "TGlsCliente"
         Top             =   900
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
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
         Container       =   "frmDocVentas_OC.frx":1EA27
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Chofer 
         Height          =   285
         Left            =   900
         TabIndex        =   40
         Tag             =   "TidPerChofer"
         Top             =   2400
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":1EA43
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Chofer 
         Height          =   285
         Left            =   1875
         TabIndex        =   41
         Tag             =   "TglsChofer"
         Top             =   2400
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":1EA5F
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_EmpTrans 
         Height          =   285
         Left            =   900
         TabIndex        =   42
         Tag             =   "TidPerEmpTrans"
         Top             =   3075
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":1EA7B
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_EmpTrans 
         Height          =   285
         Left            =   1875
         TabIndex        =   43
         Tag             =   "TGlsEmpTrans"
         Top             =   3075
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":1EA97
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Vehiculo 
         Height          =   285
         Left            =   900
         TabIndex        =   44
         Tag             =   "TidVehiculo"
         Top             =   3375
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":1EAB3
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vehiculo 
         Height          =   285
         Left            =   1875
         TabIndex        =   45
         Tag             =   "TGlsVehiculo"
         Top             =   3375
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":1EACF
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Almacen 
         Height          =   285
         Left            =   6840
         TabIndex        =   46
         Tag             =   "TidAlmacen"
         Top             =   1530
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":1EAEB
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Almacen 
         Height          =   285
         Left            =   7785
         TabIndex        =   47
         Top             =   1530
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":1EB07
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   285
         Left            =   6840
         TabIndex        =   48
         Tag             =   "TidPerVendedor"
         Top             =   1860
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":1EB23
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   285
         Left            =   7800
         TabIndex        =   49
         Tag             =   "TGlsVendedor"
         Top             =   1860
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":1EB3F
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   285
         Left            =   6825
         TabIndex        =   50
         Tag             =   "TidMoneda"
         Top             =   2205
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":1EB5B
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   285
         Left            =   7800
         TabIndex        =   51
         Tag             =   "Tglsmoneda"
         Top             =   2205
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":1EB77
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_RUC 
         Height          =   285
         Left            =   900
         TabIndex        =   52
         Tag             =   "TRUCCliente"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
         MaxLength       =   11
         Container       =   "frmDocVentas_OC.frx":1EB93
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Direccion 
         Height          =   285
         Left            =   900
         TabIndex        =   53
         Tag             =   "TdirCliente"
         Top             =   1500
         Visible         =   0   'False
         Width           =   4515
         _ExtentX        =   7964
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
         MaxLength       =   255
         Container       =   "frmDocVentas_OC.frx":1EBAF
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Partida 
         Height          =   285
         Left            =   900
         TabIndex        =   54
         Tag             =   "TPartida"
         Top             =   1800
         Visible         =   0   'False
         Width           =   4515
         _ExtentX        =   7964
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
         MaxLength       =   255
         Container       =   "frmDocVentas_OC.frx":1EBCB
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Llegada 
         Height          =   285
         Left            =   900
         TabIndex        =   55
         Tag             =   "Tllegada"
         Top             =   2100
         Visible         =   0   'False
         Width           =   4515
         _ExtentX        =   7964
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
         MaxLength       =   255
         Container       =   "frmDocVentas_OC.frx":1EBE7
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Brevete 
         Height          =   285
         Left            =   900
         TabIndex        =   56
         Tag             =   "TBrevete"
         Top             =   2700
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
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
         MaxLength       =   45
         Container       =   "frmDocVentas_OC.frx":1EC03
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Placa 
         Height          =   285
         Left            =   900
         TabIndex        =   57
         Tag             =   "TPlaca"
         Top             =   3675
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
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
         MaxLength       =   10
         Container       =   "frmDocVentas_OC.frx":1EC1F
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Bultos 
         Height          =   285
         Left            =   3750
         TabIndex        =   58
         Tag             =   "NBultos"
         Top             =   3675
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmDocVentas_OC.frx":1EC3B
         Estilo          =   3
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TipoCambio 
         Height          =   285
         Left            =   6825
         TabIndex        =   59
         Tag             =   "NTipoCambio"
         Top             =   855
         Visible         =   0   'False
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmDocVentas_OC.frx":1EC57
         Text            =   "0"
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Marca 
         Height          =   285
         Left            =   900
         TabIndex        =   63
         Tag             =   "TMarca"
         Top             =   3975
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
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
         MaxLength       =   128
         Container       =   "frmDocVentas_OC.frx":1EC73
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Modelo 
         Height          =   285
         Left            =   3750
         TabIndex        =   65
         Tag             =   "TModelo"
         Top             =   3975
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         MaxLength       =   128
         Container       =   "frmDocVentas_OC.frx":1EC8F
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Color 
         Height          =   285
         Left            =   900
         TabIndex        =   67
         Tag             =   "TColor"
         Top             =   4125
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
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
         MaxLength       =   128
         Container       =   "frmDocVentas_OC.frx":1ECAB
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_CodInscripcion 
         Height          =   285
         Left            =   3750
         TabIndex        =   69
         Tag             =   "TCodInsCrip"
         Top             =   4125
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         MaxLength       =   128
         Container       =   "frmDocVentas_OC.frx":1ECC7
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Lista 
         Height          =   285
         Left            =   6615
         TabIndex        =   76
         Tag             =   "TidLista"
         Top             =   4050
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":1ECE3
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Lista 
         Height          =   285
         Left            =   7725
         TabIndex        =   77
         Top             =   4050
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
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
         Container       =   "frmDocVentas_OC.frx":1ECFF
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_MotivoTraslado 
         Height          =   285
         Left            =   6750
         TabIndex        =   89
         Tag             =   "TidMotivoTraslado"
         Top             =   4425
         Visible         =   0   'False
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
         Locked          =   -1  'True
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":1ED1B
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_MotivoTraslado 
         Height          =   285
         Left            =   7725
         TabIndex        =   90
         Top             =   4425
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":1ED37
         Vacio           =   -1  'True
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gDocReferencia 
         Height          =   1200
         Left            =   6000
         OleObjectBlob   =   "frmDocVentas_OC.frx":1ED53
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   5685
      End
      Begin CATControls.CATTextBox txt_RUCEmp 
         Height          =   285
         Left            =   7815
         TabIndex        =   106
         Tag             =   "TrucEmpTrans"
         Top             =   4320
         Visible         =   0   'False
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
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   11
         Container       =   "frmDocVentas_OC.frx":213F8
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_VendedorCampo 
         Height          =   285
         Left            =   885
         TabIndex        =   116
         Tag             =   "TidPerVendedorCampo"
         Top             =   4500
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":21414
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_VendedorCampo 
         Height          =   285
         Left            =   1860
         TabIndex        =   117
         Tag             =   "TGlsVendedorCampo"
         Top             =   4500
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":21430
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_OrdenCompra 
         Height          =   285
         Left            =   900
         TabIndex        =   131
         Tag             =   "TnumOrdenCompra"
         Top             =   2850
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
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
         MaxLength       =   25
         Container       =   "frmDocVentas_OC.frx":2144C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_UnidProd 
         Height          =   285
         Left            =   6840
         TabIndex        =   136
         Tag             =   "Tidupp"
         Top             =   1215
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":21468
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_UnidProd 
         Height          =   285
         Left            =   7785
         TabIndex        =   137
         Top             =   1215
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":21484
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_FormaPago 
         Height          =   285
         Left            =   900
         TabIndex        =   140
         Tag             =   "TidFormaPago"
         Top             =   585
         Visible         =   0   'False
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
         MaxLength       =   8
         Container       =   "frmDocVentas_OC.frx":214A0
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_FormaPago 
         Height          =   285
         Left            =   1875
         TabIndex        =   141
         Tag             =   "TglsFormaPago"
         Top             =   585
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
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
         Container       =   "frmDocVentas_OC.frx":214BC
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsPlaca 
         Height          =   285
         Left            =   900
         TabIndex        =   151
         Tag             =   "TGlsPlaca"
         Top             =   2745
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
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
         MaxLength       =   15
         Container       =   "frmDocVentas_OC.frx":214D8
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label LblPlaca 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "UP. y/o Placa"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   -945
         TabIndex        =   152
         Top             =   60
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblcontacto 
         Caption         =   "Contacto"
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
         Left            =   45
         TabIndex        =   147
         Top             =   2610
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lbl_FormaPago 
         Appearance      =   0  'Flat
         Caption         =   "F. Pago"
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
         Left            =   90
         TabIndex        =   142
         Top             =   645
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lbl_upp 
         Appearance      =   0  'Flat
         Caption         =   "UUPP"
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
         Left            =   5985
         TabIndex        =   138
         Top             =   1260
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lbl_Comision 
         Appearance      =   0  'Flat
         Caption         =   "RUC Emp:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   8850
         TabIndex        =   134
         Top             =   3225
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lbl_OrdenCompra 
         Appearance      =   0  'Flat
         Caption         =   "Nº Sobre"
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
         Left            =   75
         TabIndex        =   132
         Top             =   2925
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lbl_Partida2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Partida 2"
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
         Left            =   75
         TabIndex        =   130
         Top             =   3900
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lbl_Llegada2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Llegada 2"
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
         Left            =   75
         TabIndex        =   129
         Top             =   4200
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lbl_TipoTicket 
         Appearance      =   0  'Flat
         Caption         =   "Tipo Ticket"
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
         Left            =   30
         TabIndex        =   126
         Top             =   4245
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbl_VendedorCampo 
         Appearance      =   0  'Flat
         Caption         =   "V. Campo"
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
         Left            =   60
         TabIndex        =   118
         Top             =   4575
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_CentroCosto 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6000
         TabIndex        =   114
         Top             =   3900
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lbl_RUCEmp 
         Appearance      =   0  'Flat
         Caption         =   "RUC Emp:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6840
         TabIndex        =   107
         Top             =   4320
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lbl_FechaPago 
         Appearance      =   0  'Flat
         Caption         =   "Fec. Pago"
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
         Left            =   8925
         TabIndex        =   105
         Top             =   3660
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl_IniTraslado 
         Appearance      =   0  'Flat
         Caption         =   " Ini. Tras"
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
         Left            =   6000
         TabIndex        =   103
         Top             =   3750
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Obs 
         Appearance      =   0  'Flat
         Caption         =   "Obs."
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
         TabIndex        =   14
         Top             =   4050
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_MotivoNCD 
         Appearance      =   0  'Flat
         Caption         =   "Motivo"
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
         Left            =   6000
         TabIndex        =   96
         Top             =   4275
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_MotivoTraslado 
         Appearance      =   0  'Flat
         Caption         =   "Motivo"
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
         Left            =   5985
         TabIndex        =   91
         Top             =   4500
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Lista 
         Appearance      =   0  'Flat
         Caption         =   "Lista:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6000
         TabIndex        =   78
         Top             =   4125
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_CodInscrip 
         Appearance      =   0  'Flat
         Caption         =   "Cod. Ins:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2850
         TabIndex        =   70
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl_Color 
         Appearance      =   0  'Flat
         Caption         =   "Color:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   75
         TabIndex        =   68
         Top             =   4200
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Modelo 
         Appearance      =   0  'Flat
         Caption         =   "Modelo:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2850
         TabIndex        =   66
         Top             =   4050
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbl_Marca 
         Appearance      =   0  'Flat
         Caption         =   "Marca"
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
         Left            =   75
         TabIndex        =   64
         Top             =   4050
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Bultos 
         Appearance      =   0  'Flat
         Caption         =   "Bultos:"
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
         Left            =   2850
         TabIndex        =   35
         Top             =   3750
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Vehiculo 
         Appearance      =   0  'Flat
         Caption         =   "Vehículo"
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
         Left            =   75
         TabIndex        =   34
         Top             =   3405
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_EmpTrans 
         Appearance      =   0  'Flat
         Caption         =   "Emp Trans."
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
         Left            =   75
         TabIndex        =   32
         Top             =   3105
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Brevete 
         Appearance      =   0  'Flat
         Caption         =   "Brevete"
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
         Left            =   75
         TabIndex        =   30
         Top             =   2775
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Placa 
         Appearance      =   0  'Flat
         Caption         =   "Placa"
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
         Left            =   75
         TabIndex        =   29
         Top             =   3750
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Chofer 
         Appearance      =   0  'Flat
         Caption         =   "Chofer"
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
         Left            =   75
         TabIndex        =   28
         Top             =   2430
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Llegada 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Llegada"
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
         Left            =   75
         TabIndex        =   26
         Top             =   2100
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl_Partida 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Partida"
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
         Left            =   75
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl_Almacen 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6000
         TabIndex        =   21
         Top             =   1590
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_RUC 
         Appearance      =   0  'Flat
         Caption         =   "R.U.C."
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
         Left            =   75
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Direccion 
         Appearance      =   0  'Flat
         Caption         =   "Dirección"
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
         Left            =   75
         TabIndex        =   18
         Top             =   1500
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Vendedor 
         Appearance      =   0  'Flat
         Caption         =   "Responsable"
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
         Left            =   6000
         TabIndex        =   17
         Top             =   1935
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_TC 
         Appearance      =   0  'Flat
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
         Height          =   240
         Left            =   6000
         TabIndex        =   16
         Top             =   900
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6000
         TabIndex        =   15
         Top             =   2280
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_Cliente 
         Appearance      =   0  'Flat
         Caption         =   "Proveedor"
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
         Left            =   75
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_NumDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   225
         Left            =   9630
         TabIndex        =   11
         Top             =   225
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbl_FechaEmision 
         Appearance      =   0  'Flat
         Caption         =   "Fecha Emisión"
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
         Left            =   9075
         TabIndex        =   10
         Top             =   900
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lbl_Serie 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   7695
         TabIndex        =   9
         Top             =   225
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblDoc 
         Appearance      =   0  'Flat
         Caption         =   "Orden de Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   165
         TabIndex        =   8
         Top             =   225
         Width           =   5220
      End
      Begin VB.Label lblPrioridad 
         Appearance      =   0  'Flat
         Caption         =   "Prioridad"
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
         Left            =   2790
         TabIndex        =   154
         Top             =   1200
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Almacen:"
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   6075
      TabIndex        =   22
      Top             =   1830
      Width           =   765
   End
End
Attribute VB_Name = "frmDocVentas_OC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strTipoDoc           As String
Dim objDocVentas            As New clsDocVentas
Dim intBoton                As Integer
Dim indCargando             As Boolean
Dim indNuevoDoc             As Boolean
Private indInserta          As Boolean
Private indInsertaDocRef    As Boolean
Private strEstDocVentas     As String
Private indGeneraVale       As Boolean
Private strGlsTipoDoc       As String
Dim intRegMax               As Integer
Dim dblPorDsctoEspecial     As Double
Dim dblIgvNEw               As Double
Dim rucEmpresa              As String

Private Sub cbx_Mes_Click()
On Error GoTo Err
Dim StrMsgError As String

    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaAlmacen_Click()
    
    mostrarAyuda "ALMACEN", txtCod_Almacen, txtGls_Almacen
    If txtCod_Almacen.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaCentroCosto_Click()
    
    mostrarAyuda "CENTROCOSTO", txtCod_CentroCosto, txtGls_CentroCosto
    'If txtCod_CentroCosto.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaChofer_Click()

    mostrarAyuda "CHOFER", txtCod_Chofer, txtGls_Chofer
    If txtCod_Chofer.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaEmpTrans_Click()

    mostrarAyuda "EMPTRANS", txtCod_EmpTrans, txtGls_EmpTrans
    If txtCod_EmpTrans.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaFormaPago_Click()

    mostrarAyuda "FORMASPAGO", txtCod_FormaPago, txtGls_FormaPago, "and visible_ventas in('C','A') "
    
End Sub

Private Sub cmbAyudaLista_Click()

    mostrarAyuda "LISTAPRECIOS", txtCod_Lista, txtGls_Lista
    If txtCod_Lista.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaMoneda_Click()

    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    'If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaMotivoNCD_Click()

    mostrarAyuda "MOTIVONCD", txtCod_MotivoNCD, txtGls_MotivoNCD
    If txtCod_MotivoNCD.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaMotivoTraslado_Click()

    mostrarAyuda "MOTIVOTRASLADO", txtCod_MotivoTraslado, txtGls_MotivoTraslado
    If txtCod_MotivoTraslado.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaTipoTicket_Click()

    mostrarAyuda "TIPOTICKET", txtCod_TipoTicket, txtGls_TipoTicket
    If txtCod_TipoTicket.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaVehiculo_Click()

    mostrarAyuda "VEHICULO", txtCod_Vehiculo, txtGls_Vehiculo
    If txtCod_Vehiculo.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaVendedor_Click()

    'mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor
    mostrarAyuda "USUARIOS", txtCod_Vendedor, txtGls_Vendedor
    'If txtCod_Vendedor.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaVendedorCampo_Click()
    
    If glsModVendCampo = False Then Exit Sub
    mostrarAyuda "VENDEDOR", txtCod_VendedorCampo, txtGls_VendedorCampo
    If txtCod_VendedorCampo.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbcontactosclientes_Click()

    mostrarAyuda "CONTACTOSPROVEEDORES", txtCod_contacto, txtgls_contacto, "AND C.IDPROVEEDOR='" & txtCod_Cliente.Text & "'"

End Sub

Private Sub CmdAyudaUnidProduc_Click()
    
    mostrarAyuda "UNIDADPRODUC", txtCod_UnidProd, txtGls_UnidProd
    If txtCod_UnidProd.Text <> "" Then SendKeys "{tab}"
 
End Sub

Private Sub Command1_Click(Index As Integer)
    
    gDetalle.Dataset.Edit
    gDetalle.Columns.ColumnByFieldName("glsProducto").Value = CATTextBox1(0).Text
    gDetalle.Dataset.Post
    Frame2(0).Visible = False

End Sub

Private Sub Command2_Click(Index As Integer)
    
    Frame2(0).Visible = False

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
    
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If

End Sub
    
Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    If strTipoDoc <> "94" Then
         cbx_Mes.RemoveItem (12)
    End If
    
    If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Then
        lbl_FechaPago.Caption = "Fec. Entrega"
        lbl_Llegada.Caption = "Lugar de Entrega"
        lbl_Partida.Caption = "Plazo de Entrega"
    End If
    
    indInserta = False
    indInsertaDocRef = False
    indNuevoDoc = True
    Me.top = 0
    Me.left = 0
    
    strEstDocVentas = "GEN"
    txt_Ano.Text = Year(getFechaSistema)
    cbx_Mes.ListIndex = Month(getFechaSistema) - 1
    
    strGlsTipoDoc = traerCampo("documentos", "GlsDocumento", "idDocumento", strTipoDoc, False)
    Me.Caption = strGlsTipoDoc
    
    lblDoc.Caption = Me.Caption
    muestraControlesCabecera
    
    lbl_Serie.Visible = False
    txt_Serie.Visible = False
    lbl_Lista.Visible = False
    txtCod_Lista.Visible = False
    txtGls_Lista.Visible = False
    cmbAyudaLista.Visible = False
    
    txt_TipoCambio.Decimales = glsDecimalesTC
    muestraColumnasDetalle
    
    ConfGrid gLista, False, False, False, False
    ConfGrid gListaDetalle, False, False, False, False
    ConfGrid gDetalle, True, False, False, False
    ConfGrid gDocReferencia, True, False, False, True
    
    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    fraDetalle.Visible = False
    fraTotales.Visible = False
    habilitaBotones 8
    indNuevoDoc = False
    
    If leeParametro("NO_VALIDAR_CONTACTO_OC") = "1" Then
        txtCod_contacto.Vacio = True
        txtgls_contacto.Vacio = True
    End If
    
    If strTipoDoc = "OS" Then
        gLista.Columns.ColumnByFieldName("Referencia").Visible = True
        gLista.Columns.ColumnByFieldName("GlsPlaca").Visible = True
    Else
        gLista.Columns.ColumnByFieldName("Referencia").Visible = False
        gLista.Columns.ColumnByFieldName("GlsPlaca").Visible = False
    End If
    
    If strTipoDoc = "94" Then
        If leeParametro("VISUALIZA_CENTROCOSTO_OC") = "1" Then
            gLista.Columns.ColumnByFieldName("IdCentroCosto").Visible = True
            gLista.Columns.ColumnByFieldName("GlsUser").Visible = True
        Else
            gLista.Columns.ColumnByFieldName("IdCentroCosto").Visible = False
            gLista.Columns.ColumnByFieldName("GlsUser").Visible = False
        End If
    End If
    
    rucEmpresa = traerCampo("Empresas", "RUC", "idEmpresa", glsEmpresa, False)
    
    'Si la Empresa es INMAC y el documento es 87 la columna del detalle idDocVentasPres cambiamos el tipo de boton
    If rucEmpresa = "20513250445" And strTipoDoc = "87" Then
        gDetalle.Columns.ColumnByFieldName("idDocVentasPres").ColumnType = gedTextEdit
    End If
    
    cboPrioridad.ListIndex = 0
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo As String
Dim strMsg      As String

    getEstadoCierreMes Format(dtp_Emision.Value, "dd/mm/yyyy"), StrMsgError
    If StrMsgError <> "" Then GoTo Err

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    eliminaNulosGrilla

    If gDetalle.Count >= 1 Then
        If gDetalle.Count = 1 And (gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" Or gDetalle.Columns.ColumnByFieldName("Cantidad").Value <= 0) Then
            StrMsgError = "Falta Ingresar Detalle"
            GoTo Err
        End If
    End If
    
    eliminaNulosGrillaDocRef
    generaSTRDocReferencia
    txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
    
    If strTipoDoc = "OS" And leeParametro("VALIDAPRESUPUESTO") = "1" Then
        gDetalle.Dataset.First
        Do While Not gDetalle.Dataset.EOF
            If Len(Trim(gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value)) > 0 And Len(Trim(gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value)) > 0 And Len(Trim(gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value)) = 0 Then
                StrMsgError = "El Item " & gDetalle.Columns.ColumnByFieldName("Item").Value & " tiene asignado Hoja de Costo pero falta asignar el Presupuesto": GoTo Err
            End If
            gDetalle.Dataset.Next
        Loop
    End If
    
    If strTipoDoc = "OS" And leeParametro("VALIDA_HC_OS") = "1" Then
        gDetalle.Dataset.First
        Do While Not gDetalle.Dataset.EOF
            If Len(Trim(gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value)) = 0 Then
                StrMsgError = "El Item " & gDetalle.Columns.ColumnByFieldName("Item").Value & " no tiene asignado Hoja de Costo "
            End If
            gDetalle.Dataset.Next
        Loop
    End If
    
    If intBoton = 1 Then 'graba
        objDocVentas.EjecutaSQLFormDocVentas_OC Me, 0, StrMsgError, strTipoDoc, txt_Serie.Text, gDetalle, gDocReferencia, indGeneraVale
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabo"
        intBoton = 3
    
    Else 'modifica
        objDocVentas.EjecutaSQLFormDocVentas_OC Me, 1, StrMsgError, strTipoDoc, txt_Serie.Text, gDetalle, gDocReferencia, indGeneraVale
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modifico"
    End If
    
    fraGeneral.Enabled = False
    fraDetalle.Enabled = False
    habilitaBotones 2
    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo(ByRef StrMsgError As String)
Dim rsg As New ADODB.Recordset
Dim RsD As New ADODB.Recordset
    
    If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Then indGeneraVale = False
    limpiaForm Me
    
    strEstDocVentas = "GEN"
    lblDoc.Caption = traerCampo("documentos", "GlsDocumento", "idDocumento", strTipoDoc, False)
    lblDoc.ForeColor = &H0&
    txt_Serie.Enabled = True
    txt_NumDoc.Enabled = True
    fraGeneral.Enabled = True
    fraDetalle.Enabled = True
    
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
    
    gDocReferencia.Columns.FocusedIndex = gDocReferencia.Columns.ColumnByFieldName("idDocumento").Index
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "idCodFabricante", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idMarca", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsMarca", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVBruto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVBruto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PorDcto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "DctoVV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "DctoPV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idTipoProducto", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idDocumentoImp", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "idDocVentasImp", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerieImp", adVarChar, 4, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idUsuarioDcto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "VVUnitLista", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnitLista", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnitNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnitNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IdCentroCosto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdSucursalPres", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdDocumentoPres", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "IdSeriePres", adVarChar, 3, adFldIsNullable
    rsg.Fields.Append "IdDocVentasPres", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "FechaEmision", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "GlsPlaca", adVarChar, 50, adFldIsNullable
    rsg.Open
    
    rsg.AddNew
    rsg.Fields("Item") = 1
    rsg.Fields("idProducto") = ""
    rsg.Fields("CodigoRapido") = ""
    rsg.Fields("idCodFabricante") = ""
    rsg.Fields("GlsProducto") = ""
    rsg.Fields("idMarca") = ""
    rsg.Fields("GlsMarca") = ""
    rsg.Fields("idUM") = ""
    rsg.Fields("GlsUM") = ""
    rsg.Fields("Factor") = 1
    rsg.Fields("Afecto") = 1
    rsg.Fields("Cantidad") = 0
    rsg.Fields("VVUnit") = 0
    rsg.Fields("IGVUnit") = 0
    rsg.Fields("PVUnit") = 0
    rsg.Fields("TotalVVBruto") = 0
    rsg.Fields("TotalPVBruto") = 0
    rsg.Fields("PorDcto") = 0
    rsg.Fields("DctoVV") = 0
    rsg.Fields("DctoPV") = 0
    rsg.Fields("TotalVVNeto") = 0
    rsg.Fields("TotalIGVNeto") = 0
    rsg.Fields("TotalPVNeto") = 0
    rsg.Fields("VVUnitLista") = 0
    rsg.Fields("PVUnitLista") = 0
    rsg.Fields("VVUnitNeto") = 0
    rsg.Fields("PVUnitNeto") = 0
    rsg.Fields("IdCentroCosto") = ""
    rsg.Fields("IdSucursalPres") = ""
    rsg.Fields("IdDocumentoPres") = ""
    rsg.Fields("IdSeriePres") = ""
    rsg.Fields("IdDocVentasPres") = ""
    rsg.Fields("FechaEmision") = getFechaSistema
    rsg.Fields("GlsPlaca") = ""
    
    Set gDetalle.DataSource = Nothing
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("idProducto").Index
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gdetalle_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)

    If Action = daInsert Then
        gDetalle.Columns.ColumnByFieldName("item").Value = gDetalle.Count
        gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
        gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = ""
        gDetalle.Columns.ColumnByFieldName("idCodFabricante").Value = ""
        gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
        gDetalle.Columns.ColumnByFieldName("idMarca").Value = ""
        gDetalle.Columns.ColumnByFieldName("GlsMarca").Value = ""
        gDetalle.Columns.ColumnByFieldName("idUM").Value = ""
        gDetalle.Columns.ColumnByFieldName("GlsUM").Value = ""
        gDetalle.Columns.ColumnByFieldName("Factor").Value = 1
        gDetalle.Columns.ColumnByFieldName("Afecto").Value = 1
        gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
        gDetalle.Columns.ColumnByFieldName("VVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("PVUnit").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalVVBruto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalPVBruto").Value = 0
        gDetalle.Columns.ColumnByFieldName("PorDcto").Value = 0
        gDetalle.Columns.ColumnByFieldName("DctoVV").Value = 0
        gDetalle.Columns.ColumnByFieldName("DctoPV").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = 0
        gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = 0
        gDetalle.Columns.ColumnByFieldName("VVUnitNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("PVUnitNeto").Value = 0
        gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
        gDetalle.Columns.ColumnByFieldName("FechaEmision").Value = getFechaSistema
        gDetalle.Columns.ColumnByFieldName("GlsPlaca").Value = ""
        gDetalle.Dataset.Post
    End If

End Sub

Private Sub gdetalle_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If (gDetalle.Columns.ColumnByFieldName("idProducto").Value = "") And indInserta = False Then
            Allow = False
        Else
            If intRegMax = 0 Or gDetalle.Count < intRegMax Then
                gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("idProducto").ColIndex
            Else
                Allow = False
            End If
        End If
    End If

End Sub

Private Sub gDetalle_OnChangeColumn(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal OldColumn As DXDBGRIDLibCtl.IdxGridColumn, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn)

    If gDetalle.Columns.FocusedAbsoluteIndex = gDetalle.Columns.ColumnByFieldName("GlsProducto").Index Or gDetalle.Columns.FocusedAbsoluteIndex = gDetalle.Columns.ColumnByFieldName("PVUnit").Index Or gDetalle.Columns.FocusedAbsoluteIndex = gDetalle.Columns.ColumnByFieldName("VVUnit").Index Then
        If gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value = "06002" Then 'Servicios
            gDetalle.Columns.ColumnByFieldName("GlsProducto").DisableEditor = False
            gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = False
            gDetalle.Columns.ColumnByFieldName("PVUnit").DisableEditor = False
        Else
            gDetalle.Columns.ColumnByFieldName("GlsProducto").DisableEditor = True
        End If
    End If

End Sub

Private Sub gDetalle_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrMsgError                             As String
On Error GoTo Err
    
    If gDetalle.Columns.FocusedAbsoluteIndex = gDetalle.Columns.ColumnByFieldName("GlsProducto").Index Or gDetalle.Columns.FocusedAbsoluteIndex = gDetalle.Columns.ColumnByFieldName("PVUnit").Index Or gDetalle.Columns.FocusedAbsoluteIndex = gDetalle.Columns.ColumnByFieldName("VVUnit").Index Then
        If gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value = "06002" Then 'Servicios
            gDetalle.Columns.ColumnByFieldName("GlsProducto").DisableEditor = False
            gDetalle.Columns.ColumnByFieldName("VVUnit").DisableEditor = False
            gDetalle.Columns.ColumnByFieldName("PVUnit").DisableEditor = False
        Else
            gDetalle.Columns.ColumnByFieldName("GlsProducto").DisableEditor = True
        End If
    End If
    
    If gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = "P1" Then
        gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = True
        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = True
        
    ElseIf gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = "P2" Then
        gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = False
        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = False
    Else
        gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = True
        'Si la Empresa es INMAC y el documento es 87 la columna del detalle idDocVentasPres cambiamos a editable
        If rucEmpresa = "20513250445" And strTipoDoc = "87" Then
            gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = False
        Else
            gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = True
        End If
    End If
                
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gDetalle_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
Dim dblVVUnit As Double, dblIGVUnit As Double, dblPVUnit As Double
Dim intFila As Integer

    intFila = Node.Index + 1
    If gDetalle.Dataset.State = dsEdit Then
        gDetalle.Dataset.Post
    End If
    gDetalle.Dataset.RecNo = intFila
    
    Select Case Column.Index
        Case gDetalle.Columns.ColumnByFieldName("Afecto").Index
            procesaMoneda gDetalle.Columns.ColumnByFieldName("idMoneda").Value, txtCod_Moneda.Text, 0, gDetalle.Columns.ColumnByFieldName("VVUnit").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value, dblVVUnit, dblIGVUnit, dblPVUnit
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            procesaMoneda gDetalle.Columns.ColumnByFieldName("idMoneda").Value, txtCod_Moneda.Text, 0, gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value, dblVVUnit, dblIGVUnit, dblPVUnit
            gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = dblPVUnit
            
            If dblPVUnit > Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value) Then
                MsgBox "El precio ingresado es menor al precio lista.", vbInformation, App.Title
                gDetalle.Columns.ColumnByFieldName("VVUnit").Value = 0
                gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = 0
                gDetalle.Columns.ColumnByFieldName("PVUnit").Value = 0
            End If
            calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
            gDetalle.Dataset.Post
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
    End Select

End Sub

Private Sub gDetalle_OnDblClick()
    
    Frame2(0).Visible = True
    CATTextBox1(0).Visible = True
    CATTextBox1(0).SetFocus
    CATTextBox1(0).Text = gDetalle.Columns.ColumnByFieldName("glsProducto").Value

End Sub

Private Sub gdetalle_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String
Dim strCod As String
Dim strDes As String
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
Dim indPedido As Boolean
Dim strTipoDocImportado As String
Dim rscd As ADODB.Recordset
Dim CIdCentroCosto              As String
Dim CArrays()                   As String
Dim CIdSerie                    As String
Dim CIdDocVentas                As String

    intFila = Node.Index + 1
    
    Select Case Column.Index
        Case gDetalle.Columns.ColumnByFieldName("idProducto").Index, gDetalle.Columns.ColumnByFieldName("CodigoRapido").Index
            indPedido = False
            If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Or strTipoDoc = "69" Or strTipoDoc = "97" Then indPedido = True
            If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Or strTipoDoc = "69" Or strTipoDoc = "97" Then
                strCod = ""
                strDes = ""
                strCodUM = ""
                'FrmAyudaProdOC.ExecuteReturnTextAlm txtCod_Almacen.Text, rscd, strCod, strDes, strCodUM, glsValidaStock, "", False, True, indPedido, False, StrMsgError
                
                FrmAyudaProdOCInv.ExecuteReturnTextAlm txtCod_Cliente.Text, txtCod_Almacen.Text, rscd, strCod, strDes, strCodUM, glsValidaStock, txtCod_Lista.Text, True, True, indPedido, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                
                If strCod <> "" Then
                    gDetalle.SetFocus
                    gDetalle.Dataset.RecNo = intFila
                    gDetalle.Dataset.Edit
                    
                    If leeParametro("VIZUALIZA_CODIGO_RAPIDO") = "S" Then
                        gDetalle.Columns.ColumnByFieldName("idProducto").Value = traerCampo("Productos", "IdProducto", "CodigoRapido", strCod, True)
                        gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = strCod
                    Else
                        gDetalle.Columns.ColumnByFieldName("idProducto").Value = strCod
                        gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = traerCampo("Productos", "CodigoRapido", "IdProducto", strCod, True)
                    End If
                    
                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = strDes
                        
                    If Trim(gDetalle.Columns.ColumnByFieldName("idProducto").Value) = "" Then Exit Sub
                    If DatosProducto(gDetalle.Columns.ColumnByFieldName("idProducto").Value, strCodFabri, strCodMar, strDesMar, intAfecto, strTipoProd) = False Then
                    End If
                    strMoneda = "PEN"
                    gDetalle.Columns.ColumnByFieldName("idCodFabricante").Value = strCodFabri
                    gDetalle.Columns.ColumnByFieldName("idMarca").Value = strCodMar
                    gDetalle.Columns.ColumnByFieldName("GlsMarca").Value = strDesMar
                    gDetalle.Columns.ColumnByFieldName("Afecto").Value = intAfecto
                    gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value = strTipoProd
                    gDetalle.Columns.ColumnByFieldName("idMoneda").Value = strMoneda 'falta esta columna en el detalle de la grilla
                        
                    If DatosPrecio(gDetalle.Columns.ColumnByFieldName("idProducto").Value, strTipoProd, strCodUM, strDesUM, dblVVUnit, dblFactor) = False Then
                    End If
                    If strDesUM = "" And strCodUM <> "" Then strDesUM = traerCampo("unidadMedida", "abreUM", "idUM", strCodUM, False)
                    
                    gDetalle.Columns.ColumnByFieldName("idUM").Value = strCodUM
                    gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
                    gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
                    
                    If strTipoProd = "06002" Then gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 1
                    
                    procesaMoneda strMoneda, txtCod_Moneda.Text, 0, dblVVUnit, intAfecto, dblVVUnit, dblIGVUnit, dblPVUnit
                    
                    gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
                    gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
                    gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
                    gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = dblVVUnit
                    gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = dblPVUnit
                    gDetalle.Columns.ColumnByFieldName("PorDcto").Value = dblPorDsctoEspecial
                    gDetalle.Dataset.Post
                    gDetalle.Dataset.RecNo = intFila
                    gDetalle.Dataset.Edit
                    
                    calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
                                    
                    gDetalle.Dataset.Post
                    
                    If strCod <> "" Then
                        gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").Index
                    End If
                Else
                    If rscd.RecordCount <> 0 Then
                        mostrarDocImportado2 rscd, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                End If
            End If
            
        Case gDetalle.Columns.ColumnByFieldName("idUM").Index
            strCod = gDetalle.Columns.ColumnByFieldName("idUM").Value
            strDes = gDetalle.Columns.ColumnByFieldName("GlsUM").Value
            dblFactor = gDetalle.Columns.ColumnByFieldName("Factor").Value
            
            mostrarAyudaTextoPrecios gDetalle.Columns.ColumnByFieldName("idProducto").Value, txtCod_Lista.Text, strCod, strDes, dblFactor
            gDetalle.SetFocus
            
            gDetalle.Dataset.RecNo = intFila
            
            If DatosPrecio(gDetalle.Columns.ColumnByFieldName("idProducto").Value, gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value, strCod, strDes, dblVVUnit, dblFactor) = False Then
            End If
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idUM").Value = strCod
            gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDes
            gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
            intAfecto = gDetalle.Columns.ColumnByFieldName("afecto").Value
            procesaMoneda gDetalle.Columns.ColumnByFieldName("idMoneda").Value, txtCod_Moneda.Text, 0, dblVVUnit, intAfecto, dblVVUnit, dblIGVUnit, dblPVUnit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = dblPVUnit
            gDetalle.Dataset.Post
            gDetalle.Dataset.RecNo = intFila
            gDetalle.Dataset.Edit
            
            calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
            gDetalle.Dataset.Post
            
            If strCod <> "" Then
                gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").Index
            End If
        
        Case gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Index
            
            FrmAyudaHCosto.MostrarForm StrMsgError, "", "O", "", CIdCentroCosto, "", ""
            If StrMsgError <> "" Then GoTo Err
            
            If Len(Trim(CIdCentroCosto)) > 0 Then
                gDetalle.Dataset.Edit
                gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value = CIdCentroCosto
                ReDim CArrays(5)
                traerCampos2 Cn, "CentrosCosto A Left Join DocVentas B On A.IdEmpresa = B.IdEmpresa And A.IdCentroCosto = B.IdCentroCosto", "A.GlsCentroCosto,B.IdDocumento,B.IdSerie,B.IdDocVentas,B.IdSucursal", "A.IdCentroCosto", CIdCentroCosto, 5, CArrays, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                
                If Len(Trim("" & CArrays(0))) > 0 Then
                    traerCampos2 Cn, "CentrosCosto A Left Join DocVentas B On A.IdEmpresa = B.IdEmpresa And A.IdCentroCosto = B.IdCentroCosto", "A.GlsCentroCosto,B.IdDocumento,B.IdSerie,B.IdDocVentas,B.IdSucursal", "A.IdCentroCosto", CIdCentroCosto, 5, CArrays, False, "A.IdEmpresa = '" & glsEmpresa & "' And (B.IdDocumento In('P1','P2') Or B.IdDocumento Is Null)"
                    gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = CArrays(1)
                    
                    If CArrays(1) = "P1" Then
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = CArrays(2)
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = CArrays(3)
                        gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = CArrays(4)
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = True
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = True
                        
                    ElseIf CArrays(1) = "P2" Then
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = False
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = False
                    Else
                        gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = True
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = True
                    End If
                
                Else
                    gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = True
                    gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = True
                End If
                gDetalle.Dataset.Post
            End If
        
        Case gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Index
            'FrmAyudaPresupuestos.MostrarForm strMsgError, gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value, CIdSerie, CIdDocVentas, "O", ""
            'If strMsgError <> "" Then GoTo Err
            
            'If Len(Trim("" & CIdDocVentas)) > 0 Then
            '    gDetalle.Dataset.Edit
            '    gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = CIdSerie
            '    gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = CIdDocVentas
            '    gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = traerCampo("DocVentas", "IdSucursal", "IdDocumento", gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value, True, "IdCentroCosto = '" & gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value & "' And IdSerie = '" & gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value & "' And IdDocVentas = '" & gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value & "'")
            '    gDetalle.Dataset.Post
            'End If
    End Select
    
    If gDetalle.Columns.ColumnByFieldName("FechaEmision").Index <> Column.Index Then
        calcularTotales
        gDetalle.Dataset.RecNo = intFila
    End If
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gDetalle_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim rsp As New ADODB.Recordset
Dim StrMsgError As String
Dim strCod As String
Dim strDes As String
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
Dim dblDcto As Double
Dim dblTotalBruto As Double
Dim IndEvaluacion As Integer
Dim strCodUsuarioAutorizacion As String
Dim i As Integer
Dim sw_dcto     As Boolean
Dim ndcto       As Double
Dim dblDctoVer As Double
Dim strDcto As String
Dim strPorDcto() As String
Dim CArrays()                   As String
Dim CIdSerie                    As String
Dim CIdDocVentas                As String
Dim intFila                     As Integer

    If gDetalle.Dataset.Modified = False Then Exit Sub
    
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo

    Select Case gDetalle.Columns.FocusedColumn.Index
        Case gDetalle.Columns.ColumnByFieldName("idProducto").Index
            strCod = gDetalle.Columns.ColumnByFieldName("idProducto").Value
            strDes = gDetalle.Columns.ColumnByFieldName("GlsProducto").Value
            
            csql = "Select P.IdProducto,P.GlsProducto,P.idUMVenta " & _
                    "FROM Productos p " & _
                    "INNER JOIN productosproveedor x " & _
                        "ON p.IdEmpresa = x.IdEmpresa And p.idProducto = x.idProducto and x.idProveedor = '" & Trim(txtCod_Cliente.Text) & "' " & _
                   "WHERE p.idempresa = '" & glsEmpresa & "' " & _
                   "AND (p.idProducto = '" & strCod & "' OR p.idFabricante = '" & strCod & "' OR p.CodigoRapido = '" & strCod & "') " & _
                   "Group By p.IdEmpresa,P.IdProducto"
            
            rsp.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
            If rsp.EOF Or rsp.BOF Then
                StrMsgError = "No se encuentra registrado el producto"
                gDetalle.Dataset.Edit
                gDetalle.Columns.ColumnByFieldName("idProducto").Value = ""
                gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = ""
                gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
                gDetalle.Dataset.Post
                GoTo Err
            Else
                strCod = "" & rsp.Fields("idProducto")
                strDes = "" & rsp.Fields("GlsProducto")
                strCodUM = "" & rsp.Fields("idUMVenta")
            End If
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idProducto").Value = strCod
            gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = traerCampo("Productos", "CodigoRapido", "IdProducto", strCod, True)
            gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = strDes
            If DatosProducto(strCod, strCodFabri, strCodMar, strDesMar, intAfecto, strTipoProd) = False Then
            End If
                
            strMoneda = traerCampo("Listaprecios", "idMoneda", "idLista", txtCod_Lista.Text, True)
            gDetalle.Columns.ColumnByFieldName("idCodFabricante").Value = strCodFabri
            gDetalle.Columns.ColumnByFieldName("idMarca").Value = strCodMar
            gDetalle.Columns.ColumnByFieldName("GlsMarca").Value = strDesMar
            gDetalle.Columns.ColumnByFieldName("Afecto").Value = intAfecto
            gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value = strTipoProd
            gDetalle.Columns.ColumnByFieldName("idMoneda").Value = strMoneda 'falta esta columna en el detalle de la grilla
            If DatosPrecio(strCod, strTipoProd, strCodUM, strDesUM, dblVVUnit, dblFactor) = False Then
            End If
            gDetalle.Columns.ColumnByFieldName("idUM").Value = strCodUM
            gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
            gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
            
            If strTipoProd = "06002" Then gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 1
            procesaMoneda strMoneda, txtCod_Moneda.Text, 0, dblVVUnit, intAfecto, dblVVUnit, dblIGVUnit, dblPVUnit
            
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = dblPVUnit
            gDetalle.Columns.ColumnByFieldName("PorDcto").Value = dblPorDsctoEspecial
            gDetalle.Dataset.Post
            gDetalle.Dataset.Edit
            calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
            gDetalle.Dataset.Post
                    
            If strCod <> "" Then
                gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").ColIndex '.Index
            End If
            
            If dblVVUnit = 0# Then
                MsgBox "El producto NO registra precio. Verifique.", vbCritical, App.Title
            End If
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("VVUnit").Index
            procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 0, Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), gDetalle.Columns.ColumnByFieldName("Afecto").Value, dblVVUnit, dblIGVUnit, dblPVUnit
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
            gDetalle.Dataset.Post
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("PVUnit").Index
            procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 1, gDetalle.Columns.ColumnByFieldName("PVUnit").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value, dblVVUnit, dblIGVUnit, dblPVUnit
            If Val("" & gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value) <= dblPVUnit Then
                gDetalle.Dataset.Edit
                gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
                gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
                gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
                calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
                gDetalle.Dataset.Post
                calcularTotales
            Else
                MsgBox "El precio ingresado es menor al precio lista.", vbInformation, App.Title
                gDetalle.Dataset.Edit
                gDetalle.Columns.ColumnByFieldName("VVUnit").Value = gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value
                gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value - gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value
                gDetalle.Columns.ColumnByFieldName("PVUnit").Value = gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value
                gDetalle.Dataset.Post
            End If
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("Cantidad").Index
            gDetalle.Dataset.Edit
            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("IGVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value), gDetalle.Columns.ColumnByFieldName("PorDcto").Value, Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
            gDetalle.Dataset.Post
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("PorDcto").Index
            IndEvaluacion = 0
            If ("" & gDetalle.Columns.ColumnByFieldName("PorDcto").Value) <> "0" Then
                strDcto = "" & gDetalle.Columns.ColumnByFieldName("PorDcto").Value
                If Trim(strDcto) <> "" And strDcto <> "0" Then
                    strPorDcto = Split(strDcto, "+")
                    For i = 0 To UBound(strPorDcto)
                        dblDctoVer = dblDctoVer + (Val("" & strPorDcto(i)))
                    Next
                Else
                    dblDctoVer = 0
                End If
            End If
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idUsuarioDcto").Value = strCodUsuarioAutorizacion
            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("IGVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value), ("" & gDetalle.Columns.ColumnByFieldName("PorDcto").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
            gDetalle.Dataset.Post
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("DctoVV").Index
            IndEvaluacion = 0
            If Val("" & gDetalle.Columns.ColumnByFieldName("DctoVV").Value) > 0 Then
                frmAprobacion.MostrarForm "02", IndEvaluacion, strCodUsuarioAutorizacion, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                If IndEvaluacion = 0 Then
                    strCodUsuarioAutorizacion = ""
                    gDetalle.Columns.ColumnByFieldName("DctoVV").Value = 0
                End If
            End If
            dblDcto = Val("" & gDetalle.Columns.ColumnByFieldName("DctoVV").Value)
            dblTotalBruto = Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value) 'DCTO POR PRECIO UNITARIO
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("idUsuarioDcto").Value = strCodUsuarioAutorizacion
            gDetalle.Columns.ColumnByFieldName("PorDcto").Value = (dblDcto * 100) / dblTotalBruto
            gDetalle.Dataset.Post
            gDetalle.Dataset.Edit
            calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("IGVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value), ("" & gDetalle.Columns.ColumnByFieldName("PorDcto").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
            gDetalle.Dataset.Post
            calcularTotales
            gDetalle.Dataset.RecNo = intFila
        
        Case gDetalle.Columns.ColumnByFieldName("DctoPV").Index
            IndEvaluacion = 0
            If Val("" & gDetalle.Columns.ColumnByFieldName("DctoPV").Value) > 0 Then
                frmAprobacion.MostrarForm "02", IndEvaluacion, strCodUsuarioAutorizacion, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                
                If IndEvaluacion = 0 Then
                    strCodUsuarioAutorizacion = ""
                    gDetalle.Columns.ColumnByFieldName("DctoPV").Value = 0
                End If
            End If
            dblDcto = Val("" & gDetalle.Columns.ColumnByFieldName("DctoPV").Value)
            dblTotalBruto = Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value)  'DCTO POR PRECIO UNITARIO
            
            If glsDctoMinMonto > 0 Then
                sw_dcto = True
                ndcto = dblTotalBruto * (glsDctoMinMonto / 100)
                If dblDcto > ndcto Then
                    sw_dcto = False
                    MsgBox "El monto del descuento es mayor al " & Format$(glsDctoMinMonto, "00") & "%", vbCritical, App.Title
                    gDetalle.Dataset.Edit
                    gDetalle.Columns.ColumnByFieldName("DctoPV").Value = 0
                    gDetalle.Dataset.Post
                End If
            Else
                sw_dcto = True
            End If
            
            If sw_dcto = True Then
                gDetalle.Dataset.Edit
                gDetalle.Columns.ColumnByFieldName("idUsuarioDcto").Value = strCodUsuarioAutorizacion
                gDetalle.Columns.ColumnByFieldName("PorDcto").Value = (dblDcto * 100) / dblTotalBruto
                gDetalle.Dataset.Post
                gDetalle.Dataset.Edit
                calculaTotalesFila Val("" & gDetalle.Columns.ColumnByFieldName("Cantidad").Value), Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("IGVUnit").Value), Val("" & gDetalle.Columns.ColumnByFieldName("PVUnit").Value), ("" & gDetalle.Columns.ColumnByFieldName("PorDcto").Value), Val("" & gDetalle.Columns.ColumnByFieldName("Afecto").Value)
                gDetalle.Dataset.Post
                calcularTotales
            End If
            gDetalle.Dataset.RecNo = intFila
            
        Case gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Index
            If Len(Trim("" & gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value)) > 0 Then
                gDetalle.Dataset.Edit
                ReDim CArrays(5)
                traerCampos2 Cn, "CentrosCosto A Left Join DocVentas B On A.IdEmpresa = B.IdEmpresa And A.IdCentroCosto = B.IdCentroCosto", "A.GlsCentroCosto,B.IdDocumento,B.IdSerie,B.IdDocVentas,B.IdSucursal", "A.IdCentroCosto", gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value, 5, CArrays, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                
                If Len(Trim("" & CArrays(0))) > 0 Then
                    traerCampos2 Cn, "CentrosCosto A Left Join DocVentas B On A.IdEmpresa = B.IdEmpresa And A.IdCentroCosto = B.IdCentroCosto", "A.GlsCentroCosto,B.IdDocumento,B.IdSerie,B.IdDocVentas,B.IdSucursal", "A.IdCentroCosto", gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value, 5, CArrays, False, "A.IdEmpresa = '" & glsEmpresa & "' And (B.IdDocumento In('P1','P2') Or B.IdDocumento Is Null)"
                    gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = CArrays(1)
                    
                    If CArrays(1) = "P1" Then
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = CArrays(2)
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = CArrays(3)
                        gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = CArrays(4)
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = True
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = True
                        
                    ElseIf CArrays(1) = "P2" Then
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = False
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = False
                    Else
                        gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
                        gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = True
                        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = True
                    End If
                Else
                    gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = True
                    gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = True
                End If
                gDetalle.Dataset.Post
            
            Else
                gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value = ""
                gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = ""
                gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
                gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
                gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
                gDetalle.Columns.ColumnByFieldName("IdSeriePres").DisableEditor = True
                gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").DisableEditor = True
                gDetalle.Dataset.Post
            End If
            
        Case gDetalle.Columns.ColumnByFieldName("IdSeriePres").Index, gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Index
            If Len(Trim("" & gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value)) > 0 And Len(Trim("" & gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value)) > 0 Then
                If Len(Trim(traerCampo("DocVentas", "IdSucursal", "IdDocumento", gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value, True, "IdCentroCosto = '" & gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value & "' And IdSerie = '" & gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value & "' And IdDocVentas = '" & gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value & "' And IndAprobado = '1' And (IndTerminado = '0' Or Length(Trim(IfNull(IndTerminado,''))) = 0)"))) = 0 Then
                    gDetalle.Dataset.Edit
                    gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
                    gDetalle.Dataset.Post
                    StrMsgError = "Ingrese Serie y Número de Presupuesto correcto": GoTo Err
                End If
            End If
            
        Case gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Index
            procesaMoneda txtCod_Moneda.Text, txtCod_Moneda.Text, 0, Val("" & gDetalle.Columns.ColumnByFieldName("VVUnit").Value), gDetalle.Columns.ColumnByFieldName("Afecto").Value, dblVVUnit, dblIGVUnit, dblPVUnit
            gDetalle.Dataset.Edit
            gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
            gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
            gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
            
            calculaTotalesFilaPVNeto gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
            gDetalle.Dataset.Post
            calcularTotales
            gDetalle.SetFocus
            gDetalle.Dataset.RecNo = intFila
    End Select
    If rsp.State = 1 Then rsp.Close: Set rsp = Nothing
    
    Exit Sub

Err:
    If rsp.State = 1 Then rsp.Close: Set rsp = Nothing
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
                    gDetalle.Columns.ColumnByFieldName("idCodFabricante").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idMarca").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsMarca").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idUM").Value = ""
                    gDetalle.Columns.ColumnByFieldName("GlsUM").Value = ""
                    gDetalle.Columns.ColumnByFieldName("Factor").Value = 1
                    gDetalle.Columns.ColumnByFieldName("Afecto").Value = 1
                    gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 0
                    gDetalle.Columns.ColumnByFieldName("VVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("PVUnit").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalVVBruto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalPVBruto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("PorDcto").Value = "0"
                    gDetalle.Columns.ColumnByFieldName("DctoVV").Value = 0
                    gDetalle.Columns.ColumnByFieldName("DctoPV").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("idMoneda").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idDocumentoImp").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idDocVentasImp").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idSerieImp").Value = ""
                    gDetalle.Columns.ColumnByFieldName("NumLote").Value = ""
                    gDetalle.Columns.ColumnByFieldName("FecVencProd").Value = ""
                    gDetalle.Columns.ColumnByFieldName("idUsuarioDcto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = 0
                    gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = 0
                    gDetalle.Columns.ColumnByFieldName("VVUnitNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("PVUnitNeto").Value = 0
                    gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
                    gDetalle.Columns.ColumnByFieldName("FechaEmision").Value = getFechaSistema
                    gDetalle.Columns.ColumnByFieldName("GlsPlaca").Value = ""
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
                calcularTotales
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
Dim strTipoDocImportado As String
Dim rscd As New ADODB.Recordset
Dim indPedido As Boolean

    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    intFila = gDetalle.Dataset.RecNo
    
    If Key <> 9 And Key <> 13 And Key <> 27 Then
        Select Case gDetalle.Columns.FocusedColumn.Index
            Case gDetalle.Columns.ColumnByFieldName("idProducto").Index, gDetalle.Columns.ColumnByFieldName("CodigoRapido").Index
                If glsLeeCodigoBarras = "N" Then
                    strCod = gDetalle.Columns.ColumnByFieldName("idProducto").Value
                    strDes = gDetalle.Columns.ColumnByFieldName("GlsProducto").Value
                    indPedido = False
                    If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Then indPedido = True
                    If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Then
                        FrmAyudaProdOCInv.ExecuteReturnTextAlm txtCod_Cliente.Text, txtCod_Almacen.Text, rscd, strCod, strDes, strCodUM, glsValidaStock, txtCod_Lista.Text, True, True, indPedido, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                        If rscd.RecordCount <> 0 Then
                            mostrarDocImportado2 rscd, StrMsgError
                            If StrMsgError <> "" Then GoTo Err
                        End If
                    Else
                        mostrarAyudaTextoProdAlm2 txtCod_Almacen.Text, strCod, strDes, strCodUM, glsValidaStock, txtCod_Lista.Text, True, True, indPedido, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                    gDetalle.SetFocus
                    gDetalle.Dataset.RecNo = intFila
                    gDetalle.Dataset.Edit
                    gDetalle.Columns.ColumnByFieldName("idProducto").Value = strCod
                    gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = traerCampo("Productos", "CodigoRapido", "IdProducto", strCod, True)
                    gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = strDes
                        
                    If Trim(strCod) = "" Then Exit Sub
                    If DatosProducto(strCod, strCodFabri, strCodMar, strDesMar, intAfecto, strTipoProd) = False Then
                    End If
                    strMoneda = traerCampo("Listaprecios", "idMoneda", "idLista", txtCod_Lista.Text, True)
                    gDetalle.Columns.ColumnByFieldName("idCodFabricante").Value = strCodFabri
                    gDetalle.Columns.ColumnByFieldName("idMarca").Value = strCodMar
                    gDetalle.Columns.ColumnByFieldName("GlsMarca").Value = strDesMar
                    gDetalle.Columns.ColumnByFieldName("Afecto").Value = intAfecto
                    gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value = strTipoProd
                    gDetalle.Columns.ColumnByFieldName("idMoneda").Value = strMoneda 'falta esta columna en el detalle de la grilla
                    If DatosPrecio(strCod, strTipoProd, strCodUM, strDesUM, dblVVUnit, dblFactor) = False Then
                    End If
                    
                    If strDesUM = "" And strCodUM <> "" Then strDesUM = traerCampo("unidadMedida", "abreUM", "idUM", strCodUM, False)
                    gDetalle.Columns.ColumnByFieldName("idUM").Value = strCodUM
                    gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
                    gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
                    If strTipoProd = "06002" Then gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 1
                    procesaMoneda strMoneda, txtCod_Moneda.Text, 0, dblVVUnit, intAfecto, dblVVUnit, dblIGVUnit, dblPVUnit
                    gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
                    gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
                    gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
                    gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = dblVVUnit
                    gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = dblPVUnit
                    gDetalle.Columns.ColumnByFieldName("PorDcto").Value = dblPorDsctoEspecial
                    gDetalle.Dataset.Post
            
                    gDetalle.Dataset.RecNo = intFila
                    gDetalle.Dataset.Edit
                    calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
                    gDetalle.Dataset.Post
                    
                    If strCod <> "" Then
                        gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").Index
                    End If
                End If
            
            Case gDetalle.Columns.ColumnByFieldName("idUM").Index
                strCod = gDetalle.Columns.ColumnByFieldName("idUM").Value
                strDes = gDetalle.Columns.ColumnByFieldName("GlsUM").Value
                mostrarAyudaKeyasciiTextoPrecios Key, gDetalle.Columns.ColumnByFieldName("idProducto").Value, txtCod_Lista.Text, strCod, strDes
                Key = 0
                gDetalle.SetFocus
                gDetalle.Dataset.RecNo = intFila
                If DatosPrecio(gDetalle.Columns.ColumnByFieldName("idProducto").Value, gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value, strCod, strDes, dblVVUnit, dblFactor) = False Then
                End If
                gDetalle.Dataset.Edit
                gDetalle.Columns.ColumnByFieldName("idUM").Value = strCod
                gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDes
                intAfecto = gDetalle.Columns.ColumnByFieldName("Afecto").Value
                procesaMoneda gDetalle.Columns.ColumnByFieldName("idMoneda").Value, txtCod_Moneda.Text, 0, dblVVUnit, intAfecto, dblVVUnit, dblIGVUnit, dblPVUnit
                gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
                gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
                gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
                gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = dblVVUnit
                gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = dblPVUnit
                gDetalle.Dataset.Post
                gDetalle.Dataset.RecNo = intFila
                gDetalle.Dataset.Edit
                calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
                gDetalle.Dataset.Post
                
                If strCod <> "" Then
                    gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").Index
                End If
                calcularTotales
                gDetalle.Dataset.RecNo = intFila
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
            gDocReferencia.Columns.FocusedIndex = gDocReferencia.Columns.ColumnByFieldName("idDocumento").Index
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
            mostrarAyudaTexto "DOCUMENTOS", strCod, strDes
            gDocReferencia.Dataset.Edit
            gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value = strCod
            gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value = strDes
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
            gDocReferencia.Columns.ColumnByFieldName("idSerie").Value = Format(gDocReferencia.Columns.ColumnByFieldName("idSerie").Value, "000")
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
                    gDocReferencia.Columns.ColumnByFieldName("idDocumento").Value = ""
                    gDocReferencia.Columns.ColumnByFieldName("GlsDocumento").Value = ""
                    gDocReferencia.Columns.ColumnByFieldName("idSerie").Value = ""
                    gDocReferencia.Columns.ColumnByFieldName("idNumDOc").Value = ""
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
                
                mostrarAyudaKeyasciiTexto Key, "DOCUMENTOS", strCod, strDes
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
Dim strestDocImport As String

    mostrarDocVentas gLista.Columns.ColumnByName("idDocVentas").Value, gLista.Columns.ColumnByName("idSerie").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraDetalle.Visible = True
    fraTotales.Visible = True
    fraGeneral.Enabled = False
    fraDetalle.Enabled = False
    intBoton = 3
    habilitaBotones 2
    
    ValidaAprobacion StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Then
        strestDocImport = "" & traerCampo("docventas", "estdocimportado", "idDocumento", strTipoDoc, True, "iddocventas = '" & txt_NumDoc.Text & "' and idSerie = '" & txt_Serie.Text & "' And  idsucursal ='" & glsSucursal & "'")
             If strestDocImport = "S" Then
                Toolbar1.Buttons(3).Visible = False 'MODIFICAR
                Toolbar1.Buttons(5).Visible = False 'ELIMINAR
             End If
    End If
    
    If strTipoDoc = "87" Then
        If Len(Trim("" & traerCampo("DocReferencia", "TipoDocReferencia", "TipoDocOrigen", "87", True, "TipoDocReferencia In('P1','P2') And SerieDocOrigen = '" & txt_Serie.Text & "' And NumDocOrigen = '" & txt_NumDoc.Text & "'"))) > 0 Then
            Toolbar1.Buttons(3).Visible = False 'MODIFICAR
            Toolbar1.Buttons(5).Visible = False 'ELIMINAR
            Toolbar1.Buttons(6).Visible = False 'ANULAR
        End If
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim rscd                As ADODB.Recordset
Dim rsdd                As ADODB.Recordset
Dim strPeriodo          As String
Dim StrMsgError         As String
Dim rucEmp              As String
Dim strTipoDocImportado As String
Dim strAno              As String
Dim strParamImpOC       As String

    Select Case Button.Index
        Case 1 'Nuevo
            intBoton = Button.Index
            nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            habilitaBotones Button.Index
            valoresIniciales
            
            txtCod_TipoTicket.Text = "08002"
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraDetalle.Visible = True
            fraTotales.Visible = True
            
            If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Then
                lbl_Cliente.Caption = "Proveedor"
            End If
            strPeriodo = Format(dtp_Emision.Value, "yyyymm")
            
            If Trim("" & traerCampo("parametros", "valparametro", "glsparametro", "PERIODO_CAMBIO_IGV", True)) > strPeriodo Then
                dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV_ANT", True), "0.00")) / 100, "0.00")
            Else
                dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV", True), "0.00")) / 100, "0.00")
            End If
            
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            intBoton = 3
            
            ValidaAprobacion StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        Case 3 'Modificar
            If glsModVendCampo = False Then
                txtCod_VendedorCampo.Locked = True
            Else
                txtCod_VendedorCampo.Locked = False
            End If
            fraGeneral.Enabled = True
            fraDetalle.Enabled = True
            habilitaBotones Button.Index
            strPeriodo = Format(dtp_Emision.Value, "yyyymm")
            If Trim("" & traerCampo("parametros", "valparametro", "glsparametro", "PERIODO_CAMBIO_IGV", True)) > strPeriodo Then
                dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV_ANT", True), "0.00")) / 100, "0.00")
            Else
                dblIgvNEw = Format(Val(Format(traerCampo("parametros", "valparametro", "glsparametro", "IGV", True), "0.00")) / 100, "0.00")
            End If
            
        Case 4 'Cancelar
            fraGeneral.Enabled = False
            fraDetalle.Enabled = False
            habilitaBotones Button.Index
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 6 'Anular
            anularDoc StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 7 'Imprimir
            If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Or strTipoDoc = "97" Then
                strParamImpOC = Trim("" & traerCampo("parametros", "Valparametro", "glsparametro", "FORMATO_OC", True))
                If strParamImpOC = "2" Then
                    imprimeDocVentas strTipoDoc, txt_NumDoc.Text, txt_Serie.Text, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    csql = "UPDATE docventas SET estDocventas = 'IMP' where idDocumento = '" & strTipoDoc & "' and idDocVentas = '" & Trim(txt_NumDoc.Text) & "' and idSerie = '" & Trim(txt_Serie.Text) & "' And idEmpresa ='" & glsEmpresa & "'"
                    Cn.Execute csql
                    strEstDocVentas = "IMP"
                    
                    listaDocVentas StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    
                Else
                    imprimeOCompra StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                End If
            End If
            habilitaBotones Button.Index
        
        Case 8 'Lista
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraDetalle.Visible = False
            fraTotales.Visible = False
            strAno = Year(Now)
            txt_Ano.Text = strAno
            habilitaBotones Button.Index
        
        Case 9 'Excel
            gLista.m.ExportToXLS App.Path & "\Temporales\Listado.xls"
            ShellEx App.Path & "\Temporales\Listado.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        
        Case 10 'Importar
            frmListaDocExportar.MostrarForm strTipoDoc, txtCod_Cliente.Text, rscd, rsdd, strTipoDocImportado, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
            If strTipoDocImportado <> "" Then
                If leeParametro("AGRUPAPRODUCTOS") = "N" Then
                    If leeParametro("RECUPERA_PRECIO_REQUERIMIENTO_COMPRA") = "S" Then
                        MostrarDocImportadoSinAgrupar2 rscd, rsdd, strTipoDocImportado, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    Else
                        MostrarDocImportadoSinAgrupar rscd, rsdd, strTipoDocImportado, StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                Else
                    mostrarDocImportado rscd, rsdd, strTipoDocImportado, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                End If
            End If
            Unload frmListaDocExportar
            
        Case 11 'Enviar Correo OC
            Enviar_Correo StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
        Case 12 'Salir
            Unload Me
            
    End Select
    
    If TypeName(rscd) <> "Nothing" Then
        If rscd.State = 1 Then rscd.Close: Set rscd = Nothing
        If rsdd.State = 1 Then rsdd.Close: Set rsdd = Nothing
    End If
    
    Exit Sub
    
Err:
    If TypeName(rscd) <> "Nothing" Then
        If rscd.State = 1 Then rscd.Close: Set rscd = Nothing
        If rsdd.State = 1 Then rsdd.Close: Set rsdd = Nothing
    End If
  
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub habilitaBotones(indexBoton)
On Error GoTo Err
Dim StrMsgError As String

    Select Case indexBoton
        Case 1:
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(10).Visible = True
            Toolbar1.Buttons(11).Visible = False
            Toolbar1.Buttons(12).Visible = True
            
        Case 2: 'Grabar
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = True
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = True
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(10).Visible = False
            Toolbar1.Buttons(11).Visible = True
            Toolbar1.Buttons(12).Visible = True
            
            ocultarColumnasEstado
        
        Case 3 'Modificar
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = True
            Toolbar1.Buttons(5).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = True
            Toolbar1.Buttons(8).Visible = False
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(10).Visible = True
            Toolbar1.Buttons(11).Visible = True
            Toolbar1.Buttons(12).Visible = True
            ocultarColumnasEstado
        
        Case 4 'Cancelar
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = True
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = True
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(10).Visible = False
            Toolbar1.Buttons(11).Visible = False
            Toolbar1.Buttons(12).Visible = True
            ocultarColumnasEstado
            
        Case 5 'Eliminar
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(10).Visible = False
            Toolbar1.Buttons(11).Visible = False
            Toolbar1.Buttons(12).Visible = True
            ocultarColumnasEstado
        
        Case 6 'Anular
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            If glsGrabaTodo = "S" Then
                Toolbar1.Buttons(13).Visible = False
            End If
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(10).Visible = False
            Toolbar1.Buttons(11).Visible = False
            Toolbar1.Buttons(12).Visible = True
            ocultarColumnasEstado
        
        Case 7 'Imprimir
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = True
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = True
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(10).Visible = False
            Toolbar1.Buttons(11).Visible = True
            Toolbar1.Buttons(12).Visible = True
        
        Case 8 'Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
            Toolbar1.Buttons(8).Visible = False
            Toolbar1.Buttons(9).Visible = True
            Toolbar1.Buttons(10).Visible = False
            Toolbar1.Buttons(11).Visible = False
            Toolbar1.Buttons(12).Visible = True
        
        Case 10, 11
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = True
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = True
            Toolbar1.Buttons(8).Visible = True
            Toolbar1.Buttons(9).Visible = False
            Toolbar1.Buttons(10).Visible = True
            Toolbar1.Buttons(11).Visible = True
            Toolbar1.Buttons(12).Visible = True
            ocultarColumnasEstado
    End Select
    
    ValidaAprobacion StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_Ano_Change()
On Error GoTo Err
Dim StrMsgError As String

    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_NumDoc_Change()

    If Len(Trim(txt_NumDoc.Text)) = 8 And Len(Trim(txt_Serie.Text)) = 3 Then
        Me.Caption = strGlsTipoDoc & " (" & txt_Serie.Text & " - " & Format(txt_NumDoc.Text, "00000000") & ")"
    Else
        Me.Caption = strGlsTipoDoc
    End If

End Sub

Private Sub txt_numdoc_LostFocus()
    
    txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")

End Sub

Private Sub txt_ruc_KeyPress(KeyAscii As Integer)
Dim idCliente As String
    
    If KeyAscii = 13 Then
        If txt_RUC.Text = "" Then
            txtCod_Cliente.Text = ""
        Else
            idCliente = traerCampo("personas", "idPersona", "ruc", txt_RUC.Text, False)
            txtCod_Cliente.Text = idCliente
        End If
    End If

End Sub

Private Sub txt_Serie_Change()

    If Len(Trim(txt_NumDoc.Text)) = 8 And Len(Trim(txt_Serie.Text)) = 3 Then
        Me.Caption = strGlsTipoDoc & " (" & txt_Serie.Text & " - " & Format(txt_NumDoc.Text, "00000000") & ")"
    Else
        Me.Caption = strGlsTipoDoc
    End If

End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaDocVentas(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond     As String
Dim strMes      As String
Dim RsListaCab As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = "%" & strCond & "%"
    End If
    
    If strTipoDoc = "94" Then
        If (cbx_Mes.ListIndex + 1) = "13" Then  '--- TODO EL AÑO
            strMes = ""
            gLista.Columns.ColumnByFieldName("Mes").Visible = True
        Else
            strMes = cbx_Mes.ListIndex + 1
            gLista.Columns.ColumnByFieldName("Mes").Visible = False
        End If
    Else
        strMes = cbx_Mes.ListIndex + 1
    End If
    
    csql = "EXEC spu_ListaDocventasOrdendeCompra '" & glsEmpresa & "','" & glsSucursal & "','" & strTipoDoc & "'," & Val(txt_Ano.Text) & "," & Val(strMes) & ",'" & strCond & "'"
    
    If RsListaCab.State = 1 Then RsListaCab.Close: Set RsListaCab = Nothing
    RsListaCab.Open csql, Cn, adOpenStatic, adLockOptimistic
        
    Set gLista.DataSource = RsListaCab
    
'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = "CALL spu_ListaDocventasOrdendeCompra ('" & glsEmpresa & "','" & glsSucursal & "','" & strTipoDoc & "'," & Val(txt_Ano.Text) & ",'" & strMes & "' ,'" & strCond & "') "
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "item"
'    End With
    
    ListaDetalle
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarDocVentas(strNum As String, strSerie As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim RsD As New ADODB.Recordset

    indCargando = True
    csql = "SELECT * " & _
           "FROM docventas d " & _
           "WHERE d.idEmpresa = '" & glsEmpresa & "' AND  d.idSucursal = '" & glsSucursal & "'  AND d.idDocumento = '" & strTipoDoc & "' AND d.idDocVentas = '" & strNum & "' AND d.idSerie = '" & strSerie & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    lblDoc.Caption = traerCampo("documentos", "GlsDocumento", "idDocumento", strTipoDoc, False)
    lblDoc.ForeColor = &H0&
    txt_Serie.Enabled = False
    txt_NumDoc.Enabled = False
    
    If Not rst.EOF Then
        strEstDocVentas = Trim("" & rst.Fields("estDocventas"))
        If strEstDocVentas = "ANU" Then
            lblDoc.ForeColor = &HFF&
            lblDoc.Caption = lblDoc.Caption & " - ANULADA"
            fraGeneral.Enabled = False
            fraDetalle.Enabled = False
        ElseIf strEstDocVentas = "IMP" Then
            fraGeneral.Enabled = False
            fraDetalle.Enabled = False
        Else
            txt_Serie.Enabled = False
            txt_NumDoc.Enabled = False
            fraGeneral.Enabled = True
            fraDetalle.Enabled = True
        End If
        indGeneraVale = IIf(("" & rst.Fields("indVale")) = "S", True, False)
    End If
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    csql = "SELECT idDocumento, idDocVentas, idSerie, idProducto, glsProducto, idMarca, idUM, Factor, Afecto, Cantidad, VVUnit, IGVUnit, PVUnit, TotalVVBruto, TotalPVBruto, PorDcto, DctoVV, DctoPV, TotalVVNeto, TotalIGVNeto, TotalPVNeto, item, GlsMarca, GlsUM, idTipoProducto, idMoneda, idCodFabricante, idEmpresa, idSucursal, estDocImportado, idDocumentoImp, idDocVentasImp, idSerieImp, NumLote, FecVencProd, idUsuarioDcto, VVUnitLista, PVUnitLista, CantidadImp, VVUnitNeto, PVUnitNeto, Cantidad2, CodigoRapido, idTallaPeso, CantidadAnt, Simbolo1, Simbolo2, Simbolo3, itemPro, PorcUtilidad, IdCentroCosto, IdSucursalPres, IdDocumentoPres, IdSeriePres, IdDocVentasPres, GlsPlaca, IdUPCliente, IsNull(FechaEmision,GETDATE()) FechaEmision " & _
           "FROM docventasdet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND  idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTipoDoc & "' AND idDocVentas = '" & strNum & "' AND idSerie = '" & strSerie & "' ORDER BY ITEM"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idProducto", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 40, adFldIsNullable
    rsg.Fields.Append "idCodFabricante", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idMarca", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsMarca", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVBruto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVBruto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PorDcto", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "DctoVV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "DctoPV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idTipoProducto", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idDocumentoImp", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "idDocVentasImp", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerieImp", adVarChar, 4, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idUsuarioDcto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "VVUnitLista", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnitLista", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnitNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnitNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IdCentroCosto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdSucursalPres", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdDocumentoPres", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "IdSeriePres", adVarChar, 3, adFldIsNullable
    rsg.Fields.Append "IdDocVentasPres", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "FechaEmision", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "GlsPlaca", adVarChar, 50, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("CodigoRapido") = ""
        rsg.Fields("idCodFabricante") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("idMarca") = ""
        rsg.Fields("GlsMarca") = ""
        rsg.Fields("idUM") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("Factor") = 1
        rsg.Fields("Afecto") = 1
        rsg.Fields("Cantidad") = 0
        rsg.Fields("VVUnit") = 0
        rsg.Fields("IGVUnit") = 0
        rsg.Fields("PVUnit") = 0
        rsg.Fields("TotalVVBruto") = 0
        rsg.Fields("TotalPVBruto") = 0
        rsg.Fields("PorDcto") = "0"
        rsg.Fields("DctoVV") = 0
        rsg.Fields("DctoPV") = 0
        rsg.Fields("TotalVVNeto") = 0
        rsg.Fields("TotalIGVNeto") = 0
        rsg.Fields("TotalPVNeto") = 0
        rsg.Fields("idTipoProducto") = ""
        rsg.Fields("idMoneda") = ""
        rsg.Fields("VVUnitLista") = 0
        rsg.Fields("PVUnitLista") = 0
        rsg.Fields("VVUnitNeto") = 0
        rsg.Fields("PVUnitNeto") = 0
        rsg.Fields("IdCentroCosto") = ""
        rsg.Fields("IdSucursalPres") = ""
        rsg.Fields("IdDocumentoPres") = ""
        rsg.Fields("IdSeriePres") = ""
        rsg.Fields("IdDocVentasPres") = ""
        rsg.Fields("FechaEmision") = getFechaSistema
        rsg.Fields("GlsPlaca") = ""
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idProducto") = "" & rst.Fields("idProducto")
            rsg.Fields("CodigoRapido") = traerCampo("Productos", "CodigoRapido", "IdProducto", "" & rst.Fields("IdProducto"), True)
            rsg.Fields("idCodFabricante") = "" & rst.Fields("idCodFabricante")
            rsg.Fields("GlsProducto") = "" & rst.Fields("GlsProducto")
            rsg.Fields("idMarca") = "" & rst.Fields("idMarca")
            rsg.Fields("GlsMarca") = "" & rst.Fields("GlsMarca")
            rsg.Fields("idUM") = "" & rst.Fields("idUM")
            rsg.Fields("GlsUM") = "" & rst.Fields("GlsUM")
            rsg.Fields("Factor") = "" & rst.Fields("Factor")
            rsg.Fields("Afecto") = "" & rst.Fields("Afecto")
            rsg.Fields("Cantidad") = "" & rst.Fields("Cantidad")
            rsg.Fields("VVUnit") = "" & rst.Fields("VVUnit")
            rsg.Fields("IGVUnit") = "" & rst.Fields("IGVUnit")
            rsg.Fields("PVUnit") = "" & rst.Fields("PVUnit")
            rsg.Fields("TotalVVBruto") = "" & rst.Fields("TotalVVBruto")
            rsg.Fields("TotalPVBruto") = "" & rst.Fields("TotalPVBruto")
            rsg.Fields("PorDcto") = "" & rst.Fields("PorDcto")
            rsg.Fields("DctoVV") = "" & rst.Fields("DctoVV")
            rsg.Fields("DctoPV") = "" & rst.Fields("DctoPV")
            rsg.Fields("TotalVVNeto") = "" & rst.Fields("TotalVVNeto")
            rsg.Fields("TotalIGVNeto") = "" & rst.Fields("TotalIGVNeto")
            rsg.Fields("TotalPVNeto") = "" & rst.Fields("TotalPVNeto")
            rsg.Fields("idTipoProducto") = "" & rst.Fields("idTipoProducto")
            rsg.Fields("idMoneda") = "" & rst.Fields("idMoneda")
            rsg.Fields("idDocumentoImp") = "" & rst.Fields("idDocumentoImp")
            rsg.Fields("idDocVentasImp") = "" & rst.Fields("idDocVentasImp")
            rsg.Fields("idSerieImp") = "" & rst.Fields("idSerieImp")
            rsg.Fields("NumLote") = "" & rst.Fields("NumLote")
            rsg.Fields("FecVencProd") = "" & rst.Fields("FecVencProd")
            rsg.Fields("idUsuarioDcto") = "" & rst.Fields("idUsuarioDcto")
            rsg.Fields("VVUnitLista") = "" & rst.Fields("VVUnitLista")
            rsg.Fields("PVUnitLista") = "" & rst.Fields("PVUnitLista")
            rsg.Fields("VVUnitNeto") = "" & rst.Fields("VVUnitNeto")
            rsg.Fields("PVUnitNeto") = "" & rst.Fields("PVUnitNeto")
            rsg.Fields("IdCentroCosto") = "" & rst.Fields("IdCentroCosto")
            rsg.Fields("IdSucursalPres") = "" & rst.Fields("IdSucursalPres")
            rsg.Fields("IdDocumentoPres") = "" & rst.Fields("IdDocumentoPres")
            rsg.Fields("IdSeriePres") = "" & rst.Fields("IdSeriePres")
            rsg.Fields("IdDocVentasPres") = "" & rst.Fields("IdDocVentasPres")
            rsg.Fields("GlsPlaca") = "" & rst.Fields("GlsPlaca")
            
            If Len(Trim("" & rst.Fields("FechaEmision"))) = 0 Then
                rsg.Fields("FechaEmision") = getFechaSistema
            Else
                rsg.Fields("FechaEmision") = Format("" & rst.Fields("FechaEmision"), "dd/mm/yyyy")
            End If
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    csql = "SELECT r.item, r.tipoDocReferencia idDocumento,d.GlsDocumento, r.numDocReferencia idNumDoc, r.serieDocReferencia idSerie " & _
            "FROM docreferencia r , documentos d " & _
            "WHERE r.idEmpresa = '" & glsEmpresa & "' AND r.idSucursal = '" & glsSucursal & "' AND r.tipoDocReferencia = d.idDocumento AND tipoDocOrigen = '" & strTipoDoc & "' AND numDocOrigen = '" & strNum & "' AND serieDocOrigen = '" & strSerie & "' ORDER BY ITEM"
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
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
    
    lbl_TotalLetras.Caption = EnLetras(Format(txt_TotalNeto.Value, "0.00"), txtGls_Moneda.Text)
    txt_MontoLetras.Text = lbl_TotalLetras.Caption
    indCargando = False
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
End Sub

Private Sub cmbAyudaCliente_Click()
    
    If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Or strTipoDoc = "69" Or strTipoDoc = "97" Then
        Me.MousePointer = 11
        sw_proveedor = False
        ayuda_proveedores.Show 1
        If Len(Trim("" & cod_Prov)) > 0 Then
            txtCod_Cliente.Text = "" & cod_Prov
            txtGls_Cliente.Text = "" & des_prov
            txt_RUC.Text = "" & ruc_prov
            txt_Direccion.Text = "" & dir_prov
        End If
        Me.MousePointer = 1
    Else
        'mostrarAyudaClientes txtCod_Cliente, txtGls_Cliente, txt_RUC, txt_Direccion
        'If txtCod_Cliente.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Function traerControl(strName As String) As Object
On Error GoTo Err
Dim pCtrl As Control

    For Each pCtrl In Me
        If pCtrl.Name = strName Then
            Set traerControl = pCtrl
            Exit For
        End If
    Next
    
    Exit Function

Err:
    Exit Function
End Function

Private Sub muestraControlesCabecera()
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim pCtrl As Object
Dim strSerie As String

    strSerie = traerCampo("seriexusuario", "idSerie", "idUsuario", glsUser, True, " idDocumento = '" & strTipoDoc & "'")
    csql = "SELECT GlsObj,intLeft,intTop,tipoDato,Decimales,intTabIndex,ISNULL(Etiqueta,'') AS Etiqueta  FROM objdocventas " & _
            "where idEmpresa = '" & glsEmpresa & "' and idDocumento = '" & strTipoDoc & "' AND idSerie = '" & strSerie & "' and (tipoObj = 'C' or tipoObj = 'T') and indVisible = 'V' ORDER BY intTabIndex"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    Do While Not rst.EOF
        Set pCtrl = traerControl(rst.Fields("GlsObj"))
        pCtrl.left = Val(rst.Fields("intLeft") & "")
        pCtrl.top = Val(rst.Fields("intTop") & "")
        If (rst.Fields("tipoDato") & "") = "N" And TypeOf pCtrl Is CATTextBox Then
            pCtrl.Decimales = Val(rst.Fields("Decimales") & "")
        End If
        
        '******** CAMBIA LA ETIQUETA A LOS CAPTIÓN ************
        If (rst.Fields("tipoDato") & "") = "T" And TypeOf pCtrl Is Label Then
            If Trim(rst.Fields("Etiqueta") & "") <> "" Then
                pCtrl.Caption = Trim(rst.Fields("Etiqueta") & "")
                pCtrl.Width = Len(Trim(rst.Fields("Etiqueta") & "")) * 100
            End If
        End If
        '******************************************************
        
        pCtrl.TabIndex = Val(rst.Fields("intTabIndex") & "")
        pCtrl.Visible = True
        rst.MoveNext
    Loop
    
    Exit Sub

Err:
    MsgBox Err.Description
End Sub

Private Sub muestraColumnasDetalle()
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim pCtrl As Object
Dim strSerie As String

    strSerie = traerCampo("seriexusuario", "idSerie", "idUsuario", glsUser, True, " idDocumento = '" & strTipoDoc & "'")
    csql = "SELECT GlsObj,etiqueta,numCol,ancho,Tipodato,Decimales  FROM objdocventas " & _
            "where idEmpresa = '" & glsEmpresa & "' and idDocumento = '" & strTipoDoc & "' and idserie = '" & strSerie & "' and tipoObj = 'D' and indVisible = 'V' order by numcol"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    Do While Not rst.EOF
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Caption = rst.Fields("etiqueta") & ""
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").ColIndex = Val(rst.Fields("numCol") & "")
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Width = Val(rst.Fields("ancho") & "")
        gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").Visible = True
        If (rst.Fields("Tipodato") & "") = "N" Then
            gDetalle.Columns.ColumnByFieldName(rst.Fields("GlsObj") & "").DecimalPlaces = Val(rst.Fields("Decimales") & "")
        End If
        rst.MoveNext
    Loop
    
    Exit Sub

Err:
    MsgBox Err.Description
End Sub

Private Function DatosPrecio(ByVal strCodProd As String, ByVal strTipoProd As String, ByVal strCodUM As String, ByRef strglsum As String, ByRef dblVVUnit As Double, ByRef dblFactor As Double) As Boolean
Dim rst As New ADODB.Recordset
            
    csql = "Select TOP 1 U.AbreUM GlsUM,IIf(A.IdMoneda = 'PEN',A.ValorVenta,A.ValorVenta * " & txt_TipoCambio.Text & ") VVUnit,1  As Factor " & _
               "From ProductosProveedor A " & _
               "Inner Join Productos P " & _
                    "On A.idEmpresa = P.idEmpresa And A.idProducto = P.idProducto " & _
               "Inner Join UnidadMedida U " & _
                    "On P.idUMCompra = U.IdUM " & _
               "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdProveedor = '" & txtCod_Cliente.Text & "' And A.IdProducto = '" & strCodProd & "' " & _
               "Order By A.Fecha Desc "
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        DatosPrecio = True
        strglsum = "" & rst.Fields("GlsUM")
        dblVVUnit = "" & rst.Fields("VVUnit")
        dblFactor = "" & rst.Fields("Factor")
    Else
        DatosPrecio = False
        strglsum = ""
        dblVVUnit = 0
        dblFactor = 1
    End If
    rst.Close: Set rst = Nothing
    
End Function

Private Function DatosProducto(strCodProd As String, ByRef strCodFabri As String, ByRef strCodMar As String, ByRef strGlsMarca As String, ByRef intAfecto As Integer, ByRef strTipoProd As String) As Boolean
Dim rst As New ADODB.Recordset

    csql = "SELECT p.idFabricante,p.idMarca,m.GlsMarca,p.AfectoIGV,p.idTipoProducto " & _
            "FROM productos p LEFT JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca AND m.idEmpresa = '" & glsEmpresa & "' " & _
            "WHERE p.idEmpresa = '" & glsEmpresa & "' " & _
            "AND p.idProducto = '" & strCodProd & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If Not rst.EOF Then
        DatosProducto = True
        strCodFabri = "" & rst.Fields("idFabricante")
        strCodMar = "" & rst.Fields("idMarca")
        strGlsMarca = "" & rst.Fields("GlsMarca")
        intAfecto = "" & rst.Fields("AfectoIGV")
        strTipoProd = "" & rst.Fields("idTipoProducto")
    Else
        DatosProducto = False
        strCodFabri = ""
        strCodMar = ""
        strGlsMarca = ""
        intAfecto = 1
        strTipoProd = ""
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

Private Sub calculaTotalesFila(dblCantidad As Double, dblVVUnit As Double, dblIGVUnit As Double, dblPVUnit As Double, strDcto As String, intAfecto As Integer)
Dim dblTotalVVBruto As Double
Dim dblTotalPVBruto As Double
Dim dblDctoVV As Double
Dim dblDctoPV As Double
Dim dblTotalVVNeto As Double
Dim dblTotalIGVNeto As Double
Dim dblTotalPVNeto As Double
Dim dblDctoVVT        As Double
Dim dblDctoPVT        As Double
Dim strPorDcto()        As String
Dim strPorDctoRpt        As String
Dim i As Integer
      
    dblTotalVVBruto = dblCantidad * dblVVUnit
    dblTotalPVBruto = dblCantidad * dblPVUnit
    If Trim(strDcto) <> "" And strDcto <> "0" Then
        strPorDcto = Split(strDcto, "+")
        dblDctoVV = 0
        strPorDctoRpt = ""
        For i = 0 To UBound(strPorDcto)
            dblDctoVVT = (dblVVUnit - dblDctoVV) * (Val(strPorDcto(i)) / 100)
            dblDctoPVT = (dblPVUnit - dblDctoPV) * (Val(strPorDcto(i)) / 100)

            dblDctoVV = dblDctoVV + dblDctoVVT
            dblDctoPV = dblDctoPV + dblDctoPVT
            strPorDctoRpt = strPorDctoRpt & CStr(Val(strPorDcto(i))) & "+"
        Next
    Else
        dblDctoVV = 0
        dblDctoPV = 0
        strPorDctoRpt = "0"
    End If
    
    If Len(strPorDctoRpt) > 1 Then strPorDctoRpt = left(strPorDctoRpt, Len(strPorDctoRpt) - 1)
    gDetalle.Columns.ColumnByFieldName("VVUnitNeto").Value = dblVVUnit - dblDctoVV
    gDetalle.Columns.ColumnByFieldName("PVUnitNeto").Value = dblPVUnit - dblDctoPV
    
    dblTotalVVNeto = Val((dblCantidad * (dblVVUnit - dblDctoVV))) 'DCTO PRECIO UNITARIO  CAMBIO
    
    If intAfecto = 1 Then
        dblTotalIGVNeto = Val((dblTotalVVNeto * dblIgvNEw))
    Else
        dblTotalIGVNeto = 0
    End If
    dblTotalPVNeto = dblTotalVVNeto + dblTotalIGVNeto
    
    gDetalle.Columns.ColumnByFieldName("TotalVVBruto").Value = dblTotalVVBruto
    gDetalle.Columns.ColumnByFieldName("TotalPVBruto").Value = dblTotalPVBruto
    gDetalle.Columns.ColumnByFieldName("DctoVV").Value = dblDctoVV
    gDetalle.Columns.ColumnByFieldName("DctoPV").Value = dblDctoPV
    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = dblTotalVVNeto
    gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = dblTotalIGVNeto
    gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = dblTotalPVNeto
    gDetalle.Columns.ColumnByFieldName("porDcto").Value = strPorDctoRpt
    
End Sub

Private Sub ListaDetalle()
Dim rsdatos         As New ADODB.Recordset
Dim StrMsgError     As String
On Error GoTo Err

    If strTipoDoc = "87" Or strTipoDoc = "OS" Or strTipoDoc = "69" Or strTipoDoc = "97" Then
        gDetalle.Columns.ColumnByFieldName("PorDcto").Visible = False
    Else
        gDetalle.Columns.ColumnByFieldName("PorDcto").Visible = True
    End If
    
    If gLista.Count = 0 Then Exit Sub
    
    csql = "SELECT item, idProducto, CAST(GlsProducto AS VARCHAR (500)) AS GlsProducto, GlsMarca, GlsUM, CAST(Cantidad AS DECIMAL(10,2)) AS Cantidad, CAST(VVUnit AS DECIMAL(10,2)) AS PVUnit, CAST(PorDcto AS DECIMAL(10,2)) AS PorDcto, CAST(TotalVVNeto AS DECIMAL(10,2)) AS TotalPVNeto " & _
           "FROM docventasdet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'  AND idDocumento = '" & strTipoDoc & _
           "' AND idDocVentas = '" & gLista.Columns.ColumnByFieldName("idDocVentas").Value & _
           "' AND idSerie = '" & gLista.Columns.ColumnByFieldName("idSerie").Value & "'"
           
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
    
    Set gListaDetalle.DataSource = rsdatos
    
'    csql = "SELECT item, idProducto, GlsProducto, GlsMarca, GlsUM, Format(Cantidad,2) AS Cantidad, Format(PVUnit,2) AS PVUnit, FORMAT(PorDcto,2) AS PorDcto, Format(TotalPVNeto,2) AS TotalPVNeto " & _
'           "FROM docventasdet " & _
'           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'  AND idDocumento = '" & strTipoDoc & _
'           "' AND idDocVentas = '" & gLista.Columns.ColumnByFieldName("idDocVentas").Value & _
'           "' AND idSerie = '" & gLista.Columns.ColumnByFieldName("idSerie").Value & "'"
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
'
    
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
End Sub

Private Sub txtCod_Almacen_Change()
    
    txtGls_Almacen.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", txtCod_Almacen.Text, True)

End Sub

Private Sub txtCod_Almacen_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "ALMACEN", txtCod_Almacen, txtGls_Almacen
        KeyAscii = 0
        If txtCod_Almacen.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_CentroCosto_Change()
    
    txtGls_CentroCosto.Text = traerCampo("centroscosto", "GlsCentroCosto", "idCentroCosto", txtCod_CentroCosto.Text, True)

End Sub

Private Sub txtCod_CentroCosto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CENTROCOSTO", txtCod_CentroCosto, txtGls_CentroCosto
        KeyAscii = 0
        If txtCod_CentroCosto.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Chofer_Change()
    
    If indCargando = False And txtCod_Chofer.Text <> "" Then
        txt_Brevete.Text = traerCampo("choferes", "NroBrevete", "idChofer", Trim(txtCod_Chofer.Text), True)
    End If

End Sub

Private Sub txtCod_Chofer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CHOFER", txtCod_Chofer, txtGls_Chofer
        KeyAscii = 0
        If txtCod_Chofer.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Cliente_Change()
On Error GoTo Err
Dim StrMsgError         As String
Dim rst                 As New ADODB.Recordset
Dim strCodAlmUsu        As String

    strCodAlmUsu = Trim("" & traerCampo("Usuarios", "idAlmacen", "idUsuario", glsUser, True))

    If indCargando = False And txtCod_Cliente.Text <> "" Then
        csql = "SELECT p.ruc,concat(p.direccion,' ',u.glsUbigeo,' ',d.glsUbigeo) as direccion,p.GlsPersona,p.direccionEntrega " & _
               "FROM personas p LEFT JOIN ubigeo u ON P.idDistrito = u.idDistrito AND p.idPais = u.idPais  LEFT JOIN ubigeo d ON left(u.idDistrito,2) = d.idDpto AND d.idProv = '00' AND d.idDist = '00'   AND u.idPais = d.idPais  " & _
               "Where p.idPersona = '" & txtCod_Cliente.Text & "'"
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rst.EOF Then
            txt_RUC.Text = "" & rst.Fields("ruc")
            txt_Direccion.Text = "" & rst.Fields("direccion")
            If (strTipoDoc <> "94" Or strTipoDoc <> "87") Then
                If strCodAlmUsu <> "" Then
                   txt_Llegada.Text = traerCampo("Almacenes", "GlsDireccion", "idAlmacen", strCodAlmUsu, True)
                Else
                    If txt_Llegada.Visible Then txt_Llegada.Text = Trim("" & traerCampo("Personas", "direccionEntrega", "idpersona", Trim("" & traerCampo("Empresas", "idpersona", "idempresa", glsEmpresa, False)), False))
                End If
                txtGls_Cliente.Text = "" & rst.Fields("GlsPersona")
                txtCod_VendedorCampo.Text = traerCampo("clientes", "idVendedorCampo", "idCliente", txtCod_Cliente.Text, True)
                txtCod_EmpTrans.Text = traerCampo("clientes", "idEmpTrans", "idCliente", txtCod_Cliente.Text, True)
                dblPorDsctoEspecial = Val(traerCampo("clientes", "Val_Dscto", "idCliente", txtCod_Cliente.Text, True) & "")
            End If
        End If
        rst.Close: Set rst = Nothing
        
        traerListaPrecios StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
    Else
        txt_RUC.Text = ""
        txt_Direccion.Text = ""
        If strTipoDoc <> "94" Or strTipoDoc <> "87" Then
            If txt_Llegada.Visible Then txt_Llegada.Text = ""
        End If
        txtGls_Cliente.Text = ""
        txtCod_VendedorCampo.Text = ""
        txtCod_EmpTrans.Text = ""
        dblPorDsctoEspecial = 0
    End If
    
    '--- Asignamos moneda segun proveedor
    txtCod_Moneda.Text = traerCampo("Proveedores", "idMoneda", "idProveedor", txtCod_Cliente.Text, True)
    If txtCod_Moneda.Text = "" Then
        txtCod_Moneda.Text = glsMonVentas
    End If
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Cliente_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Then
            mostrarAyuda "PROVEEDOR", txtCod_Cliente, txtGls_Cliente
        Else
            'mostrarAyudaClientesKeyascii KeyAscii, txtCod_Cliente, txtGls_Cliente, txt_RUC, txt_Direccion
        End If
        KeyAscii = 0
        If txtCod_Cliente.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtCod_contacto_Change()

    txtgls_contacto.Text = traerCampo("PERSONAS", "GLSPERSONA", "IDPERSONA", txtCod_contacto.Text, False)

End Sub

Private Sub txtCod_EmpTrans_Change()
    
    txtGls_EmpTrans.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_EmpTrans.Text, False)
    txt_RUCEmp.Text = traerCampo("personas", "ruc", "idPersona", txtCod_EmpTrans.Text, False)
    
    If glsPersonaEmpresa <> "" Then
        If glsPersonaEmpresa <> txtCod_EmpTrans.Text Then
            'If txt_Partida2.Text = "" Then txt_Partida2.Text = txt_Llegada.Text
            If txt_Llegada2.Text = "" Then txt_Llegada2.Text = txt_Direccion.Text
        End If
    End If
    
End Sub

Private Sub txtCod_EmpTrans_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "EMPTRANS", txtCod_EmpTrans, txtGls_EmpTrans
        KeyAscii = 0
        If txtCod_EmpTrans.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_FormaPago_Change()
    
    If txtCod_FormaPago.Text = "" Then
        txtGls_FormaPago.Text = ""
    Else
        txtGls_FormaPago.Text = traerCampo("formaspagos", "glsFormaPago", "idFormaPago", txtCod_FormaPago.Text, False)
    End If

End Sub

Private Sub txtCod_Lista_Change()
    
    txtGls_Lista.Text = traerCampo("listaprecios", "GlsLista", "idLista", txtCod_Lista.Text, True)

End Sub

Private Sub txtCod_Lista_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "LISTA", txtCod_Lista, txtGls_Lista
        KeyAscii = 0
        If txtCod_Lista.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
    lbl_SimbMonBruto.Caption = traerCampo("monedas", "Simbolo", "idMoneda", txtCod_Moneda.Text, False)
    lbl_SimbMonIGV.Caption = lbl_SimbMonBruto.Caption
    lbl_SimbMonNeto.Caption = lbl_SimbMonBruto.Caption
    txt_SimboloMonBruto.Text = lbl_SimbMonBruto.Caption
    txt_SimboloMonIGV.Text = lbl_SimbMonBruto.Caption
    txt_SimboloMonNeto.Text = lbl_SimbMonBruto.Caption
    
    If Len(Trim(txtCod_Moneda.Text)) > 0 Then
        If traerCampo("parametros", "valparametro", "GLSPARAMETRO", "VIZUALIZA_SON", True) = "S" Then
            lbl_TotalLetras.Caption = "SON:" & Cadenanum(Format(txt_TotalNeto.Value, "0.00"), txtGls_Moneda.Text)
        Else
            lbl_TotalLetras.Caption = Cadenanum(Format(txt_TotalNeto.Value, "0.00"), txtGls_Moneda.Text)
        End If
        txt_MontoLetras.Text = lbl_TotalLetras.Caption
    End If

End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MONEDA", txtCod_Moneda, txtGls_Moneda
        KeyAscii = 0
        If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_MotivoNCD_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MOTIVONCD", txtCod_MotivoNCD, txtGls_MotivoNCD
        KeyAscii = 0
        If txtCod_MotivoNCD.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_MotivoTraslado_Change()
    
    txtGls_MotivoTraslado.Text = traerCampo("motivostraslados", "GlsMotivoTraslado", "idMotivoTraslado", txtCod_MotivoTraslado.Text, False)

End Sub

Private Sub txtCod_MotivoTraslado_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MOTIVOTRASLADO", txtCod_MotivoTraslado, txtGls_MotivoTraslado
        KeyAscii = 0
        If txtCod_MotivoTraslado.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_TipoTicket_Change()
    
    txtGls_TipoTicket.Text = traerCampo("datos", "GlsDato", "idDato", txtCod_TipoTicket.Text, False)

End Sub

Private Sub txtCod_TipoTicket_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "TIPOTICKET", txtCod_TipoTicket, txtGls_TipoTicket
        KeyAscii = 0
        If txtCod_TipoTicket.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_UnidProd_Change()
    
    If txtCod_UnidProd.Text = "" Then
        txtGls_UnidProd.Text = ""
    Else
        txtGls_UnidProd.Text = Trim("" & traerCampo("unidadproduccion", "Descunidad", "CodUnidProd", txtCod_UnidProd.Text, True))
        txtCod_Almacen.Text = Trim("" & traerCampo("unidadproduccion", "IdAlmacen", "CodUnidProd", txtCod_UnidProd.Text, True))
    End If

End Sub

Private Sub txtCod_Vehiculo_Change()
Dim rst As New ADODB.Recordset
    
    If indCargando = False And txtCod_Vehiculo.Text <> "" Then
        csql = "SELECT GlsPlaca,GlsDato as Marca, GlsModelo, GlsColor, GlsCodInscripcion " & _
               "FROM vehiculos,datos " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND idMarcaVehi = idDato AND idVehiculo = '" & txtCod_Vehiculo.Text & "'"
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rst.EOF Then
            txt_Placa.Text = "" & rst.Fields("GlsPlaca")
            txt_Marca.Text = "" & rst.Fields("Marca")
            txt_Modelo.Text = "" & rst.Fields("GlsModelo")
            txt_Color.Text = "" & rst.Fields("GlsColor")
            txt_CodInscripcion.Text = "" & rst.Fields("GlsCodInscripcion")
        End If
        rst.Close: Set rst = Nothing
    End If

End Sub

Private Sub txtCod_Vehiculo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "VEHICULO", txtCod_Vehiculo, txtGls_Vehiculo
        KeyAscii = 0
        If txtCod_Vehiculo.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Vendedor_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)
    traerListaPrecios StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub valoresIniciales()
Dim strAlmxDef      As String

    strAlmxDef = Trim("" & traerCampo("Usuarios", "idAlmacen", "idUsuario", glsUser, True))
    
    txt_Serie.Text = traerCampo("seriexusuario", "idSerie", "idUsuario", glsUser, True, " idDocumento = '" & strTipoDoc & "'")
    'txt_TipoCambio.Text = glsTC
    txt_TipoCambio.Text = Val(traerCampo("TiposDeCambio", "TcVenta", "Fecha", Format(dtp_Emision.Value, "yyyy-mm-dd"), False))
    txtCod_Moneda.Text = glsMonVentas
    
    If traerCampo("vendedores", "idVendedor", "idVendedor", glsUser, True) <> "" Then
        txtCod_Vendedor.Text = glsUser
    End If
    
    'PROVEEDOR_POR_DEFECTO
    'Si el TD es Requeriento Asignamos el proveedor por defecto
     If strTipoDoc = "87" Then
        txtCod_Cliente.Text = Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "PROVEEDOR_POR_DEFECTO", True))
     End If
    
    'Si el campo idAlmacen de la Tabla Usuarios esta lleno asignamos el almacen y el punto de partida si NO se asigna el valor del parametro
     If strTipoDoc = "94" And strAlmxDef <> "" Then
        txtCod_Almacen.Text = strAlmxDef
        txt_Llegada.Text = traerCampo("Almacenes", "GlsDireccion", "idAlmacen", strAlmxDef, True)
     Else
        txtCod_Almacen.Text = Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "ALMACEN_COMPRAS", True))
     End If
     
     'Si el campo idCentroCosto de la Tabla Usuarios esta lleno asignamos el dato
     If strTipoDoc = "94" Then
        txtCod_CentroCosto.Text = Trim("" & traerCampo("Usuarios", "idCentroCosto", "idUsuario", glsUser, True))
     End If
     
End Sub

Private Sub anularDoc(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim iniTrans As Boolean
Dim strNumVale As String
Dim strMovCaja As String
Dim IndEvaluacion As Integer
Dim strCodUsuarioAutorizacion As String
Dim cnn_empresa As New ADODB.Connection
Dim rscta       As New ADODB.Recordset
Dim cabrev      As String, cdocumento As String, cbusca As String, cdirectorio As String, cruta As String
Dim cconex_empresa  As String
Dim ncorrela        As Double

    getEstadoCierreMes Format(dtp_Emision.Value, "dd/mm/yyyy"), StrMsgError
    If StrMsgError <> "" Then GoTo Err

    If MsgBox("Seguro de Anular el Documento", vbQuestion + vbYesNo, App.Title) = vbYes Then
        csql = "select idProducto, glsProducto, Cantidad, CantidadImp " & _
                 "from docventasdet where CantidadImp <> 0 and iddocumento = '" & strTipoDoc & "' " & _
                 "and idserie = '" & txt_Serie.Text & "' and iddocventas = '" & txt_NumDoc.Text & _
                 "' and idempresa = '" & glsEmpresa & "' And idSucursal = '" & glsSucursal & "'"
        If rst.State = adStateOpen Then rst.Close
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If rst.RecordCount <> 0 Then
            StrMsgError = "No se puede anular el documento por que ha sido importado en un vale"
            GoTo Err
        End If
        If rst.State = adStateOpen Then rst.Close: Set rst = Nothing
        
        IndEvaluacion = 0
        frmAprobacion.MostrarForm "01", IndEvaluacion, strCodUsuarioAutorizacion, StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        If IndEvaluacion = 0 Then Exit Sub
            
        Cn.BeginTrans
        iniTrans = True
    
        'Anulando el Documento
        csql = "UPDATE docventas SET estDocventas = 'ANU', idUsuarioAnulacion = '" & strCodUsuarioAutorizacion & _
              "' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & _
              "' AND idDocumento = '" & strTipoDoc & "' and idDocVentas = '" & Trim(txt_NumDoc.Text) & _
              "' and idSerie = '" & Trim(txt_Serie.Text) & "'"
        Cn.Execute csql
        
        '----------------------------------------------------------------------------------
        '---- ANULAMOS DE CTASXCOBRAR - ACCESS
        cdirectorio = traerCampo("empresas", "Carpeta", "idEmpresa", glsEmpresa, False)
        If Len(Trim(cdirectorio)) > 0 Then
            cruta = glsRuta_Access & cdirectorio
            If cnn_empresa.State = adStateOpen Then cnn_empresa.Close
            cconex_empresa = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cruta & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
            cnn_empresa.Open cconex_empresa
                    
            cabrev = traerCampo("documentos", "AbreDocumento", "idDocumento", strTipoDoc, False)
            cdocumento = cabrev & Format(txt_Serie.Text, "000") & "/" & Format(txt_NumDoc.Text, "0000000")
            
            cbusca = "SELECT * FROM CTA_DCTO WHERE NRO_COMP='" & cdocumento & "' and idempresa = '" & glsEmpresa & "'"
            If rscta.State = adStateOpen Then rscta.Close
            rscta.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
            If Not rscta.EOF Then
                ncorrela = Val(rscta.Fields("CORRELA") & "")
                cbusca = "UPDATE CTA_DCTO SET TOTAL=0,TOTALO=0,SALDO=0,ANULADO='A' WHERE CORRELA=" & ncorrela & ""
                cnn_empresa.Execute (cbusca)
            End If
            rscta.Close: Set rscta = Nothing
        End If
        
        'Capturamos datos a anular
        csql = "SELECT idValesCab,idMovCaja " & _
               "FROM docventas " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTipoDoc & "' and idDocVentas = '" & Trim(txt_NumDoc.Text) & "' and idSerie = '" & Trim(txt_Serie.Text) & "'"
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        'Anulando el vale
        If indGeneraVale Then
            strNumVale = "" & rst.Fields("idValesCab")
            
            If strNumVale <> "" Then
                'actualizaStock strNumVale, 1, StrMsgError, False
                'If StrMsgError <> "" Then GoTo Err
            End If
            
            csql = "UPDATE valescab SET estValeCab = 'ANU' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idValesCab = '" & strNumVale & "'"
            Cn.Execute csql
        End If
        
        strMovCaja = "" & rst.Fields("idMovCaja")
        
        csql = "SELECT d.idDocumentoImp,d.idDocVentasImp,d.idSerieImp,d.idProducto,d.idUM,d.Cantidad " & _
                "FROM docventasdet d " & _
                "WHERE d.idEmpresa = '" & glsEmpresa & "'" & _
                "AND d.idSucursal = '" & glsSucursal & "'" & _
                "AND d.idDocumento = '" & strTipoDoc & "'" & _
                "AND d.idDocVentas = '" & Trim(txt_NumDoc.Text) & "'" & _
                "AND d.idSerie = '" & Trim(txt_Serie.Text) & "'" & _
                "AND d.idDocVentasImp <> ''"
                
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            'ACTUALIZAMOS CANTIDAD IMPORTADA
            csql = "UPDATE docventasdet dd  SET dd.CantidadImp = dd.CantidadImp - " & CStr(rst.Fields("Cantidad")) & ", dd.estDocImportado = 'N' " & _
                   "WHERE dd.idEmpresa = '" & glsEmpresa & "' AND dd.idSucursal = '" & glsSucursal & "' AND dd.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND dd.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND dd.idSerie = '" & rst.Fields("idSerieImp") & "' " & _
                     "AND dd.idProducto = '" & rst.Fields("idProducto") & "' AND dd.idUM = '" & rst.Fields("idUM") & "'"
            Cn.Execute csql
            rst.MoveNext
        Loop
        
        csql = "SELECT DISTINCT idDocumentoImp,idDocVentasImp,idSerieImp  FROM docventasdet " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTipoDoc & "' AND idDocVentas = '" & Trim(txt_NumDoc.Text) & "' AND idSerie = '" & Trim(txt_Serie.Text) & "' AND idDocVentasImp <> '' AND ISNULL(idDocVentasImp) = False"
        If rst.State = 1 Then rst.Close
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            csql = "UPDATE docVentas c  SET c.estDocImportado = 'N'  ,  c.estDocventas = 'GEN '" & _
                   "WHERE c.idEmpresa = '" & glsEmpresa & "' AND c.idSucursal = '" & glsSucursal & "' AND c.idDocumento = '" & rst.Fields("idDocumentoImp") & "' AND c.idDocVentas = '" & rst.Fields("idDocVentasImp") & "' AND c.idSerie = '" & rst.Fields("idSerieImp") & "'"
            Cn.Execute csql
            rst.MoveNext
        Loop
        
        Cn.CommitTrans
        
        'Cambiando valores a variables
        strEstDocVentas = "ANU"
        lblDoc.ForeColor = &HFF&
        lblDoc.Caption = lblDoc.Caption & " - ANULADA"
        fraGeneral.Enabled = False
        fraDetalle.Enabled = False
        
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        habilitaBotones 6
    End If
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub

Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If iniTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Vendedor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor
        KeyAscii = 0
        If txtCod_Vendedor.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_VendedorCampo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        If glsModVendCampo = False Then
            KeyAscii = 0
            Exit Sub
        End If
        mostrarAyudaKeyascii KeyAscii, "VENDEDOR", txtCod_VendedorCampo, txtGls_VendedorCampo
        KeyAscii = 0
        If txtCod_VendedorCampo.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub calcularTotales()
Dim DblImponible As Double
Dim DblExonerado As Double
Dim DblTotIgv    As Double
    
    txt_TotalBruto.Text = 0#
    txt_TotalIGV.Text = 0#
    txt_TotalNeto.Text = 0#
    
    txt_TotalDsctoVV.Text = 0#
    txt_TotalDsctoPV.Text = 0#
    txt_TotalBaseImponible.Text = 0#
    txt_TotalExonerado.Text = 0#
    
    DblImponible = 0#
    DblExonerado = 0#
    DblTotIgv = 0#
    
    gDetalle.Dataset.DisableControls
    
    gDetalle.Dataset.First
    Do While Not gDetalle.Dataset.EOF
        txt_TotalBruto.Text = Val(txt_TotalBruto.Value + gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value)
        '---- CALCULO DEL PRECIO UNITARIO
        txt_TotalDsctoVV.Text = Val(txt_TotalDsctoVV.Value + ((gDetalle.Columns.ColumnByFieldName("TotalVVBruto").Value) - (gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value)))
        txt_TotalDsctoPV.Text = Val(txt_TotalDsctoPV.Value + ((gDetalle.Columns.ColumnByFieldName("TotalPVBruto").Value) - (gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value)))
        
        If gDetalle.Columns.ColumnByFieldName("Afecto").Value = 0 Then
            txt_TotalExonerado.Text = Val(txt_TotalExonerado.Value + gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value)
            
            DblExonerado = Val(DblExonerado + gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value)
        Else
            txt_TotalBaseImponible.Text = Val(txt_TotalBaseImponible.Value + gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value)
            DblImponible = Val(DblImponible + gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value)
        End If
        
        gDetalle.Dataset.Next
    Loop
     
    txt_TotalIGV.Text = Val(Format((DblImponible * dblIgvNEw), "0.00"))
    DblTotIgv = DblImponible * dblIgvNEw
    txt_TotalNeto.Text = Val(Format(txt_TotalBaseImponible.Text, "0.00")) + Val(Format(txt_TotalExonerado.Text, "0.00")) + Val(Format(txt_TotalIGV.Text, "0.00"))
    
    gDetalle.Dataset.EnableControls

    lbl_TotalLetras.Caption = EnLetras(Format(txt_TotalNeto.Value, "0.00"), txtGls_Moneda.Text)
    txt_MontoLetras.Text = lbl_TotalLetras.Caption
    
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

Private Sub ocultarColumnasEstado()

    Select Case strEstDocVentas
        Case "GEN"
            If strTipoDoc = "86" Or strTipoDoc = "92" Or strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Or strTipoDoc = "97" Then  'Guia ----Or strTipoDoc = "40" Pedido
                Toolbar1.Buttons(7).Visible = True 'IMPRIMIR
            Else
                Toolbar1.Buttons(7).Visible = False 'IMPRIMIR
            End If
        Case "ANU"
            Toolbar1.Buttons(2).Visible = False 'GRABAR
            Toolbar1.Buttons(3).Visible = False 'MODIFICAR
            Toolbar1.Buttons(5).Visible = False 'ELIMINAR
            Toolbar1.Buttons(6).Visible = False 'ANULAR
            Toolbar1.Buttons(7).Visible = False 'IMPRIMIR
            Toolbar1.Buttons(8).Visible = True 'LISTAR
    End Select

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
            End If
            gDocReferencia.Dataset.Next
        Loop
    End If

End Sub

Private Sub mostrarDocImportado2(ByVal rsdd As ADODB.Recordset, ByRef StrMsgError As String)
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
Dim primero As Boolean
    
    primero = True
    rsdd.MoveFirst
    Do While Not rsdd.EOF
        If primero = True Then
            primero = False
        Else
            gDetalle.Dataset.Insert
        End If
        gDetalle.SetFocus
        gDetalle.Dataset.RecNo = intFila
        gDetalle.Dataset.Edit
        gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" & rsdd.Fields("idProducto")
        gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = "" & rsdd.Fields("CodigoRapido")
        gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = "" & rsdd.Fields("GlsProducto")
        strCodUM = traerCampo("productos", "idUMCompra", "idProducto", "" & rsdd.Fields("idProducto"), True)
        If Trim("" & rsdd.Fields("idProducto")) = "" Then Exit Sub
        If DatosProducto("" & rsdd.Fields("idProducto"), strCodFabri, strCodMar, strDesMar, intAfecto, strTipoProd) = False Then
        End If

        strMoneda = "PEN"
        
        gDetalle.Columns.ColumnByFieldName("idCodFabricante").Value = strCodFabri
        gDetalle.Columns.ColumnByFieldName("idMarca").Value = strCodMar
        gDetalle.Columns.ColumnByFieldName("GlsMarca").Value = strDesMar
        gDetalle.Columns.ColumnByFieldName("Afecto").Value = intAfecto
        gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value = strTipoProd
        gDetalle.Columns.ColumnByFieldName("idMoneda").Value = strMoneda 'falta esta columna en el detalle de la grilla

        If DatosPrecio("" & rsdd.Fields("idProducto"), strTipoProd, strCodUM, strDesUM, dblVVUnit, dblFactor) = False Then
        End If
        If strDesUM = "" And strCodUM <> "" Then strDesUM = traerCampo("unidadMedida", "abreUM", "idUM", strCodUM, False)

        gDetalle.Columns.ColumnByFieldName("idUM").Value = strCodUM
        gDetalle.Columns.ColumnByFieldName("GlsUM").Value = strDesUM
        gDetalle.Columns.ColumnByFieldName("Factor").Value = dblFactor
        If strTipoProd = "06002" Then gDetalle.Columns.ColumnByFieldName("Cantidad").Value = 1
        procesaMoneda strMoneda, txtCod_Moneda.Text, 0, dblVVUnit, intAfecto, dblVVUnit, dblIGVUnit, dblPVUnit
        gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
        gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
        gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
        gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = dblVVUnit
        gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = dblPVUnit
        gDetalle.Columns.ColumnByFieldName("PorDcto").Value = dblPorDsctoEspecial
        gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
        gDetalle.Columns.ColumnByFieldName("FechaEmision").Value = getFechaSistema
        gDetalle.Columns.ColumnByFieldName("GlsPlaca").Value = ""
        
        gDetalle.Dataset.Post
        gDetalle.Dataset.RecNo = intFila
        gDetalle.Dataset.Edit

        calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
        gDetalle.Dataset.Post
        If "" & rsdd.Fields("idProducto") <> "" Then
            gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").Index
        End If
        rsdd.MoveNext
    Loop

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
 
Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsValida As New ADODB.Recordset
Dim indTrans As Boolean
Dim strNumOpe As String
Dim strSerie As String
Dim strCodVale As String
Dim strNumMovCaja As String

    getEstadoCierreMes Format(dtp_Emision.Value, "dd/mm/yyyy"), StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If MsgBox("¿Seguro de eliminar el documento?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    
    strNumOpe = Trim(txt_NumDoc.Text)
    strSerie = Trim(txt_Serie.Text)
    
    csql = "select idProducto, glsProducto, Cantidad, CantidadImp " & _
             "from docventasdet " & _
             "where iddocumento = '" & strTipoDoc & "' " & _
             "and iddocventas = '" & txt_NumDoc.Text & " ' and idserie = '" & txt_Serie.Text & _
             "' and idempresa = '" & glsEmpresa & "' and idSucursal  = '" & glsSucursal & "' And CantidadImp <> 0  "
    If rst.State = adStateOpen Then rst.Close
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If rst.RecordCount <> 0 Then
        StrMsgError = "No se puede eliminar el documento por que ha sido importado en un vale"
        GoTo Err
    End If
    If rst.State = adStateOpen Then rst.Close: Set rst = Nothing

    Cn.BeginTrans
    indTrans = True
    strCodVale = Trim(traerCampo("docventas", "idValesCab", "idDocumento", strTipoDoc, True, " idDocVentas = '" & strNumOpe & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'"))

    If strCodVale <> "" Then
        'actualizaStock strCodVale, 1, StrMsgError, False
        'If StrMsgError <> "" Then GoTo Err
    End If
    
    csql = "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = '99' AND numDocOrigen = '" & strCodVale & "' AND serieDocOrigen = '000'"
    Cn.Execute csql
    
    csql = "DELETE FROM valesdet WHERE idValesCab = '" & strCodVale & "' AND idEmpresa = '" & glsEmpresa & "'  AND idSucursal = '" & glsSucursal & "' And  tipoVale  ='I' "
    Cn.Execute csql
    
    csql = "DELETE FROM valescab WHERE idValesCab = '" & strCodVale & "' AND idEmpresa = '" & glsEmpresa & "'  AND idSucursal = '" & glsSucursal & "' And  tipoVale  ='I' "
    Cn.Execute csql
    
    'Eliminando docreferencias
    csql = "DELETE FROM docreferencia WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND tipoDocOrigen = '" & strTipoDoc & "' AND numDocOrigen = '" & strNumOpe & "' AND serieDocOrigen = '" & strSerie & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM docventasdet WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' " & _
              "AND idDocumento = '" & strTipoDoc & "' AND idDocVentas = '" & strNumOpe & "' AND idSerie = '" & strSerie & "'"
    Cn.Execute csql
    
    csql = "DELETE FROM docventas WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' " & _
              "AND idDocumento = '" & strTipoDoc & "' AND idDocVentas = '" & strNumOpe & "' AND idSerie = '" & strSerie & "'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    MsgBox "Registro eliminado satisfactoriamente.", vbInformation, App.Title
    
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub
    
Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
 
Private Sub traerListaPrecios(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodLista As String

    If txtGls_Cliente.Text <> "" Then
        strCodLista = traerCampo("clientes", "IdLista", "idCliente", txtCod_Cliente.Text, True)
    End If
    
    If txtGls_Vendedor.Text <> "" And strCodLista = "" Then
        strCodLista = traerCampo("vendedores", "IdLista", "idVendedor", txtCod_Vendedor.Text, True)
    End If
    If strCodLista <> "" Then txtCod_Lista.Text = strCodLista
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtGls_Cliente_KeyPress(KeyAscii As Integer)

'''''    If KeyAscii <> 13 Then
'''''        mostrarAyudaClientesKeyascii KeyAscii, txtCod_Cliente, txtGls_Cliente, txt_RUC, txt_Direccion
'''''        KeyAscii = 0
'''''        If txtGls_Cliente.Text <> "" Then SendKeys "{tab}"
'''''    End If

End Sub

Private Sub mostrarDocImportado(ByVal rscd As ADODB.Recordset, ByVal rsdd As ADODB.Recordset, ByVal strTipoDocImportado As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg                         As New ADODB.Recordset
Dim RsD                         As New ADODB.Recordset
Dim strSerieDocVentas           As String
Dim i                           As Integer
Dim indExisteDocRef             As Boolean
Dim cAuxIdProducto              As String

    indCargando = True
    i = 0
    strEstDocVentas = "GEN"
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "numOrdenCompra", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idProducto", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "idCodFabricante", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idMarca", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsMarca", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Cantidad2", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVBruto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVBruto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PorDcto", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "DctoVV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "DctoPV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idTipoProducto", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idDocumentoImp", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "idDocVentasImp", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerieImp", adVarChar, 4, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idUsuarioDcto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "VVUnitLista", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnitLista", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnitNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnitNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IdCentroCosto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdSucursalPres", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdDocumentoPres", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "IdSeriePres", adVarChar, 3, adFldIsNullable
    rsg.Fields.Append "IdDocVentasPres", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "FechaEmision", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "GlsPlaca", adVarChar, 50, adFldIsNullable
    rsg.Open
    
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adChar, 8, adFldIsNullable
    RsD.Open , , adOpenKeyset, adLockOptimistic
    
    If rscd.RecordCount = 0 Then
        txt_OrdenCompra.Text = ""
        txt_Partida.Text = ""
        txtCod_Almacen.Text = ""
        txtCod_FormaPago.Text = ""
        txtObs.Text = ""
        txtCod_Moneda.Text = ""
        txtCod_Vendedor.Text = ""
        dtp_Emision.Value = getFechaSistema
    Else
        txt_Partida.Text = "" & rscd.Fields("Partida")
        txtCod_Almacen.Text = "" & rscd.Fields("idAlmacen")
        txtObs.Text = "" & rscd.Fields("ObsDocVentas")
    End If
        
    If rsdd.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("CodigoRapido") = ""
        rsg.Fields("idCodFabricante") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("idMarca") = ""
        rsg.Fields("GlsMarca") = ""
        rsg.Fields("idUM") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("Factor") = 1
        rsg.Fields("Afecto") = 1
        rsg.Fields("Cantidad") = 0
        rsg.Fields("Cantidad2") = 0
        rsg.Fields("VVUnit") = 0
        rsg.Fields("IGVUnit") = 0
        rsg.Fields("PVUnit") = 0
        rsg.Fields("TotalVVBruto") = 0
        rsg.Fields("TotalPVBruto") = 0
        rsg.Fields("PorDcto") = "0"
        rsg.Fields("DctoVV") = 0
        rsg.Fields("DctoPV") = 0
        rsg.Fields("TotalVVNeto") = 0
        rsg.Fields("TotalIGVNeto") = 0
        rsg.Fields("TotalPVNeto") = 0
        rsg.Fields("idTipoProducto") = ""
        rsg.Fields("idMoneda") = ""
        rsg.Fields("idDocumentoImp") = ""
        rsg.Fields("idDocVentasImp") = ""
        rsg.Fields("idSerieImp") = ""
        rsg.Fields("NumLote") = ""
        rsg.Fields("FecVencProd") = ""
        rsg.Fields("VVUnitLista") = 0
        rsg.Fields("PVUnitLista") = 0
        rsg.Fields("VVUnitNeto") = 0
        rsg.Fields("PVUnitNeto") = 0
        rsg.Fields("IdCentroCosto") = ""
        rsg.Fields("IdSucursalPres") = ""
        rsg.Fields("IdDocumentoPres") = ""
        rsg.Fields("IdSeriePres") = ""
        rsg.Fields("IdDocVentasPres") = ""
        rsg.Fields("FechaEmision") = getFechaSistema
        rsg.Fields("GlsPlaca") = ""
        
    Else
        rsdd.MoveFirst
        rsdd.Sort = "idProducto"
        Do While Not rsdd.EOF
            cAuxIdProducto = "" & rsdd.Fields("idProducto")
            rsg.AddNew
            i = i + 1
            rsg.Fields("Item") = i
            rsg.Fields("idProducto") = "" & rsdd.Fields("idProducto")
            rsg.Fields("CodigoRapido") = traerCampo("Productos", "CodigoRapido", "IdProducto", "" & rsdd.Fields("IdProducto"), True)
            rsg.Fields("idCodFabricante") = "" & rsdd.Fields("idCodFabricante")
            rsg.Fields("GlsProducto") = "" & rsdd.Fields("GlsProducto")
            rsg.Fields("idMarca") = "" & rsdd.Fields("idMarca")
            rsg.Fields("GlsMarca") = "" & rsdd.Fields("GlsMarca")
            rsg.Fields("idUM") = "" & rsdd.Fields("idUM")
            rsg.Fields("GlsUM") = "" & rsdd.Fields("GlsUM")
            rsg.Fields("Factor") = "" & rsdd.Fields("Factor")
            rsg.Fields("Afecto") = "" & rsdd.Fields("Afecto")
            rsg.Fields("idTipoProducto") = "" & rsdd.Fields("idTipoProducto")
            rsg.Fields("idMoneda") = "" & rsdd.Fields("idMoneda")
            rsg.Fields("idDocumentoImp") = strTipoDocImportado
            rsg.Fields("idDocVentasImp") = "" & rsdd.Fields("idDocVentas")
            rsg.Fields("idSerieImp") = "" & rsdd.Fields("idSerie")
            rsg.Fields("NumLote") = "" & rsdd.Fields("NumLote")
            rsg.Fields("FecVencProd") = "" & rsdd.Fields("FecVencProd")
            rsg.Fields("VVUnit") = Val("" & rsdd.Fields("VVUnit"))
            rsg.Fields("IGVUnit") = Val("" & rsdd.Fields("IGVUnit"))
            rsg.Fields("PVUnit") = Val("" & rsdd.Fields("PVUnit"))
            rsg.Fields("VVUnitLista") = Val("" & rsdd.Fields("VVUnitLista"))
            rsg.Fields("PVUnitLista") = Val("" & rsdd.Fields("PVUnitLista"))
            rsg.Fields("VVUnitNeto") = Val("" & rsdd.Fields("VVUnitNeto"))
            rsg.Fields("PVUnitNeto") = Val("" & rsdd.Fields("PVUnitNeto"))
            rsg.Fields("IdCentroCosto") = ""
            rsg.Fields("IdSucursalPres") = ""
            rsg.Fields("IdDocumentoPres") = ""
            rsg.Fields("IdSeriePres") = ""
            rsg.Fields("IdDocVentasPres") = ""
            rsg.Fields("FechaEmision") = getFechaSistema
            rsg.Fields("GlsPlaca") = ""
    
            Do While cAuxIdProducto = "" & rsdd.Fields("idProducto")
                rsg.Fields("Cantidad") = Val("" & rsg.Fields("Cantidad")) + Val("" & rsdd.Fields("Cantidad"))
                rsg.Fields("TotalVVBruto") = Val("" & rsg.Fields("TotalVVBruto")) + Val("" & rsdd.Fields("TotalVVBruto"))
                rsg.Fields("TotalPVBruto") = Val("" & rsg.Fields("TotalPVBruto")) + Val("" & rsdd.Fields("TotalPVBruto"))
                rsg.Fields("PorDcto") = Val("" & rsg.Fields("PorDcto")) + Val("" & rsdd.Fields("PorDcto"))
                rsg.Fields("DctoVV") = Val("" & rsg.Fields("DctoVV")) + Val("" & rsdd.Fields("DctoVV"))
                rsg.Fields("DctoPV") = Val("" & rsg.Fields("DctoPV")) + Val("" & rsdd.Fields("DctoPV"))
                rsg.Fields("TotalVVNeto") = Val("" & rsg.Fields("TotalVVNeto")) + Val("" & rsdd.Fields("TotalVVNeto"))
                rsg.Fields("TotalIGVNeto") = Val("" & rsg.Fields("TotalIGVNeto")) + Val("" & rsdd.Fields("TotalIGVNeto"))
                rsg.Fields("TotalPVNeto") = Val("" & rsg.Fields("TotalPVNeto")) + Val("" & rsdd.Fields("TotalPVNeto"))
            
                If RsD.RecordCount > 0 Then RsD.MoveFirst
                indExisteDocRef = False
                Do While Not RsD.EOF
                   If RsD.Fields("idDocumento") = strTipoDocImportado And RsD.Fields("idSerie") = "" & rsdd.Fields("idSerie") And RsD.Fields("idNumDOc") = "" & rsdd.Fields("idDocVentas") Then
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
                End If
                rsdd.MoveNext
                If rsdd.EOF Then Exit Do
            Loop
                   
        Loop
    End If
    
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
        
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
            
    gDetalle.Dataset.First
    Do While Not gDetalle.Dataset.EOF
        gDetalle.Dataset.Edit
        calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, gDetalle.Columns.ColumnByFieldName("VVUnit").Value, gDetalle.Columns.ColumnByFieldName("IGVUnit").Value, gDetalle.Columns.ColumnByFieldName("PVUnit").Value, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
        gDetalle.Dataset.Post
        gDetalle.Dataset.Next
    Loop
    
    calcularTotales
    indCargando = False
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub MostrarDocImportadoSinAgrupar(ByVal rscd As ADODB.Recordset, ByVal rsdd As ADODB.Recordset, ByVal strTipoDocImportado As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg                         As New ADODB.Recordset
Dim RsD                         As New ADODB.Recordset
Dim strSerieDocVentas           As String
Dim i                           As Integer
Dim indExisteDocRef             As Boolean
Dim cAuxIdProducto              As String

    indCargando = True
    i = 0
    strEstDocVentas = "GEN"
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "numOrdenCompra", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idProducto", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "CodigoRapido", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "idCodFabricante", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    rsg.Fields.Append "idMarca", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsMarca", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Cantidad2", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVBruto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVBruto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PorDcto", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "DctoVV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "DctoPV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idTipoProducto", adChar, 5, adFldIsNullable
    rsg.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idDocumentoImp", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "idDocVentasImp", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerieImp", adVarChar, 4, adFldIsNullable
    rsg.Fields.Append "NumLote", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "idUsuarioDcto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "VVUnitLista", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnitLista", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "VVUnitNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PVUnitNeto", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IdCentroCosto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdSucursalPres", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdDocumentoPres", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "IdSeriePres", adVarChar, 3, adFldIsNullable
    rsg.Fields.Append "IdDocVentasPres", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "FechaEmision", adVarChar, 10, adFldIsNullable
    rsg.Open
    
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adChar, 8, adFldIsNullable
    RsD.Open , , adOpenKeyset, adLockOptimistic
    
    If rscd.RecordCount = 0 Then
        txt_OrdenCompra.Text = ""
        txt_Partida.Text = ""
        txtCod_Almacen.Text = ""
        txtCod_FormaPago.Text = ""
        txtObs.Text = ""
        txtCod_Moneda.Text = ""
        txtCod_Vendedor.Text = ""
        dtp_Emision.Value = getFechaSistema
    Else
        txt_Partida.Text = "" & rscd.Fields("Partida")
        txtCod_Almacen.Text = "" & rscd.Fields("idAlmacen")
        txtObs.Text = "" & rscd.Fields("ObsDocVentas")
        txtCod_CentroCosto.Text = "" & rscd.Fields("IdCentroCosto")
    End If
        
    If rsdd.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idProducto") = ""
        rsg.Fields("CodigoRapido") = ""
        rsg.Fields("idCodFabricante") = ""
        rsg.Fields("GlsProducto") = ""
        rsg.Fields("idMarca") = ""
        rsg.Fields("GlsMarca") = ""
        rsg.Fields("idUM") = ""
        rsg.Fields("GlsUM") = ""
        rsg.Fields("Factor") = 1
        rsg.Fields("Afecto") = 1
        rsg.Fields("Cantidad") = 0
        rsg.Fields("Cantidad2") = 0
        rsg.Fields("VVUnit") = 0
        rsg.Fields("IGVUnit") = 0
        rsg.Fields("PVUnit") = 0
        rsg.Fields("TotalVVBruto") = 0
        rsg.Fields("TotalPVBruto") = 0
        rsg.Fields("PorDcto") = "0"
        rsg.Fields("DctoVV") = 0
        rsg.Fields("DctoPV") = 0
        rsg.Fields("TotalVVNeto") = 0
        rsg.Fields("TotalIGVNeto") = 0
        rsg.Fields("TotalPVNeto") = 0
        rsg.Fields("idTipoProducto") = ""
        rsg.Fields("idMoneda") = ""
        rsg.Fields("idDocumentoImp") = ""
        rsg.Fields("idDocVentasImp") = ""
        rsg.Fields("idSerieImp") = ""
        rsg.Fields("NumLote") = ""
        rsg.Fields("FecVencProd") = ""
        rsg.Fields("VVUnitLista") = 0
        rsg.Fields("PVUnitLista") = 0
        rsg.Fields("VVUnitNeto") = 0
        rsg.Fields("PVUnitNeto") = 0
        rsg.Fields("IdCentroCosto") = ""
        rsg.Fields("IdSucursalPres") = ""
        rsg.Fields("IdDocumentoPres") = ""
        rsg.Fields("IdSeriePres") = ""
        rsg.Fields("IdDocVentasPres") = ""
        rsg.Fields("FechaEmision") = getFechaSistema
        
    Else
        rsdd.MoveFirst
        rsdd.Sort = "idProducto"
        Do While Not rsdd.EOF
            rsg.AddNew
            i = i + 1
            rsg.Fields("Item") = i
            rsg.Fields("idProducto") = "" & rsdd.Fields("idProducto")
            rsg.Fields("CodigoRapido") = traerCampo("Productos", "CodigoRapido", "IdProducto", "" & rsdd.Fields("IdProducto"), True)
            rsg.Fields("idCodFabricante") = "" & rsdd.Fields("idCodFabricante")
            rsg.Fields("GlsProducto") = "" & rsdd.Fields("GlsProducto")
            rsg.Fields("idMarca") = "" & rsdd.Fields("idMarca")
            rsg.Fields("GlsMarca") = "" & rsdd.Fields("GlsMarca")
            rsg.Fields("idUM") = "" & rsdd.Fields("idUM")
            rsg.Fields("GlsUM") = "" & rsdd.Fields("GlsUM")
            rsg.Fields("Factor") = "" & rsdd.Fields("Factor")
            rsg.Fields("Afecto") = "" & rsdd.Fields("Afecto")
            rsg.Fields("idTipoProducto") = "" & rsdd.Fields("idTipoProducto")
            rsg.Fields("idMoneda") = "" & rsdd.Fields("idMoneda")
            rsg.Fields("idDocumentoImp") = strTipoDocImportado
            rsg.Fields("idDocVentasImp") = "" & rsdd.Fields("idDocVentas")
            rsg.Fields("idSerieImp") = "" & rsdd.Fields("idSerie")
            rsg.Fields("NumLote") = "" & rsdd.Fields("NumLote")
            rsg.Fields("FecVencProd") = "" & rsdd.Fields("FecVencProd")
            rsg.Fields("VVUnit") = Val("" & rsdd.Fields("VVUnit"))
            rsg.Fields("IGVUnit") = Val("" & rsdd.Fields("IGVUnit"))
            rsg.Fields("PVUnit") = Val("" & rsdd.Fields("PVUnit"))
            rsg.Fields("VVUnitLista") = Val("" & rsdd.Fields("VVUnitLista"))
            rsg.Fields("PVUnitLista") = Val("" & rsdd.Fields("PVUnitLista"))
            rsg.Fields("VVUnitNeto") = Val("" & rsdd.Fields("VVUnitNeto"))
            rsg.Fields("PVUnitNeto") = Val("" & rsdd.Fields("PVUnitNeto"))
            rsg.Fields("IdCentroCosto") = ""
            rsg.Fields("IdSucursalPres") = ""
            rsg.Fields("IdDocumentoPres") = ""
            rsg.Fields("IdSeriePres") = ""
            rsg.Fields("IdDocVentasPres") = ""
            rsg.Fields("FechaEmision") = getFechaSistema
            rsg.Fields("Cantidad") = Val("" & rsdd.Fields("Cantidad"))
            rsg.Fields("TotalVVBruto") = Val("" & rsdd.Fields("TotalVVBruto"))
            rsg.Fields("TotalPVBruto") = Val("" & rsdd.Fields("TotalPVBruto"))
            rsg.Fields("PorDcto") = Val("" & rsdd.Fields("PorDcto"))
            rsg.Fields("DctoVV") = Val("" & rsdd.Fields("DctoVV"))
            rsg.Fields("DctoPV") = Val("" & rsdd.Fields("DctoPV"))
            rsg.Fields("TotalVVNeto") = Val("" & rsdd.Fields("TotalVVNeto"))
            rsg.Fields("TotalIGVNeto") = Val("" & rsdd.Fields("TotalIGVNeto"))
            rsg.Fields("TotalPVNeto") = Val("" & rsdd.Fields("TotalPVNeto"))
            
            If RsD.RecordCount > 0 Then RsD.MoveFirst
            indExisteDocRef = False
            Do While Not RsD.EOF
               If RsD.Fields("idDocumento") = strTipoDocImportado And RsD.Fields("idSerie") = "" & rsdd.Fields("idSerie") And RsD.Fields("idNumDOc") = "" & rsdd.Fields("idDocVentas") Then
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
            End If
            rsdd.MoveNext
            If rsdd.EOF Then Exit Do
        Loop
    End If
    
    mostrarDatosGridSQL gDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
        
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
            
    gDetalle.Dataset.First
    Do While Not gDetalle.Dataset.EOF
        gDetalle.Dataset.Edit
        calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, gDetalle.Columns.ColumnByFieldName("VVUnit").Value, gDetalle.Columns.ColumnByFieldName("IGVUnit").Value, gDetalle.Columns.ColumnByFieldName("PVUnit").Value, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
        gDetalle.Dataset.Post
        gDetalle.Dataset.Next
    Loop
    
    calcularTotales
    indCargando = False
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub imprimeOCompra(ByRef StrMsgError As String)
Dim rsReporte       As New ADODB.Recordset
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim strTitulo       As String
Dim strReporte      As String
Dim strParamImpOC   As String
Dim strSP As String
On Error GoTo Err

    Screen.MousePointer = 11
    gStrRutaRpts = App.Path + "\Reportes\"
    
    strParamImpOC = Trim("" & traerCampo("parametros", "Valparametro", "glsparametro", "FORMATO_OC", True))
    
    strSP = "spu_ImpOC"
    
    If strParamImpOC = "1" Then
         If strTipoDoc = "94" Then
              strReporte = "rptImpOCompra.rpt"
              strTitulo = "Orden de Compra"
         ElseIf strTipoDoc = "87" Then
             strReporte = "rptImpRCompra.rpt"
             strTitulo = "Requerimiento de Compra"
         ElseIf strTipoDoc = "97" Then
            strReporte = "rptImpRInterno.rpt"
            strTitulo = "Requerimiento Interno"
            strSP = "spu_ImpReiCadillo"
         ElseIf strTipoDoc = "OS" Then
            strReporte = "rptImpOServicio.rpt"
            strTitulo = "Orden de Servicio"
            strSP = "spu_ImpOS"
         End If
     ElseIf strParamImpOC = "3" Then
         If strTipoDoc = "94" Then
              strReporte = "rptImpOCompra_PROVEEPERU.rpt"
              strTitulo = "Orden de Compra"
         ElseIf strTipoDoc = "87" Then
             strReporte = "rptImpRCompra.rpt"
             strTitulo = "Requerimiento de Compra"
         End If
     ElseIf strParamImpOC = "4" Then
         If strTipoDoc = "94" Then
              strReporte = "rptImpOCompra_Hac.rpt"
              strTitulo = "Orden de Compra"
         ElseIf strTipoDoc = "87" Then
             strReporte = "rptImpRCompra.rpt"
             strTitulo = "Requerimiento de Compra"
         End If
     ElseIf strParamImpOC = "5" Then
        If strTipoDoc = "94" Then
              strReporte = "rptImpOCompra_Solvet.rpt"
              strTitulo = "Orden de Compra"
              strSP = "spu_ImpOC_Solvet"
         ElseIf strTipoDoc = "87" Then
             strReporte = "rptImpRCompra.rpt"
             strTitulo = "Requerimiento de Compra"
         End If
     ElseIf strParamImpOC = "6" Then
        If strTipoDoc = "94" Then
              strReporte = "rptImpOCompra_MM.rpt"
              strTitulo = "Orden de Compra"
              strSP = "spu_ImpOCMM"
         ElseIf strTipoDoc = "87" Then
             strReporte = "rptImpRCompra.rpt"
             strTitulo = "Requerimiento de Compra"
         ElseIf strTipoDoc = "OS" Then
             strReporte = "rptImpOServicio_MM.rpt"
             strTitulo = "Orden de Servicio"
             strSP = "spu_ImpOSMM"
         End If
         
    ElseIf strParamImpOC = "7" Then
        If strTipoDoc = "94" Then
              strReporte = "rptImpOCompra_Indupol.rpt"
              strTitulo = "Orden de Compra"
              strSP = "spu_ImpOSIndupol"
        ElseIf strTipoDoc = "OS" Then
             strReporte = "rptImpOServicio_Indupol.rpt"
             strTitulo = "Orden de Servicio"
             strSP = "spu_ImpOSIndupol"
        End If
         
    '--- INMAC
    ElseIf strParamImpOC = "8" Then
        If strTipoDoc = "94" Then
              strReporte = "rptImpOCompra_Inmac.rpt"
              strTitulo = "Orden de Compra"
              strSP = "spu_ImpOC_Imnac"
        ElseIf strTipoDoc = "OS" Then
             strReporte = "rptImpOServicio_Inmac.rpt"
             strTitulo = "Orden de Servicio"
             strSP = "spu_ImpOS_Imnac"
        ElseIf strTipoDoc = "87" Then
             strReporte = "rptImpRCompra_Inmac.rpt"
             strTitulo = "Requerimiento de Compra"
             strSP = "spu_ImpRC_Inmac"
        End If
    '--- EITAL
    ElseIf strParamImpOC = "9" Then
        If strTipoDoc = "94" Then
             strReporte = "rptImpOCompra_Eital.rpt"
             strTitulo = "Orden de Compra"
             strSP = "spu_ImpOC_Eital"
        ElseIf strTipoDoc = "OS" Then
             strReporte = "rptImpOServicio.rpt"
             strTitulo = "Orden de Servicio"
             strSP = "spu_ImpOS"
        ElseIf strTipoDoc = "87" Then
             strReporte = "rptImpRCompra.rpt"
             strTitulo = "Requerimiento de Compra"
        End If
    
    ElseIf strParamImpOC = "10" Then
        
        If strTipoDoc = "94" Then
            strReporte = "rptImpOCompra_Proy3ctar.rpt"
            strTitulo = "Orden de Compra"
            strSP = "spu_ImpOC_Proy3ctar"
        ElseIf strTipoDoc = "OS" Then
            strReporte = "rptImpOServicio.rpt"
            strTitulo = "Orden de Servicio"
            strSP = "spu_ImpOS"
        ElseIf strTipoDoc = "87" Then
            strReporte = "rptImpRCompra.rpt"
            strTitulo = "Requerimiento de Compra"
        End If
    
    ElseIf strParamImpOC = "11" Then ' Paez
        If strTipoDoc = "94" Then
             strReporte = "rptImpOCompra_Paez.rpt"
             strTitulo = "Orden de Compra"
             strSP = "spu_ImpOC_Paez"
        ElseIf strTipoDoc = "87" Then
            strReporte = "rptImpRCompra.rpt"
            strTitulo = "Requerimiento de Compra"
        End If
    
    ElseIf strParamImpOC = "12" Then ' Salas
        If strTipoDoc = "94" Then
             strReporte = "rptImpOCompra_Salas.rpt"
             strTitulo = "Orden de Compra"
             strSP = "spu_ImpOC_Paez"
        ElseIf strTipoDoc = "87" Then
            strReporte = "rptImpRCompra.rpt"
            strTitulo = "Requerimiento de Compra"
        End If
        
    Else
        If strTipoDoc = "94" Then
             strReporte = "rptImpOCompra.rpt"
             strTitulo = "Orden de Compra"
             strSP = "spu_ImpOC"
        ElseIf strTipoDoc = "87" Then
            strReporte = "rptImpRCompra.rpt"
            strTitulo = "Requerimiento de Compra"
        End If
    End If
    
    If strReporte = "" Then Screen.MousePointer = 0: Exit Sub
    Set reporte = aplicacion.OpenReport(gStrRutaRpts & strReporte)
    Set rsReporte = DataProcedimiento(strSP, StrMsgError, glsEmpresa, glsSucursal, strTipoDoc, txt_Serie.Text, txt_NumDoc.Text)
    If StrMsgError <> "" Then GoTo Err
    If Not rsReporte.EOF And Not rsReporte.BOF Then
         reporte.database.SetDataSource rsReporte, 3
         vistaPrevia.CRViewer91.ReportSource = reporte
         vistaPrevia.Caption = strTitulo
         vistaPrevia.CRViewer91.ViewReport
         vistaPrevia.CRViewer91.DisplayGroupTree = False
         Screen.MousePointer = 0
         vistaPrevia.WindowState = 2
         vistaPrevia.Show
    Else
        Screen.MousePointer = 0
        MsgBox "No existen Registros  Seleccionados", vbInformation, App.Title
    End If
    Screen.MousePointer = 0
    Set rsReporte = Nothing
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
    
'''''''''''''    mostrarReporte strReporte, "varEmpresa|varSucursal|varTipoDoc|varSerie|varDocVentas", glsEmpresa & "|" & glsSucursal & "|" & strTipoDoc & "|" & txt_Serie.Text & "|" & txt_NumDoc.Text, "Situacion de la Orden de Compra", StrMsgError
'''''''''''''    If StrMsgError <> "" Then GoTo Err
    
'    Set reporte = aplicacion.OpenReport(gStrRutaRpts & strReporte)
'    Set rsReporte = DataProcedimiento(strSP, StrMsgError, glsEmpresa, glsSucursal, strTipoDoc, txt_serie.Text, txt_numdoc.Text)
'    If StrMsgError <> "" Then GoTo ERR
'
'    If Not rsReporte.EOF And Not rsReporte.BOF Then
'        reporte.Database.SetDataSource rsReporte, 3
'        vistaPrevia.CRViewer91.ReportSource = reporte
'        vistaPrevia.Caption = strTitulo
'        vistaPrevia.CRViewer91.ViewReport
'        vistaPrevia.CRViewer91.DisplayGroupTree = False
'        Screen.MousePointer = 0
'        vistaPrevia.WindowState = 2
'        vistaPrevia.Show
'    Else
'        Screen.MousePointer = 0
'        MsgBox "No existen Registros Seleccionados.", vbInformation, App.Title
'    End If
'
'    Screen.MousePointer = 0
'    Set rsReporte = Nothing
'    Set vistaPrevia = Nothing
'    Set aplicacion = Nothing
'    Set reporte = Nothing
    
    Exit Sub
    
Err:
    Screen.MousePointer = 0
    If StrMsgError = "" Then StrMsgError = Err.Description
    If rsReporte.State = 1 Then rsReporte.Close
    Set rsReporte = Nothing
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
End Sub

Private Sub calculaTotalesFilaPVNeto(dblCantidad As Double, dblVVUnit As Double, dblIGVUnit As Double, dblPVUnit As Double, strDcto As String, intAfecto As Integer)
Dim dblTotalVVBruto     As Double
Dim dblTotalPVBruto     As Double
Dim dblDctoVV           As Double
Dim dblDctoPV           As Double
Dim dblTotalVVNeto      As Double
Dim dblTotalIGVNeto     As Double
Dim dblTotalPVNeto      As Double
Dim dblDctoVVT          As Double
Dim dblDctoPVT          As Double
Dim strPorDcto()        As String
Dim strPorDctoRpt       As String
Dim i                   As Integer
    
    dblTotalPVNeto = Val(Format(gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value, "0.00"))
    dblTotalVVNeto = Val(Format((dblTotalPVNeto / (Val(1 & "." & right(dblIgvNEw, 2)))), "0.00"))
    
    If Trim(strDcto) <> "" And strDcto <> "0" Then
        strPorDcto = Split(strDcto, "+")
        dblDctoVV = 0
        strPorDctoRpt = ""
        For i = 0 To UBound(strPorDcto)
            dblDctoVVT = (dblVVUnit - dblDctoVV) * (Val(strPorDcto(i)) / 100)
            dblDctoPVT = (dblPVUnit - dblDctoPV) * (Val(strPorDcto(i)) / 100)

            dblDctoVV = dblDctoVV + dblDctoVVT
            dblDctoPV = dblDctoPV + dblDctoPVT
            strPorDctoRpt = strPorDctoRpt & CStr(Val(strPorDcto(i))) & "+"
        Next
    Else
        dblDctoVV = 0
        dblDctoPV = 0
        strPorDctoRpt = "0"
    End If
    If Len(strPorDctoRpt) > 1 Then strPorDctoRpt = left(strPorDctoRpt, Len(strPorDctoRpt) - 1)
    
    If intAfecto = 1 Then
        dblTotalIGVNeto = Val(Format((dblTotalVVNeto * dblIgvNEw), "0.00"))
    Else
        dblTotalIGVNeto = 0#
    End If
           
    dblTotalPVBruto = Val(Format(gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value, "0.00"))
    dblTotalVVBruto = Val(Format((dblTotalPVNeto / (Val(1 & "." & right(dblIgvNEw, 2)))), "0.00"))
    
    gDetalle.Columns.ColumnByFieldName("VVUnit").Value = Val(Format((dblTotalVVNeto / dblCantidad), "0.00000"))
    If intAfecto = 1 Then
        gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = Val(Format((Val(Format((dblTotalVVNeto / dblCantidad), "0.00")) * dblIgvNEw), "0.00"))
    Else
        dblTotalIGVNeto = 0#
    End If
    
    gDetalle.Columns.ColumnByFieldName("PVUnit").Value = Val(Format((dblTotalPVNeto / dblCantidad), "0.00000"))
    gDetalle.Columns.ColumnByFieldName("VVUnitNeto").Value = Val(Format((dblTotalVVNeto / dblCantidad), "0.00000"))
    gDetalle.Columns.ColumnByFieldName("PVUnitNeto").Value = Val(Format((dblTotalPVNeto / dblCantidad), "0.00000"))
    gDetalle.Columns.ColumnByFieldName("TotalVVBruto").Value = dblTotalVVBruto
    gDetalle.Columns.ColumnByFieldName("TotalPVBruto").Value = dblTotalPVBruto
    gDetalle.Columns.ColumnByFieldName("DctoVV").Value = dblDctoVV
    gDetalle.Columns.ColumnByFieldName("DctoPV").Value = dblDctoPV
    gDetalle.Columns.ColumnByFieldName("TotalVVNeto").Value = dblTotalVVNeto
    gDetalle.Columns.ColumnByFieldName("TotalIGVNeto").Value = dblTotalIGVNeto
    gDetalle.Columns.ColumnByFieldName("TotalPVNeto").Value = dblTotalPVNeto
    gDetalle.Columns.ColumnByFieldName("porDcto").Value = strPorDctoRpt
    
End Sub

Private Sub MostrarDocImportadoSinAgrupar2(ByVal rscd As ADODB.Recordset, ByVal rsdd As ADODB.Recordset, ByVal strTipoDocImportado As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg                    As New ADODB.Recordset
Dim RsD                    As New ADODB.Recordset
Dim strSerieDocVentas      As String
Dim dblTC                  As Double
Dim strCodFabri            As String
Dim strCodMar              As String
Dim strDesMar              As String
Dim intAfecto              As Integer
Dim strTipoProd            As String
Dim strMoneda              As String
Dim strCodUM               As String
Dim strDesUM               As String
Dim dblVVUnit              As Double
Dim dblIGVUnit             As Double
Dim dblPVUnit              As Double
Dim dblFactor              As Double
Dim intFila                As Integer
Dim i                      As Integer
Dim indExisteDocRef        As Boolean
Dim primero                As Boolean
        
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    RsD.Fields.Append "idNumDOc", adChar, 8, adFldIsNullable
    RsD.Open , , adOpenKeyset, adLockOptimistic
    
    If rscd.RecordCount = 1 Then
        txtObs.Text = "" & rscd.Fields("ObsDocVentas")
        txtCod_CentroCosto.Text = "" & rscd.Fields("IdCentroCosto")
    End If
    
    primero = True
    rsdd.MoveFirst
    Do While Not rsdd.EOF
        If primero = True Then
            primero = False
        Else
            gDetalle.Dataset.Insert
        End If
        
        If Trim("" & rsdd.Fields("idProducto")) = "" Then Exit Sub
        
        gDetalle.SetFocus
        gDetalle.Dataset.RecNo = intFila
        gDetalle.Dataset.Edit
        gDetalle.Columns.ColumnByFieldName("idProducto").Value = "" & rsdd.Fields("idProducto")
        gDetalle.Columns.ColumnByFieldName("CodigoRapido").Value = "" & rsdd.Fields("CodigoRapido")
        gDetalle.Columns.ColumnByFieldName("GlsProducto").Value = "" & rsdd.Fields("GlsProducto")
        gDetalle.Columns.ColumnByFieldName("idCodFabricante").Value = "" & rsdd.Fields("idCodFabricante")
        gDetalle.Columns.ColumnByFieldName("idMarca").Value = "" & rsdd.Fields("idMarca")
        gDetalle.Columns.ColumnByFieldName("GlsMarca").Value = "" & rsdd.Fields("GlsMarca")
        gDetalle.Columns.ColumnByFieldName("Afecto").Value = "" & rsdd.Fields("Afecto")
        gDetalle.Columns.ColumnByFieldName("idTipoProducto").Value = "" & rsdd.Fields("idTipoProducto")
        gDetalle.Columns.ColumnByFieldName("idMoneda").Value = "" & rsdd.Fields("idMoneda")
        gDetalle.Columns.ColumnByFieldName("idDocumentoImp").Value = strTipoDocImportado
        gDetalle.Columns.ColumnByFieldName("idDocVentasImp").Value = "" & rsdd.Fields("idDocVentas")
        gDetalle.Columns.ColumnByFieldName("idSerieImp").Value = "" & rsdd.Fields("idSerie")
        gDetalle.Columns.ColumnByFieldName("NumLote").Value = "" & rsdd.Fields("NumLote")
        gDetalle.Columns.ColumnByFieldName("FecVencProd").Value = "" & rsdd.Fields("FecVencProd")
        gDetalle.Columns.ColumnByFieldName("idUM").Value = "" & rsdd.Fields("idUM")
        gDetalle.Columns.ColumnByFieldName("GlsUM").Value = "" & rsdd.Fields("GlsUM")
        gDetalle.Columns.ColumnByFieldName("Factor").Value = "" & rsdd.Fields("Factor")
        gDetalle.Columns.ColumnByFieldName("Cantidad").Value = "" & rsdd.Fields("Cantidad")
    
        If DatosPrecio("" & rsdd.Fields("idProducto"), "" & rsdd.Fields("idTipoProducto"), "" & rsdd.Fields("idUM"), "" & rsdd.Fields("GlsUM"), dblVVUnit, "" & rsdd.Fields("Factor")) = False Then
        End If
        
        procesaMoneda strMoneda, txtCod_Moneda.Text, 0, dblVVUnit, intAfecto, dblVVUnit, dblIGVUnit, dblPVUnit
        gDetalle.Columns.ColumnByFieldName("VVUnit").Value = dblVVUnit
        gDetalle.Columns.ColumnByFieldName("IGVUnit").Value = dblIGVUnit
        gDetalle.Columns.ColumnByFieldName("PVUnit").Value = dblPVUnit
        gDetalle.Columns.ColumnByFieldName("VVUnitLista").Value = dblVVUnit
        gDetalle.Columns.ColumnByFieldName("PVUnitLista").Value = dblPVUnit
        gDetalle.Columns.ColumnByFieldName("PorDcto").Value = dblPorDsctoEspecial
        gDetalle.Columns.ColumnByFieldName("IdCentroCosto").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdSucursalPres").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdDocumentoPres").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdSeriePres").Value = ""
        gDetalle.Columns.ColumnByFieldName("IdDocVentasPres").Value = ""
        gDetalle.Columns.ColumnByFieldName("FechaEmision").Value = getFechaSistema
        gDetalle.Columns.ColumnByFieldName("GlsPlaca").Value = ""
            
        gDetalle.Dataset.Post
        gDetalle.Dataset.RecNo = intFila
        gDetalle.Dataset.Edit
    
        calculaTotalesFila gDetalle.Columns.ColumnByFieldName("Cantidad").Value, dblVVUnit, dblIGVUnit, dblPVUnit, gDetalle.Columns.ColumnByFieldName("PorDcto").Value, gDetalle.Columns.ColumnByFieldName("Afecto").Value
        gDetalle.Dataset.Post
        If "" & rsdd.Fields("idProducto") <> "" Then
            gDetalle.Columns.FocusedIndex = gDetalle.Columns.ColumnByFieldName("Cantidad").Index
        End If
        
        If RsD.RecordCount > 0 Then RsD.MoveFirst
        indExisteDocRef = False
        Do While Not RsD.EOF
           If RsD.Fields("idDocumento") = strTipoDocImportado And RsD.Fields("idSerie") = "" & rsdd.Fields("idSerie") And RsD.Fields("idNumDOc") = "" & rsdd.Fields("idDocVentas") Then
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
        End If
        rsdd.MoveNext
    Loop

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
End Sub

Private Sub Enviar_Correo(ByRef StrMsgError As String)
On Error GoTo Err
Dim StrNomDocProv           As String
Dim CorreoProv              As String
Dim CorreoResp              As String

    StrNomDocProv = ""
    CorreoProv = Trim("" & traerCampo("Personas", "mail", "idPersona", Trim(txtCod_Cliente.Text), False))
    CorreoResp = Trim("" & traerCampo("Personas", "mail", "idPersona", Trim(txtCod_Vendedor.Text), False))

    StrNomDocProv = "Orden de Compra N°" & txt_NumDoc.Text


    If Len(Trim("" & CorreoProv)) > 0 Then

        ExportarReporte "rptImpOCompra_Eital.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parDocVentas", glsEmpresa & "|" & glsSucursal & "|" & strTipoDoc & "|" & txt_Serie.Text & "|" & txt_NumDoc.Text, "Orden de Compra", StrNomDocProv, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    
        With MAPISession1
            .NewSession = False
            .SignOn
        End With

        With MAPIMessages1
            .SessionID = MAPISession1.SessionID
            .Compose ' CREAMOS EL MENSAJE
            .MsgSubject = "Orden de Compra" ' ASUNTO DEL MENSAJE
            .MsgNoteText = "Se Adjunta Orden de Compra N°" & txt_NumDoc.Text ' MENSAJE

            'XXXXXXCORREOSXXXXXX
            .RecipIndex = 0
            .RecipDisplayName = CorreoProv  ' Receptor
            .RecipType = mapToList

            If Len(Trim("" & CorreoResp)) > 0 Then
                .RecipIndex = 1
                .RecipDisplayName = CorreoResp ' Copia
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


    MsgBox ("El Proceso finalizó satisfactoriamente."), vbInformation, App.Title

    Exit Sub
    
Err:
    MAPISession1.SignOff
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub ValidaAprobacion(StrMsgError As String)
On Error GoTo Err
Dim strDocApro  As String

    'Si no esta aprobado no puede imprimir tampoco enviar correo
    If traerCampo("Parametros", "ValParametro", "GlsParametro", "VALIDA_APROBACION_BANDEJA", True) = "S" Then
        If strTipoDoc = "94" Or strTipoDoc = "87" Or strTipoDoc = "OS" Then
            strDocApro = traerCampo("Docventas", "indAprobado", "idDocumento", strTipoDoc, True, "iddocventas = '" & txt_NumDoc.Text & "' And idSerie = '" & txt_Serie.Text & "' And  idSucursal ='" & glsSucursal & "'")
            If strDocApro <> "1" Then
               Toolbar1.Buttons(7).Visible = False 'IMPRIMIR
               Toolbar1.Buttons(11).Visible = False 'ENVIAR CORREO
            End If
        End If
     Else
        Toolbar1.Buttons(11).Visible = False 'ENVIAR CORREO
     End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
