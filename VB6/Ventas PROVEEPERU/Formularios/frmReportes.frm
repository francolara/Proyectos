VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmReportes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   10260
   ClientLeft      =   5235
   ClientTop       =   2085
   ClientWidth     =   7230
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   7230
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   720
      Index           =   13
      Left            =   225
      TabIndex        =   84
      Top             =   8295
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaCliente 
         Height          =   315
         Left            =   6090
         Picture         =   "frmReportes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   240
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   285
         Left            =   1350
         TabIndex        =   85
         Tag             =   "TidPerCliente"
         Top             =   270
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
         Container       =   "frmReportes.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   285
         Left            =   2340
         TabIndex        =   87
         Tag             =   "TGlsCliente"
         Top             =   270
         Width           =   3690
         _ExtentX        =   6509
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
         Locked          =   -1  'True
         Container       =   "frmReportes.frx":03A6
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtdir 
         Height          =   285
         Left            =   6525
         TabIndex        =   102
         Tag             =   "TidPerCliente"
         Top             =   225
         Width           =   195
         _ExtentX        =   344
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
         Container       =   "frmReportes.frx":03C2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtruc 
         Height          =   285
         Left            =   6705
         TabIndex        =   103
         Tag             =   "TidPerCliente"
         Top             =   225
         Width           =   150
         _ExtentX        =   265
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
         Container       =   "frmReportes.frx":03DE
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label lbl_Cliente 
         Appearance      =   0  'Flat
         Caption         =   "Cliente:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   165
         TabIndex        =   88
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Caption         =   "Tipo:"
      ForeColor       =   &H00C00000&
      Height          =   630
      Index           =   14
      Left            =   240
      TabIndex        =   89
      Top             =   8865
      Visible         =   0   'False
      Width           =   6915
      Begin VB.OptionButton OptSerie 
         Caption         =   "Detallado por Serie"
         Height          =   240
         Left            =   4800
         TabIndex        =   97
         Top             =   255
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.OptionButton OptGeneral 
         Caption         =   "General"
         Height          =   240
         Left            =   855
         TabIndex        =   90
         Top             =   255
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton OptDetallado 
         Caption         =   "Detallado"
         Height          =   240
         Left            =   2880
         TabIndex        =   91
         Top             =   255
         Width           =   2025
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   12
      Left            =   225
      TabIndex        =   79
      Top             =   7635
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaChofer 
         Height          =   315
         Left            =   6075
         Picture         =   "frmReportes.frx":03FA
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   270
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Chofer 
         Height          =   285
         Left            =   1335
         TabIndex        =   82
         Tag             =   "TidPerChofer"
         Top             =   285
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
         Container       =   "frmReportes.frx":0784
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Chofer 
         Height          =   285
         Left            =   2310
         TabIndex        =   83
         Tag             =   "TglsChofer"
         Top             =   285
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":07A0
         Vacio           =   -1  'True
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         Caption         =   "Chofer:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   195
         TabIndex        =   80
         Top             =   375
         Width           =   1035
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   630
      Index           =   16
      Left            =   225
      TabIndex        =   98
      Top             =   9945
      Visible         =   0   'False
      Width           =   6915
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   855
         TabIndex        =   101
         Top             =   270
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton OptProvincia 
         Caption         =   "Provincias"
         Height          =   240
         Left            =   4860
         TabIndex        =   100
         Top             =   255
         Width           =   1440
      End
      Begin VB.OptionButton OptLima 
         Caption         =   "Lima y Callao"
         Height          =   240
         Left            =   2970
         TabIndex        =   99
         Top             =   255
         Width           =   1755
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   630
      Index           =   15
      Left            =   225
      TabIndex        =   92
      Top             =   9405
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaVendedor 
         Height          =   315
         Left            =   6075
         Picture         =   "frmReportes.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   195
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   285
         Left            =   1380
         TabIndex        =   94
         Tag             =   "TidPerCliente"
         Top             =   225
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
         Container       =   "frmReportes.frx":0B46
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   285
         Left            =   2325
         TabIndex        =   95
         Tag             =   "TGlsCliente"
         Top             =   225
         Width           =   3690
         _ExtentX        =   6509
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
         Locked          =   -1  'True
         Container       =   "frmReportes.frx":0B62
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Vendedor:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   195
         TabIndex        =   96
         Top             =   285
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   11
      Left            =   225
      TabIndex        =   74
      Top             =   7005
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaTipoDoc 
         Height          =   315
         Left            =   6060
         Picture         =   "frmReportes.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   300
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Documento 
         Height          =   285
         Left            =   1335
         TabIndex        =   76
         Tag             =   "TidMoneda"
         Top             =   315
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
         Container       =   "frmReportes.frx":0F08
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Documento 
         Height          =   285
         Left            =   2310
         TabIndex        =   77
         Top             =   315
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":0F24
         Vacio           =   -1  'True
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         Caption         =   "Documento:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   195
         TabIndex        =   78
         Top             =   375
         Width           =   1035
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   9
      Left            =   225
      TabIndex        =   69
      Top             =   5685
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaUsuario 
         Height          =   315
         Left            =   6060
         Picture         =   "frmReportes.frx":0F40
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   300
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Usuario 
         Height          =   285
         Left            =   1335
         TabIndex        =   71
         Tag             =   "TidMoneda"
         Top             =   315
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
         Container       =   "frmReportes.frx":12CA
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Usuario 
         Height          =   285
         Left            =   2310
         TabIndex        =   72
         Top             =   315
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":12E6
         Vacio           =   -1  'True
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         Caption         =   "Usuario:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   73
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   10
      Left            =   225
      TabIndex        =   63
      Top             =   6405
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaCaja 
         Height          =   315
         Left            =   6060
         Picture         =   "frmReportes.frx":1302
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   300
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Caja 
         Height          =   285
         Left            =   1335
         TabIndex        =   65
         Tag             =   "TidMoneda"
         Top             =   315
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
         Container       =   "frmReportes.frx":168C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Caja 
         Height          =   285
         Left            =   2310
         TabIndex        =   66
         Top             =   315
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":16A8
         Vacio           =   -1  'True
      End
      Begin VB.Label lblPrueba 
         Caption         =   "---"
         Height          =   135
         Left            =   5640
         TabIndex        =   68
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Caja:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   67
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   8
      Left            =   210
      TabIndex        =   58
      Top             =   5040
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaEmpresa 
         Height          =   315
         Left            =   6060
         Picture         =   "frmReportes.frx":16C4
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   300
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Empresa 
         Height          =   285
         Left            =   1335
         TabIndex        =   60
         Tag             =   "TidMoneda"
         Top             =   315
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
         Container       =   "frmReportes.frx":1A4E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Empresa 
         Height          =   285
         Left            =   2310
         TabIndex        =   61
         Top             =   315
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":1A6A
         Vacio           =   -1  'True
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "Empresa:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   62
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   1035
      Index           =   7
      Left            =   225
      TabIndex        =   52
      Top             =   4155
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CheckBox ChkAgrupado 
         Caption         =   "Agrupado por Vendedor de Campo"
         Height          =   195
         Left            =   3645
         TabIndex        =   57
         Top             =   720
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.CommandButton cmbAyudaMonedaGrupo 
         Height          =   315
         Left            =   6060
         Picture         =   "frmReportes.frx":1A86
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   270
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_MonedaGrupo 
         Height          =   285
         Left            =   1350
         TabIndex        =   54
         Tag             =   "TidMoneda"
         Top             =   300
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
         Container       =   "frmReportes.frx":1E10
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_MonedaGrupo 
         Height          =   285
         Left            =   2325
         TabIndex        =   55
         Top             =   300
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":1E2C
         Vacio           =   -1  'True
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Moneda:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   225
         TabIndex        =   56
         Top             =   375
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Caption         =   " Tipo de Producto "
      ForeColor       =   &H00000000&
      Height          =   630
      Index           =   6
      Left            =   240
      TabIndex        =   31
      Top             =   3570
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaNivel 
         Height          =   315
         Index           =   0
         Left            =   6090
         Picture         =   "frmReportes.frx":1E48
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaNivel 
         Height          =   315
         Index           =   1
         Left            =   6090
         Picture         =   "frmReportes.frx":21D2
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   600
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaNivel 
         Height          =   315
         Index           =   2
         Left            =   6090
         Picture         =   "frmReportes.frx":255C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   960
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaNivel 
         Height          =   315
         Index           =   3
         Left            =   6090
         Picture         =   "frmReportes.frx":28E6
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1320
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaNivel 
         Height          =   315
         Index           =   4
         Left            =   6090
         Picture         =   "frmReportes.frx":2C70
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1680
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Nivel 
         Height          =   285
         Index           =   0
         Left            =   1365
         TabIndex        =   37
         Tag             =   "TidNivelPred"
         Top             =   270
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
         Container       =   "frmReportes.frx":2FFA
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nivel 
         Height          =   285
         Index           =   0
         Left            =   2340
         TabIndex        =   38
         Top             =   270
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":3016
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Nivel 
         Height          =   285
         Index           =   1
         Left            =   1365
         TabIndex        =   39
         Tag             =   "TidNivelPred"
         Top             =   630
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
         Container       =   "frmReportes.frx":3032
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nivel 
         Height          =   285
         Index           =   1
         Left            =   2340
         TabIndex        =   40
         Top             =   630
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":304E
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Nivel 
         Height          =   285
         Index           =   2
         Left            =   1365
         TabIndex        =   41
         Tag             =   "TidNivelPred"
         Top             =   990
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
         Container       =   "frmReportes.frx":306A
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nivel 
         Height          =   285
         Index           =   2
         Left            =   2340
         TabIndex        =   42
         Top             =   990
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":3086
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Nivel 
         Height          =   285
         Index           =   3
         Left            =   1365
         TabIndex        =   43
         Tag             =   "TidNivelPred"
         Top             =   1350
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
         Container       =   "frmReportes.frx":30A2
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nivel 
         Height          =   285
         Index           =   3
         Left            =   2340
         TabIndex        =   44
         Top             =   1350
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":30BE
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Nivel 
         Height          =   285
         Index           =   4
         Left            =   1365
         TabIndex        =   45
         Tag             =   "TidNivelPred"
         Top             =   1710
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
         Container       =   "frmReportes.frx":30DA
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nivel 
         Height          =   285
         Index           =   4
         Left            =   2340
         TabIndex        =   46
         Top             =   1710
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":30F6
         Vacio           =   -1  'True
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nivel:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   51
         Top             =   285
         Width           =   405
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nivel:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   50
         Top             =   645
         Width           =   405
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nivel:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   49
         Top             =   1005
         Width           =   405
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nivel:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   48
         Top             =   1365
         Width           =   405
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nivel:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   47
         Top             =   1725
         Width           =   405
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   5
      Left            =   225
      TabIndex        =   14
      Top             =   2925
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   6060
         Picture         =   "frmReportes.frx":3112
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   270
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   285
         Left            =   1350
         TabIndex        =   16
         Tag             =   "TidMoneda"
         Top             =   300
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
         Container       =   "frmReportes.frx":349C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   285
         Left            =   2325
         TabIndex        =   17
         Top             =   300
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":34B8
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         Caption         =   "Moneda:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   225
         TabIndex        =   18
         Top             =   375
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   4
      Left            =   225
      TabIndex        =   13
      Top             =   2325
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6030
         Picture         =   "frmReportes.frx":34D4
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   300
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   285
         Left            =   1335
         TabIndex        =   24
         Tag             =   "TidMoneda"
         Top             =   315
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
         Container       =   "frmReportes.frx":385E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   285
         Left            =   2310
         TabIndex        =   25
         Top             =   315
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":387A
         Vacio           =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Sucursal:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Caption         =   " Fecha "
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   3
      Left            =   225
      TabIndex        =   7
      Top             =   1725
      Visible         =   0   'False
      Width           =   6915
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107610113
         CurrentDate     =   38667
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "al"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2160
         TabIndex        =   12
         Top             =   375
         Width           =   165
      End
   End
   Begin VB.Frame fraBotones 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   240
      TabIndex        =   3
      Top             =   10155
      Width           =   6915
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   2
      Left            =   225
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaProducto 
         Height          =   315
         Left            =   6060
         Picture         =   "frmReportes.frx":3896
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   300
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   285
         Left            =   1335
         TabIndex        =   28
         Tag             =   "TidMoneda"
         Top             =   315
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
         Container       =   "frmReportes.frx":3C20
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   285
         Left            =   2310
         TabIndex        =   29
         Top             =   315
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":3C3C
         Vacio           =   -1  'True
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Producto:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   30
         Top             =   390
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Caption         =   " Rango de Fechas "
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   555
      Visible         =   0   'False
      Width           =   6915
      Begin MSComCtl2.DTPicker dtpfInicio 
         Height          =   315
         Left            =   1515
         TabIndex        =   6
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107610113
         CurrentDate     =   38667
      End
      Begin MSComCtl2.DTPicker dtpFFinal 
         Height          =   315
         Left            =   4515
         TabIndex        =   8
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107610113
         CurrentDate     =   38667
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3960
         TabIndex        =   11
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   900
         TabIndex        =   10
         Top             =   375
         Width           =   465
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton cmbAyudaAlmacen 
         Height          =   315
         Left            =   6045
         Picture         =   "frmReportes.frx":3C58
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   290
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Almacen 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Tag             =   "TidAlmacen"
         Top             =   300
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
         Container       =   "frmReportes.frx":3FE2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Almacen 
         Height          =   285
         Left            =   2295
         TabIndex        =   21
         Top             =   300
         Width           =   3690
         _ExtentX        =   6509
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
         Container       =   "frmReportes.frx":3FFE
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Almacen 
         Appearance      =   0  'Flat
         Caption         =   "Almacen:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   225
         TabIndex        =   22
         Top             =   375
         Width           =   765
      End
   End
   Begin CATControls.CATTextBox txtCodtienda 
      Height          =   285
      Left            =   0
      TabIndex        =   104
      Tag             =   "TidPerCliente"
      Top             =   0
      Width           =   150
      _ExtentX        =   265
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
      Container       =   "frmReportes.frx":401A
      Estilo          =   1
      EnterTab        =   -1  'True
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NumerosFrames As String
Public GlsReporte As String
Public GlsForm As String
Public IndAgrupado

Private Sub cmbAyudaAlmacen_Click()
Dim strCondicion As String

    If fraReportes(4).Visible Then

        If txtCod_Sucursal.Text = "" Then
            MsgBox "Seleccione una sucursal", vbInformation, App.Title
            txtCod_Sucursal.SetFocus
            Exit Sub
        End If
        strCondicion = " AND idSucursal = '" & txtCod_Sucursal.Text & "'"
    End If
    mostrarAyuda "ALMACENVTA", txtCod_Almacen, txtGls_Almacen, strCondicion

End Sub

Private Sub cmbAyudaCaja_Click()
    
    mostrarAyuda "CAJASUSUARIOFILTRO", txtCod_Caja, txtGls_Caja, "AND u.idUsuario = '" & txtCod_Usuario.Text & "'"

End Sub

Private Sub cmbAyudaChofer_Click()
    
    mostrarAyuda "CHOFER", txtCod_Chofer, txtGls_Chofer

End Sub

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyudaClientes txtCod_Cliente, txtGls_Cliente, txtruc, txtdir, txtCodtienda
    
End Sub

Private Sub cmbAyudaEmpresa_Click()
    
    mostrarAyuda "EMPRESA", txtCod_Empresa, txtGls_Empresa

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaMonedaGrupo_Click()
    
    mostrarAyuda "MONEDA", txtCod_MonedaGrupo, txtGls_MonedaGrupo
    If txtCod_MonedaGrupo.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub CmbAyudaNivel_Click(Index As Integer)
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

Private Sub cmbAyudaProducto_Click()
    
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmbAyudaTipoDoc_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento

End Sub

Private Sub cmbAyudaUsuario_Click()
    
    mostrarAyuda "USUARIO", txtCod_Usuario, txtGls_Usuario

End Sub

Private Sub cmbAyudaVendedor_Click()
    
    If GlsReporte = "rptVentasPorVendedorCampo.rpt" Then
        mostrarAyuda "VENDEDORCAMPO", txtCod_Vendedor, txtGls_Vendedor
    Else
        mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor
    End If

End Sub

Private Sub Command1_Click()
On Error GoTo Err
Dim rsReporte       As New ADODB.Recordset
Dim rsSubReporte    As New ADODB.Recordset
Dim fIni            As String, Ffin As String
Dim strIni          As String, strFin As String
Dim Dpto            As String
Dim fCorte          As String
Dim strCorte        As String
Dim fSaldoIni       As String
Dim strMoneda       As String
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim subReporte      As CRAXDRT.Report
Dim rsTemp          As New ADODB.Recordset
Dim strAlmacen      As String
Dim strSQL          As String
Dim StrMsgError     As String
Dim strCodSucursal  As String
Dim strCodEmpresa   As String
Dim strCodProducto  As String
Dim csucursal       As String
                    
    Screen.MousePointer = 11
    If fraReportes(5).Visible Then
        If Trim(txtCod_Moneda.Text) = "" Then
            Screen.MousePointer = 0
            MsgBox "Seleccione una Moneda", vbInformation, App.Title
            txtCod_Moneda.SetFocus
            Exit Sub
        End If
    End If
    
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strIni = Format(dtpfInicio.Value, "dd/mm/yyyy")
    strFin = Format(dtpFFinal.Value, "dd/mm/yyyy")
    
    fCorte = Format(DtpFecha.Value, "yyyy-mm-dd")
    strCorte = Format(DtpFecha.Value, "dd/mm/yyyy")
    
    strMoneda = txtCod_Moneda.Text
    strCodSucursal = Trim(txtCod_Sucursal.Text)
    strCodEmpresa = Trim(txtCod_Empresa.Text)
    strCodProducto = Trim(txtCod_Producto.Text)
    
    If fraReportes(9).Visible Then
        csql = "SELECT m.idMovCaja,m.IdSucursal " & _
                "FROM movcajas m " & _
                "WHERE m.idEmpresa = '" & glsEmpresa & "' " & _
                 "AND m.idUsuario = '" & txtCod_Usuario.Text & "' " & _
                 "AND m.idCaja = '" & txtCod_Caja.Text & "' " & _
                 "AND DATE_FORMAT(m.FecCaja ,'%d/%m/%Y') = DATE_FORMAT('" & fCorte & "','%d/%m/%Y')"
                 
        rsTemp.Open csql, Cn, adOpenForwardOnly, adLockOptimistic
        If Not rsTemp.EOF Then
            strMovCaja = "" & rsTemp.Fields("idMovCaja")
            csucursal = "" & rsTemp.Fields("IdSucursal")
        Else
            StrMsgError = "No hay caja disponible para la fecha indicada"
            GoTo Err
        End If
        
        If rsTemp.State = 1 Then rsTemp.Close
        Set rsTemp = Nothing
        
    End If
    
    Select Case GlsReporte
        Case Is = "rptKardexProducto.rpt"
            If Trim(txtCod_Almacen.Text) = "" Then
                Screen.MousePointer = 0
                MsgBox "Seleccione un Almacen", vbInformation, App.Title
                txtCod_Almacen.SetFocus
                Exit Sub
            End If
                
            If Trim(txtCod_Producto.Text) = "" Then
                Screen.MousePointer = 0
                MsgBox "Seleccione un Producto", vbInformation, App.Title
                txtCod_Producto.SetFocus
                Exit Sub
            End If
            
            txtCod_Moneda.Text = "PEN"
            Set rsReporte = kardex(StrMsgError, Trim(txtCod_Producto.Text))
            If StrMsgError <> "" Then GoTo Err
            Set reporte = aplicacion.OpenReport(gStrRutaRpts & GlsReporte)
        
        Case Is = "rptKardex.rpt"
            If Trim(txtCod_Almacen.Text) = "" Then
                Screen.MousePointer = 0
                MsgBox "Seleccione un Almacen", vbInformation, App.Title
                txtCod_Almacen.SetFocus
                Exit Sub
            End If
            
            Set rsReporte = kardex(StrMsgError, Trim(txtCod_Producto.Text))
            If StrMsgError <> "" Then GoTo Err
            Set reporte = aplicacion.OpenReport(gStrRutaRpts & GlsReporte)
        
        Case Is = "rptVentasPorCliente.rpt"
            If OptGeneral.Value = True Then
                mostrarReporte "rptVentasPorClienteGeneral.rpt", "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta|parCliente", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|" & Trim(txtCod_Cliente.Text), GlsForm, StrMsgError
            Else
                mostrarReporte "rptVentasPorCliente.rpt", "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta|parCliente", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|" & Trim(txtCod_Cliente.Text), GlsForm, StrMsgError
            End If
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptVentasPorClienteDctoEspecial.rpt"
            mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptVentasPorVendedor.rpt", "rptVentasPorVendedorCampo.rpt"
            If OptTodos.Value = True Then
                Dpto = "0"
            ElseIf OptLima.Value = True Then
                Dpto = "1"
            ElseIf OptProvincia.Value = True Then
                Dpto = "2"
            End If
            
            If GlsReporte = "rptVentasPorVendedorCampo.rpt" Then
                If OptGeneral.Value = True Then
                    mostrarReporte "rptVentasPorVendedorCampoGeneral.rpt", "parEmpresa|parSucursal|parVendedor|parMoneda|parDpto|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & Dpto & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
                Else
                    mostrarReporte GlsReporte, "parEmpresa|parSucursal|parVendedor|parMoneda|parDpto|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & Dpto & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
                End If
            Else
                mostrarReporte GlsReporte, "parEmpresa|parSucursal|parVendedor|parMoneda|parDpto|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & Dpto & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            End If
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptVentasPorVendedorPorTipoDoc.rpt"
            mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptVentasPorProducto.rpt"
            mostrarReporte left(GlsReporte, Len(GlsReporte) - 4) & Format(glsNumNiveles, "00") & ".rpt", "parEmpresa|parSucursal|parMoneda|parFechaIni|parFechaFin|parProducto", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|" & strCodProducto, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub

        Case Is = "rptVentasPorVendedorPorTipoProd.rpt"
            mostrarReporte left(GlsReporte, Len(GlsReporte) - 4) & Format(glsNumNiveles, "00") & ".rpt", "parEmpresa|parSucursal|parMoneda|parNivel|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & txtCod_Nivel(glsNumNiveles - 1).Text & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptVentasGrupoProducto.rpt"
            ChkAgrupado.Visible = True
            If ChkAgrupado.Value = 1 Then
                strMoneda = txtCod_MonedaGrupo.Text
                mostrarReporte "rptVentasGrupoProductoVendedor.rpt", "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            Else
                mostrarReporte left(GlsReporte, Len(GlsReporte) - 4) & ".rpt", "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            Exit Sub
        
        Case Is = "rptInventario.rpt"
            mostrarReporte left(GlsReporte, Len(GlsReporte) - 4) & Format(glsNumNiveles, "00") & ".rpt", "parEmpresa|parSucursal|parAlmacen|parFecha", glsEmpresa & "|" & strCodSucursal & "|" & txtCod_Almacen.Text & "|" & fCorte, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptInventarioValorizado.rpt"
            mostrarReporte left(GlsReporte, Len(GlsReporte) - 4) & Format(glsNumNiveles, "00") & ".rpt", "parEmpresa|parSucursal|parAlmacen|parMoneda|parFecha", glsEmpresa & "|" & strCodSucursal & "|" & txtCod_Almacen.Text & "|" & strMoneda & "|" & fCorte, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptStockVentasRatios.rpt"
            mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptVentasGrupo.rpt"
            mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptLiquidacionCajaDet.rpt"
            If Trim("" & traerCampo("parametros", "ValParametro", "GlsParametro", "FORMATO_LIQUIDACION", True)) = "2" Then
                mostrarReporte "rptLiquidacionCajaDet_formato2.rpt", "parEmpresa|parSucursal|parMovCaja", glsEmpresa & "|" & csucursal & "|" & strMovCaja, GlsForm, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            Else
                mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMovCaja", glsEmpresa & "|" & csucursal & "|" & strMovCaja, GlsForm, StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
            Exit Sub
            
        Case Is = "rptVentasClientesNiveles.rpt"
            mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptRecepMercaderiaPendiente.rpt"
            mostrarReporte GlsReporte, "parEmpresa|parSucursal", glsEmpresa & "|" & strCodSucursal, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
    
        Case Is = "rptVentasSucursalHoras.rpt"
            mostrarReporte GlsReporte, "parEmpresa|parMoneda|parFechaIni|parFechaFin", glsEmpresa & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptComprasMensuales.rpt"
            mostrarReporte GlsReporte, "parEmpresa|parSucursal|parAlmacen|parFechaHasta|parMoneda", glsEmpresa & "|" & strCodSucursal & "|" & Trim(txtCod_Almacen.Text) & "|" & fCorte & "|" & strMoneda, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
    
        Case Is = "rptDocumentos.rpt"
            mostrarReporte GlsReporte, "parSucursal|parFecDesde|parFecHasta|parDocumento", strCodSucursal & "|" & fIni & "|" & Ffin & "|" & Trim(txtCod_Documento.Text), GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptDespachoXChofer.rpt"
            mostrarReporte GlsReporte, "parSucursal|parChofer|parDesde|parHasta|parEmpresa", Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Chofer.Text) & "|" & fIni & "|" & Ffin & "|" & glsEmpresa, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptClientesXVendedor.rpt"
            If txtCod_Vendedor.Text = "" Then
                txtCod_Vendedor.Text = " "
            End If
            mostrarReporte GlsReporte, "parVendedor|parEmpresa", Trim(txtCod_Vendedor.Text) & "|" & glsEmpresa, GlsForm, StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Case Is = "rptClientesXNivelGeneral.rpt"
            If OptGeneral.Value = True Then
                mostrarReporte GlsReporte, "parEmpresa|parSucursal|parCliente|parMoneda|parFecDesde|parFecHasta", Trim(glsEmpresa) & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Cliente.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            ElseIf OptDetallado.Value = True Then
                mostrarReporte "rptClientesXNivelDetallado.rpt", "parEmpresa|parSucursal|parCliente|parDocumento|parMoneda|parFecDesde|parFecHasta", Trim(glsEmpresa) & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Cliente.Text) & "|" & Trim(txtCod_Documento.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            ElseIf OptSerie.Value = True Then
                mostrarReporte "rptNivelesXSerie.rpt", "parEmpresa|parSucursal|parCliente|parMoneda|parFecDesde|parFecHasta", Trim(glsEmpresa) & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Cliente.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            End If
            
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
                
    End Select
    
    If Not rsReporte.EOF And Not rsReporte.BOF Then
        reporte.Database.SetDataSource rsReporte, 3
        vistaPrevia.CRViewer91.ReportSource = reporte
        vistaPrevia.Caption = GlsForm
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
    Set rsSubReporte = Nothing
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
    Set subReporte = Nothing
    
    Exit Sub
    
Err:
    Screen.MousePointer = 0
    If StrMsgError = "" Then StrMsgError = Err.Description
    Set rsReporte = Nothing
    Set rsSubReporte = Nothing
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
    Set subReporte = Nothing
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Command2_Click()
    
    Unload Me

End Sub

Private Sub mostrarNiveles(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsj As New ADODB.Recordset
Dim i As Integer
Dim numPesos As Integer

    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    numPesos = Val("" & rsj.RecordCount)
    fraReportes(6).Height = 520 * numPesos
    i = 0
    Do While Not rsj.EOF
        lblNivel(i).Caption = "" & rsj.Fields("GlsTipoNivel")
        rsj.MoveNext
        i = i + 1
    Loop
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    
    Exit Sub

Err:
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
Dim strNumeros() As String
Dim intTop  As Integer
Dim intForm As Integer
Dim indTipoProd As Boolean
Dim i As Integer

    txtruc.Visible = False
    txtdir.Visible = False
    strNumeros = Split(NumerosFrames, ",")
    intForm = 1
    intTop = 75
    indTipoProd = False
    
    OptGeneral.left = 1425
    OptDetallado.left = 3840
    OptSerie.Visible = False
    
    If GlsReporte = "rptClientesXNivelGeneral.rpt" Then
        fraReportes(11).Enabled = False
        OptGeneral.left = 855
        OptDetallado.left = 2895
        OptSerie.Visible = True
    End If

    For i = 0 To UBound(strNumeros)
        fraReportes(strNumeros(i)).top = intTop
        fraReportes(strNumeros(i)).Visible = True
        
        If strNumeros(i) = 6 Then
            indTipoProd = True
            mostrarNiveles StrMsgError
        Else
            intForm = intForm + 1
            intTop = intTop + 825
        End If
    Next

    If GlsReporte = "rptVentasGrupoProducto.rpt" Then
        ChkAgrupado.Visible = True
        txtCod_MonedaGrupo.Text = glsMonVentas
    Else
        ChkAgrupado.Visible = False
    End If
    
    If indTipoProd Then
        intTop = intTop + fraReportes(6).Height
        Me.Height = intForm * 1100 + fraReportes(6).Height
    Else
        Me.Height = intForm * 1000
    End If

    fraBotones.top = intTop
    Me.Caption = GlsForm

    Me.top = frmPrincipal.Height / 5
    Me.left = frmPrincipal.Width / 5

    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Empresa.Text = "TODAS LAS EMPRESAS"
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    txtCod_Moneda.Text = glsMonVentas
    
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    DtpFecha.Value = Format(Date, "dd/mm/yyyy")
    ChkAgrupado.Visible = True

End Sub

Private Sub OptDetallado_Click()
    
    If GlsReporte = "rptClientesXNivelGeneral.rpt" Then
        fraReportes(11).Enabled = True
    End If

End Sub

Private Sub OptGeneral_Click()
    
    If GlsReporte = "rptClientesXNivelGeneral.rpt" Then
        fraReportes(11).Enabled = False
    End If

End Sub

Private Sub OptSerie_Click()
    
    If GlsReporte = "rptClientesXNivelGeneral.rpt" Then
        fraReportes(11).Enabled = True
    End If

End Sub

Private Sub txtCod_Almacen_Change()
    
    If txtCod_Almacen.Text <> "" Then
        txtGls_Almacen.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", txtCod_Almacen.Text, True)
    End If

End Sub

Private Sub txtCod_Almacen_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Almacen.Text = ""
    End If

End Sub

Private Sub txtCod_Almacen_KeyPress(KeyAscii As Integer)
Dim strCondicion As String

    If KeyAscii <> 13 Then
        If fraReportes(4).Visible Then
            If txtCod_Sucursal.Text = "" Then
                MsgBox "Seleccione una sucursal", vbInformation, App.Title
                txtCod_Sucursal.SetFocus
                KeyAscii = 0
                Exit Sub
            End If
            strCondicion = " AND idSucursal = '" & txtCod_Sucursal.Text & "'"
        End If
        mostrarAyudaKeyascii KeyAscii, "ALMACENVTA", txtCod_Almacen, txtGls_Almacen, strCondicion
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Caja_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CAJASUSUARIOFILTRO", txtCod_Caja, txtGls_Caja, "AND u.idUsuario = '" & txtCod_Usuario.Text & "'"
        KeyAscii = 0
        If txtCod_Caja.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtCod_Chofer_Change()
    
    If indCargando = False And txtCod_Chofer.Text <> "" Then
        txtGls_Chofer.Text = traerCampo("personas", "glsPersona", "idPersona", Trim(txtCod_Chofer.Text), False)
    End If

End Sub

Private Sub txtCod_Cliente_Change()

    If txtCod_Cliente.Text <> "" Then
        If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
            If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
                txtGls_Cliente.Text = traerCampo("personas p Inner Join  clientes c On  p.idPersona = c.idCliente Inner Join personas v On c.idVendedorCampo = v.idPersona", "p.GlsPersona", "p.idPersona", txtCod_Cliente.Text, False, "c.idVendedorCampo ='" & glsUser & "' ")
            Else
                txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
            End If
        Else
            txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
        End If
    Else
        txtGls_Cliente.Text = ""
    End If

End Sub

Private Sub txtCod_Documento_Change()
    
    txtGls_Documento.Text = traerCampo("Documentos", "GlsDocumento", "idDocumento", txtCod_Documento.Text, False)

End Sub

Private Sub txtCod_Documento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DOCUMENTOS", txtCod_Documento, txtGls_Documento
        KeyAscii = 0
        If txtCod_Caja.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtCod_Empresa_Change()
    
    If txtCod_Empresa.Text <> "" Then
        txtGls_Empresa.Text = traerCampo("empresas", "GlsEmpresa", "idEmpresa", txtCod_Empresa.Text, False)
    Else
        txtGls_Empresa.Text = "TODAS LAS EMPRESAS"
    End If

End Sub

Private Sub txtCod_Empresa_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Empresa.Text = ""
    End If

End Sub

Private Sub txtCod_Empresa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "EMPRESA", txtCod_Empresa, txtGls_Empresa
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MONEDA", txtCod_Moneda, txtGls_Moneda
        KeyAscii = 0
        If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_MonedaGrupo_Change()
    
    txtGls_MonedaGrupo.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_MonedaGrupo.Text, False)

End Sub

Private Sub txtCod_Nivel_Change(Index As Integer)
    
    txtGls_Nivel(Index).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_Nivel(Index).Text, True)

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

Private Sub txtCod_Producto_Change()
    
    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", "idProducto", txtCod_Producto.Text, True)
    Else
        txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    End If

End Sub

Private Sub txtCod_Producto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Producto.Text = ""
    End If

End Sub

Private Sub txtCod_Producto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "PRODUCTOS", txtCod_Producto, txtGls_Producto
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Sucursal_Change()

    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    
    Me.Caption = Me.Caption & " - " & txtGls_Sucursal.Text
    If fraReportes(0).Visible Then
        txtCod_Almacen.Text = ""
    End If
    
End Sub

Private Sub txtCod_Sucursal_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Sucursal.Text = ""
    End If

End Sub

Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        KeyAscii = 0
    End If

End Sub

Private Function kardex(ByRef StrMsgError As String, Optional codproducto As String) As Recordset
On Error GoTo Err
Dim rsTemp As New ADODB.Recordset
Dim rsd As New ADODB.Recordset
Dim strSQL As String
Dim i As Integer
Dim dblsaldo As Double
Dim dblValSaldo As Double
Dim CodProd As String
Dim CodProdAnt As String
Dim strFecIni As String
Dim strFecFin As String
Dim strSimboloMoneda As String

    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strSimboloMoneda = txtGls_Moneda.Text
    
    strSQL = "SELECT vc.idValesCab," & _
                   "vc.tipoVale," & _
                   "vc.fechaEmision," & _
                   "vd.IdProducto,vd.GlsProducto," & _
                   "vc.IdAlmacen," & _
                   "vc.idProvCliente,/*pe.ruc,pe.GlsPersona,*/ (select ruc from personas where idPersona = vc.idProvCliente) as ruc, (select GlsPersona from personas where idPersona = vc.idProvCliente) as GlsPersona, " & _
                   "vc.GlsDocReferencia," & _
                   "um.abreUM," & _
                   "vd.Cantidad, " & _
                   "CASE '" & txtCod_Moneda.Text & "' WHEN 'PEN' THEN  IF(vc.idMoneda = 'PEN', vd.TotalVVNeto,vd.TotalVVNeto * TipoCambio)" & _
                                           "WHEN 'USD' THEN  IF(vc.idMoneda = 'USD', vd.TotalVVNeto,vd.TotalVVNeto / TipoCambio)" & _
                   "END as VVUnit "
                   
    strSQL = strSQL & "FROM valescab vc,valesdet vd,productos pr,unidadmedida um " & _
            "WHERE vc.idValesCab = vd.idValesCab " & _
              "AND vc.idEmpresa = vd.idEmpresa " & _
              "AND vc.idSucursal = vd.idSucursal " & _
              "AND vd.idProducto = pr.idProducto " & _
              "AND vd.idEmpresa = pr.idEmpresa " & _
              "AND pr.idUMCompra = um.idUM " & _
              "AND pr.estProducto = 'A' " & _
              "AND vc.idEmpresa = '" & glsEmpresa & "'" & _
              "AND vc.idSucursal = '" & glsSucursal & "'" & _
              "AND pr.idEmpresa = '" & glsEmpresa & "'" & _
              "AND vc.idAlmacen = '" & txtCod_Almacen.Text & "'" & _
              "AND vc.idPeriodoInv = '" & glsCodPeriodoINV & "' " & _
              "AND vc.estValeCab <> 'ANU' " & _
              "AND vc.FechaEmision BETWEEN '" & strFecIni & "' AND '" & strFecFin & "' "
              
    If codproducto <> "" Then
        strSQL = strSQL & " AND vd.idProducto = '" & codproducto & "'"
    End If
              
    strSQL = strSQL & " ORDER BY vd.IdProducto,vc.FechaEmision"
            
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.Open strSQL, Cn, adOpenKeyset, adLockOptimistic
    
    Set rsTemp.ActiveConnection = Nothing
    
    rsd.Fields.Append "Item", adInteger, , adFldRowID
    rsd.Fields.Append "idValesCab", adChar, 8, adFldIsNullable
    rsd.Fields.Append "fechaEmision", adDate, 10, adFldIsNullable
    rsd.Fields.Append "IdProducto", adVarChar, 8, adFldIsNullable
    rsd.Fields.Append "GlsProducto", adVarChar, 120, adFldIsNullable
    rsd.Fields.Append "IdAlmacen", adVarChar, 8, adFldIsNullable
    rsd.Fields.Append "idProvCliente", adVarChar, 10, adFldIsNullable
    rsd.Fields.Append "ruc", adVarChar, 20, adFldIsNullable
    rsd.Fields.Append "GlsPersona", adVarChar, 180, adFldIsNullable
    rsd.Fields.Append "GlsDocReferencia", adVarChar, 180, adFldIsNullable
    rsd.Fields.Append "abreUM", adVarChar, 10, adFldIsNullable
    rsd.Fields.Append "Ingreso", adDouble, , adFldIsNullable
    rsd.Fields.Append "Salida", adDouble, , adFldIsNullable
    rsd.Fields.Append "Saldo", adDouble, , adFldIsNullable
    rsd.Fields.Append "ValorUnit", adDouble, , adFldIsNullable
    rsd.Fields.Append "ValorTotal", adDouble, , adFldIsNullable
    rsd.Fields.Append "ValorSaldo", adDouble, , adFldIsNullable
    rsd.Fields.Append "FechaInicio", adVarChar, 10, adFldIsNullable
    rsd.Fields.Append "FechaFin", adVarChar, 10, adFldIsNullable
    rsd.Fields.Append "SimboloMoneda", adVarChar, 30, adFldIsNullable
    rsd.Open
    
    CodProd = ""
    i = 0
    Do While Not rsTemp.EOF
        rsd.AddNew
        If CodProd <> rsTemp.Fields("IdProducto") Then
            i = 0
            dblsaldo = traerCantSaldo(rsTemp.Fields("IdProducto"), txtCod_Almacen.Text, strFecIni, StrMsgError)
            If StrMsgError <> "" Then GoTo Err
            
            dblValSaldo = traerCostoUnit(rsTemp.Fields("IdProducto"), txtCod_Almacen.Text, strFecIni, txtCod_Moneda.Text, StrMsgError)
            If StrMsgError <> "" Then GoTo Err
            
            dblValSaldo = dblValSaldo * dblsaldo
        
            rsd.Fields("Item") = i
            rsd.Fields("idValesCab") = ""
            rsd.Fields("fechaEmision") = Format(CDate("1988-01-01"), "yyyy-mm-dd")
            rsd.Fields("IdProducto") = "" & rsTemp.Fields("IdProducto")
            rsd.Fields("GlsProducto") = rsTemp.Fields("GlsProducto")
            rsd.Fields("IdAlmacen") = ""
            rsd.Fields("idProvCliente") = ""
            rsd.Fields("ruc") = ""
            rsd.Fields("GlsPersona") = ""
            rsd.Fields("GlsDocReferencia") = "SALDO INICIAL"
            rsd.Fields("abreUM") = "" & rsTemp.Fields("abreUM")
            rsd.Fields("Ingreso") = 0
            rsd.Fields("Salida") = 0
            rsd.Fields("Saldo") = dblsaldo
            rsd.Fields("ValorUnit") = 0
            rsd.Fields("ValorTotal") = 0
            rsd.Fields("ValorSaldo") = dblValSaldo
            rsd.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
            rsd.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
            rsd.Fields("SimboloMoneda") = strSimboloMoneda
            rsd.AddNew
        End If
        
        i = i + 1
        rsd.Fields("Item") = i
        rsd.Fields("idValesCab") = "" & rsTemp.Fields("idValesCab")
        rsd.Fields("fechaEmision") = "" & Format(rsTemp.Fields("fechaEmision"), "yyyy-mm-dd")
        rsd.Fields("IdProducto") = "" & rsTemp.Fields("IdProducto")
        rsd.Fields("GlsProducto") = "" & rsTemp.Fields("GlsProducto")
        rsd.Fields("IdAlmacen") = "" & rsTemp.Fields("IdAlmacen")
        rsd.Fields("idProvCliente") = "" & rsTemp.Fields("idProvCliente")
        rsd.Fields("ruc") = "" & rsTemp.Fields("ruc")
        rsd.Fields("GlsPersona") = "" & rsTemp.Fields("GlsPersona")
        rsd.Fields("GlsDocReferencia") = "" & rsTemp.Fields("GlsDocReferencia")
        rsd.Fields("abreUM") = "" & rsTemp.Fields("abreUM")
        rsd.Fields("ValorUnit") = Val("" & rsTemp.Fields("VVUnit"))
        
        If rsTemp.Fields("tipoVale") = "S" Then
            rsd.Fields("Ingreso") = 0
            rsd.Fields("Salida") = rsTemp.Fields("Cantidad")
            
            dblsaldo = dblsaldo - rsTemp.Fields("Cantidad")
            rsd.Fields("Saldo") = dblsaldo
            
            rsd.Fields("ValorTotal") = dblsaldo * Val("" & rsTemp.Fields("VVUnit"))
            dblValSaldo = dblValSaldo - rsd.Fields("ValorTotal")
        
        Else
            rsd.Fields("Ingreso") = rsTemp.Fields("Cantidad")
            rsd.Fields("Salida") = 0
            
            dblsaldo = dblsaldo + rsTemp.Fields("Cantidad")
            rsd.Fields("Saldo") = dblsaldo
            
            rsd.Fields("ValorTotal") = dblsaldo * Val("" & rsTemp.Fields("VVUnit"))
            dblValSaldo = dblValSaldo + rsd.Fields("ValorTotal")
        End If
            
        rsd.Fields("ValorSaldo") = dblValSaldo
        rsd.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
        rsd.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
        rsd.Fields("SimboloMoneda") = strSimboloMoneda
        CodProd = rsTemp.Fields("IdProducto")
    
        rsTemp.MoveNext
    Loop
    Set kardex = rsd
    If rsTemp.State = 1 Then rsTemp.Close: Set rsTemp = Nothing
    
    Exit Function
    
Err:
    If rsTemp.State = 1 Then rsTemp.Close
    Set rsTemp = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Function

Private Function traerCantSaldo(ByVal codproducto As String, ByVal codalmacen As String, ByVal strFecha As String, ByRef StrMsgError As String) As Double
On Error GoTo Err
Dim rsSaldo As New ADODB.Recordset
Dim strSQL As String

    traerCantSaldo = 0
    strSQL = "SELECT SUM(IF(vc.tipoVale = 'I',vd.Cantidad," & _
                                            "(vd.Cantidad * -1)" & _
                           ")" & _
                        ") AS STOCK " & _
             "FROM valescab vc,valesdet vd " & _
             "WHERE vc.idValesCab = vd.idValesCab " & _
               "AND vc.idEmpresa = vd.idEmpresa " & _
               "AND vc.idSucursal = vd.idSucursal " & _
               "AND vc.idEmpresa = '" & glsEmpresa & "'" & _
               "AND vc.idSucursal = '" & glsSucursal & "'" & _
               "AND vc.IdAlmacen = '" & codalmacen & "' " & _
               "AND vd.idProducto = '" & codproducto & "' " & _
               "AND vc.idPeriodoInv = '" & glsCodPeriodoINV & "' " & _
               "AND vc.fechaEmision < '" & strFecha & "' " & _
               "AND vc.estvalecab <> 'ANU' "
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

Private Function traerCostoUnit(ByVal codproducto As String, ByVal codalmacen As String, ByVal PFecha As String, ByVal CodMoneda As String, ByRef StrMsgError As String) As Double
On Error GoTo Err
Dim CosUni  As ADODB.Recordset

    csql = "SELECT SUM(IF(valescab.tipoVale = 'I',valesdet.Cantidad,valesdet.Cantidad*-1)) " & _
                       " / " & _
                       "Sum(IF(valescab.tipoVale = 'I'," & _
                                    "CASE '" & CodMoneda & "' WHEN 'PEN' THEN  IF(valescab.idMoneda = 'PEN', valesdet.TotalVVNeto,valesdet.TotalVVNeto * TipoCambio)" & _
                                       "WHEN 'USD' THEN  IF(valescab.idMoneda = 'USD', valesdet.TotalVVNeto,valesdet.TotalVVNeto / TipoCambio)" & _
                                    "END " & _
                                "* valesdet.Cantidad," & _
                                    "(CASE '" & CodMoneda & "' WHEN 'PEN' THEN  IF(valescab.idMoneda = 'PEN', valesdet.TotalVVNeto,valesdet.TotalVVNeto * TipoCambio)" & _
                                       "WHEN 'USD' THEN  IF(valescab.idMoneda = 'USD', valesdet.TotalVVNeto,valesdet.TotalVVNeto / TipoCambio)" & _
                                    "END " & _
                                "* valesdet.Cantidad)*-1)) " & _
                       "AS COSTO_UNITARIO "

    csql = csql & "FROM valescab,valesdet "
    csql = csql & "WHERE valescab.idValesCab = valesdet.idValesCab AND "
    csql = csql & "valescab.idEmpresa = valesdet.idEmpresa AND "
    csql = csql & "valescab.idSucursal = valesdet.idSucursal AND "
    csql = csql & "valescab.idEmpresa = '" & glsEmpresa & "' AND "
    csql = csql & "valescab.idSucursal = '" & glsSucursal & "' AND "
    csql = csql & "valescab.idPeriodoInv = '" & glsCodPeriodoINV & "' AND valescab.fechaEmision <= '" & PFecha & "' And valesdet.idProducto = '" & codproducto & "' AND "
    csql = csql & "valescab.idAlmacen = '" & codalmacen & "' AND "
    csql = csql & "valescab.estvalecab <> 'ANU' "
    
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

Private Sub txtCod_Usuario_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "USUARIO", txtCod_Usuario, txtGls_Usuario
        KeyAscii = 0
        If txtCod_Usuario.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtCod_Vendedor_Change()
    
    txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)

End Sub
