VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmVariablesSistema 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Variables del Sistema"
   ClientHeight    =   5640
   ClientLeft      =   1125
   ClientTop       =   2250
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraVales 
      Appearance      =   0  'Flat
      Caption         =   " Variables de los Vales "
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
      Height          =   1215
      Left            =   60
      TabIndex        =   19
      Top             =   4275
      Width           =   13350
      Begin CATControls.CATTextBox txt_DecimalesVales 
         Height          =   315
         Left            =   1050
         TabIndex        =   16
         Tag             =   "DECIMALESVALES"
         Top             =   525
         Width           =   1290
         _ExtentX        =   2275
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
         Container       =   "frmVariablesSistema.frx":0000
         Estilo          =   3
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Decimales Vales"
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
         Height          =   390
         Left            =   225
         TabIndex        =   31
         Top             =   450
         Width           =   765
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      Caption         =   " Variables Generales "
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
      Height          =   1410
      Left            =   75
      TabIndex        =   18
      Top             =   675
      Width           =   13320
      Begin VB.CommandButton cmbAyudaSystem 
         Height          =   315
         Left            =   6090
         Picture         =   "frmVariablesSistema.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   780
         Width           =   390
      End
      Begin CATControls.CATTextBox txtIGV 
         Height          =   315
         Left            =   1515
         TabIndex        =   0
         Tag             =   "IGV"
         Top             =   375
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmVariablesSistema.frx":03A6
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtValidaStock 
         Height          =   315
         Left            =   8880
         TabIndex        =   1
         Tag             =   "VALIDASTOCK"
         Top             =   345
         Width           =   615
         _ExtentX        =   1085
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
         MaxLength       =   1
         Container       =   "frmVariablesSistema.frx":03C2
         Decimales       =   2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox CATTextBox1 
         Height          =   315
         Left            =   11820
         TabIndex        =   2
         Tag             =   "DECIMALESTIPOCAMBIO"
         Top             =   330
         Width           =   1290
         _ExtentX        =   2275
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
         Container       =   "frmVariablesSistema.frx":03DE
         Estilo          =   3
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_System 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Tag             =   "ENTIDADSYSTEM"
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
         Container       =   "frmVariablesSistema.frx":03FA
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_System 
         Height          =   315
         Left            =   2490
         TabIndex        =   36
         Top             =   780
         Width           =   3540
         _ExtentX        =   6244
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
         Container       =   "frmVariablesSistema.frx":0416
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtRuta 
         Height          =   315
         Left            =   8880
         TabIndex        =   4
         Tag             =   "RUTAIMAGENPROD"
         Top             =   750
         Width           =   4245
         _ExtentX        =   7488
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
         Container       =   "frmVariablesSistema.frx":0432
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Ruta Imagen:"
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
         Height          =   270
         Left            =   7470
         TabIndex        =   44
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Entidad System"
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
         Height          =   270
         Left            =   165
         TabIndex        =   37
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Decimales Tipo de Cambio"
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
         Height          =   390
         Left            =   10560
         TabIndex        =   34
         Top             =   330
         Width           =   1125
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Valida Stock (S/N)"
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
         Left            =   7455
         TabIndex        =   33
         Top             =   390
         Width           =   1395
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
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
         Height          =   240
         Left            =   165
         TabIndex        =   32
         Top             =   450
         Width           =   765
      End
   End
   Begin VB.Frame fraVentas 
      Appearance      =   0  'Flat
      Caption         =   " Variables de Ventas "
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
      Height          =   2070
      Left            =   60
      TabIndex        =   17
      Top             =   2100
      Width           =   13350
      Begin VB.CommandButton cmbAyudaFormaPago 
         Height          =   315
         Left            =   6090
         Picture         =   "frmVariablesSistema.frx":044E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1320
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaCliente 
         Height          =   315
         Left            =   6090
         Picture         =   "frmVariablesSistema.frx":07D8
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   975
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaLista 
         Height          =   315
         Left            =   6090
         Picture         =   "frmVariablesSistema.frx":0B62
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   615
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   12885
         Picture         =   "frmVariablesSistema.frx":0EEC
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaAlmacen 
         Height          =   315
         Left            =   6090
         Picture         =   "frmVariablesSistema.frx":1276
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Almacen 
         Height          =   315
         Left            =   1530
         TabIndex        =   5
         Tag             =   "ALMACENVENTAS"
         Top             =   225
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
         Container       =   "frmVariablesSistema.frx":1600
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Almacen 
         Height          =   315
         Left            =   2505
         TabIndex        =   21
         Top             =   225
         Width           =   3540
         _ExtentX        =   6244
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
         Container       =   "frmVariablesSistema.frx":161C
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   8535
         TabIndex        =   6
         Tag             =   "MONEDAVENTAS"
         Top             =   225
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
         Container       =   "frmVariablesSistema.frx":1638
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   9510
         TabIndex        =   24
         Top             =   225
         Width           =   3315
         _ExtentX        =   5847
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
         Container       =   "frmVariablesSistema.frx":1654
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Lista 
         Height          =   315
         Left            =   1530
         TabIndex        =   7
         Tag             =   "LISTAVENTAS"
         Top             =   615
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
         Container       =   "frmVariablesSistema.frx":1670
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Lista 
         Height          =   315
         Left            =   2505
         TabIndex        =   27
         Top             =   615
         Width           =   3540
         _ExtentX        =   6244
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
         Container       =   "frmVariablesSistema.frx":168C
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_DecimalesCaja 
         Height          =   315
         Left            =   8535
         TabIndex        =   8
         Tag             =   "DECIMALESCAJA"
         Top             =   615
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmVariablesSistema.frx":16A8
         Estilo          =   3
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1530
         TabIndex        =   10
         Tag             =   "CLIENTEVENTAS"
         Top             =   975
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
         Container       =   "frmVariablesSistema.frx":16C4
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2505
         TabIndex        =   39
         Tag             =   "TGlsCliente"
         Top             =   975
         Width           =   3540
         _ExtentX        =   6244
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
         Container       =   "frmVariablesSistema.frx":16E0
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_FormaPago 
         Height          =   315
         Left            =   1530
         TabIndex        =   13
         Tag             =   "FORMAPAGOVENTAS"
         Top             =   1350
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
         Container       =   "frmVariablesSistema.frx":16FC
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_FormaPago 
         Height          =   315
         Left            =   2505
         TabIndex        =   42
         Tag             =   "TGlsCliente"
         Top             =   1350
         Width           =   3540
         _ExtentX        =   6244
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
         Container       =   "frmVariablesSistema.frx":1718
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_SoloGuiaMueveStock 
         Height          =   315
         Left            =   8535
         TabIndex        =   11
         Tag             =   "SOLOGUIAMUEVESTOCK"
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
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
         MaxLength       =   1
         Container       =   "frmVariablesSistema.frx":1734
         Decimales       =   2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_RecepAuto 
         Height          =   315
         Left            =   8535
         TabIndex        =   14
         Tag             =   "RECEPCIONAUTO"
         Top             =   1305
         Width           =   615
         _ExtentX        =   1085
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
         MaxLength       =   1
         Container       =   "frmVariablesSistema.frx":1750
         Decimales       =   2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_DecimalesPrecio 
         Height          =   315
         Left            =   11550
         TabIndex        =   9
         Tag             =   "DECIMALESPRECIOS"
         Top             =   615
         Width           =   1290
         _ExtentX        =   2275
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
         Container       =   "frmVariablesSistema.frx":176C
         Estilo          =   3
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_DctoMinValidacion 
         Height          =   315
         Left            =   11550
         TabIndex        =   12
         Tag             =   "DCTOMINIMOVALIDACION"
         Top             =   960
         Width           =   1290
         _ExtentX        =   2275
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
         Container       =   "frmVariablesSistema.frx":1788
         Estilo          =   3
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_ModVendCampo 
         Height          =   315
         Left            =   11550
         TabIndex        =   15
         Tag             =   "INDMODVENDEDORCAMPO"
         Top             =   1305
         Width           =   615
         _ExtentX        =   1085
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
         MaxLength       =   1
         Container       =   "frmVariablesSistema.frx":17A4
         Decimales       =   2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Mod. Vend de Campo"
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
         Left            =   9645
         TabIndex        =   49
         Top             =   1350
         Width           =   1545
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dcto en % Min. Validacion"
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
         Left            =   9645
         TabIndex        =   48
         Top             =   1020
         Width           =   1875
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Decimales Precio"
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
         Left            =   9645
         TabIndex        =   47
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Recepcion Automatica"
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
         Left            =   6780
         TabIndex        =   46
         Top             =   1350
         Width           =   1620
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Solo Guia Mueve Stock"
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
         Left            =   6780
         TabIndex        =   45
         Top             =   1005
         Width           =   1665
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
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
         Left            =   165
         TabIndex        =   43
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label lbl_Cliente 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente x Defecto"
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
         Left            =   165
         TabIndex        =   40
         Top             =   1035
         Width           =   1230
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Decimales Caja"
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
         Left            =   6780
         TabIndex        =   30
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label lbl_Lista 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Lista"
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
         Left            =   165
         TabIndex        =   28
         Top             =   690
         Width           =   345
      End
      Begin VB.Label lbl_Moneda 
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
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   6780
         TabIndex        =   25
         Top             =   300
         Width           =   570
      End
      Begin VB.Label lbl_Almacen 
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
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   165
         TabIndex        =   22
         Top             =   300
         Width           =   630
      End
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
            Picture         =   "frmVariablesSistema.frx":17C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":1B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":1FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":2346
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":26E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":2A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":2E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":31AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":3548
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":38E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":3C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariablesSistema.frx":493E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   1164
      ButtonWidth     =   1111
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmVariablesSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaAlmacen_Click()
    
    mostrarAyuda "ALMACEN", txtCod_Almacen, txtGls_Almacen
    If txtCod_Almacen.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaFormaPago_Click()

    mostrarAyuda "FORMASPAGO", txtCod_FormaPago, txtGls_FormaPago

End Sub

Private Sub cmbAyudaLista_Click()
    
    mostrarAyuda "LISTAPRECIOS", txtCod_Lista, txtGls_Lista
    If txtCod_Lista.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaMoneda_Click()
   
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaSystem_Click()

    mostrarAyuda "PERSONA", txtCod_System, txtGls_System

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    mostrarParametros StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2 'Salir
            Unload Me
    End Select

    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
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
        txtGls_Cliente.Text = " "
    End If

End Sub

Private Sub txtCod_Cliente_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CLIENTE", txtCod_Cliente, txtGls_Cliente
        KeyAscii = 0
        If txtCod_Cliente.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_FormaPago_Change()
    
    txtGls_FormaPago.Text = traerCampo("formaspagos", "GlsFormaPago", "idFormaPago", txtCod_FormaPago.Text, True)

End Sub

Private Sub txtCod_FormaPago_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "FORMAPAGO", txtCod_FormaPago, txtGls_FormaPago
        KeyAscii = 0
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

End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MONEDA", txtCod_Moneda, txtGls_Moneda
        KeyAscii = 0
        If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim C As Object

    Cn.BeginTrans
    For Each C In Me.Controls
        If TypeOf C Is TextBox Or TypeOf C Is CATTextBox Then
            If C.Tag <> "" Then
                Cn.Execute "UPDATE parametros SET ValParametro = '" & C.Text & "' WHERE GlsParametro = '" & C.Tag & "' AND idEmpresa = '" & glsEmpresa & "'"
            End If
        End If
    Next
    Cn.CommitTrans
    MsgBox "Se Registro satisfactoriamente", vbInformation, App.Title
    
    Exit Sub
    
Err:
    Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarParametros(ByRef StrMsgError As String)
On Error GoTo Err
Dim C As Object

    For Each C In Me.Controls
        If TypeOf C Is TextBox Or TypeOf C Is CATTextBox Then
            If C.Tag <> "" Then
                C.Text = traerCampo("parametros", "ValParametro", "GlsParametro", C.Tag, True)
            End If
        End If
    Next
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_System_Change()
    
    txtGls_System.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_System.Text, False)

End Sub

Private Sub txtCod_System_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PERSONA", txtCod_System, txtGls_System
        KeyAscii = 0
    End If

End Sub
