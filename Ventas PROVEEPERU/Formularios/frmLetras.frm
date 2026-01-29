VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmLetras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Letras"
   ClientHeight    =   9885
   ClientLeft      =   3360
   ClientTop       =   720
   ClientWidth     =   10725
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraGeneral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5595
      Left            =   45
      TabIndex        =   3
      Top             =   675
      Width           =   10635
      Begin VB.CheckBox chkIntereses 
         Caption         =   "Intereses Letra"
         ForeColor       =   &H00000000&
         Height          =   200
         Left            =   90
         TabIndex        =   59
         Top             =   3195
         Width           =   1440
      End
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         Caption         =   " Datos del Documento "
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1890
         Left            =   90
         TabIndex        =   37
         Top             =   135
         Width           =   10425
         Begin CATControls.CATTextBox txt_NumDoc 
            Height          =   315
            Left            =   8685
            TabIndex        =   38
            Top             =   195
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            BackColor       =   12640511
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Arial"
            FontSize        =   9
            ForeColor       =   -2147483640
            Container       =   "frmLetras.frx":0000
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Serie 
            Height          =   315
            Left            =   6510
            TabIndex        =   39
            Top             =   195
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            BackColor       =   12640511
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Arial"
            FontSize        =   9
            ForeColor       =   -2147483640
            Container       =   "frmLetras.frx":001C
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_TipoCambio 
            Height          =   315
            Left            =   900
            TabIndex        =   40
            Top             =   1100
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmLetras.frx":0038
            Text            =   "0"
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_TotalBruto 
            Height          =   315
            Left            =   900
            TabIndex        =   41
            Top             =   1500
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
            Container       =   "frmLetras.frx":0054
            Text            =   "0"
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_TotalIGV 
            Height          =   315
            Left            =   4350
            TabIndex        =   42
            Top             =   1500
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
            Container       =   "frmLetras.frx":0070
            Text            =   "0"
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_TotalNeto 
            Height          =   315
            Left            =   8685
            TabIndex        =   43
            Top             =   1500
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
            Container       =   "frmLetras.frx":008C
            Text            =   "0"
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Moneda 
            Height          =   315
            Left            =   4350
            TabIndex        =   44
            Top             =   1100
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
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
            Container       =   "frmLetras.frx":00A8
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Moneda 
            Height          =   315
            Left            =   5280
            TabIndex        =   45
            Top             =   1095
            Width           =   5070
            _ExtentX        =   8943
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
            Container       =   "frmLetras.frx":00C4
            Vacio           =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp_Emision 
            Height          =   315
            Left            =   9030
            TabIndex        =   46
            Tag             =   "FFecEmision"
            Top             =   675
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   103940097
            CurrentDate     =   38955
         End
         Begin CATControls.CATTextBox txtCod_Cliente 
            Height          =   315
            Left            =   900
            TabIndex        =   47
            Tag             =   "TidPerCliente"
            Top             =   675
            Width           =   915
            _ExtentX        =   1614
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
            Container       =   "frmLetras.frx":00E0
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Cliente 
            Height          =   315
            Left            =   1875
            TabIndex        =   48
            Tag             =   "TGlsCliente"
            Top             =   675
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   556
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
            Container       =   "frmLetras.frx":00FC
            Estilo          =   1
            Vacio           =   -1  'True
         End
         Begin VB.Label lblDoc 
            Appearance      =   0  'Flat
            Caption         =   "Boleta de Venta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   165
            TabIndex        =   58
            Top             =   225
            Width           =   3765
         End
         Begin VB.Label lbl_Moneda 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   3525
            TabIndex        =   57
            Top             =   1125
            Width           =   570
         End
         Begin VB.Label lbl_TotalBruto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Bruto"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   165
            TabIndex        =   56
            Top             =   1500
            Width           =   390
         End
         Begin VB.Label lbl_TotalIGV 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   3525
            TabIndex        =   55
            Top             =   1500
            Width           =   270
         End
         Begin VB.Label lbl_TotalNeto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Total"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   7860
            TabIndex        =   54
            Top             =   1500
            Width           =   345
         End
         Begin VB.Label lbl_Serie 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Serie"
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
            Left            =   5940
            TabIndex        =   53
            Top             =   225
            Width           =   450
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
            Left            =   7920
            TabIndex        =   52
            Top             =   225
            Width           =   675
         End
         Begin VB.Label lbl_TC 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "T/C"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   165
            TabIndex        =   51
            Top             =   1125
            Width           =   240
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   8370
            TabIndex        =   50
            Top             =   750
            Width           =   450
         End
         Begin VB.Label lbl_Cliente 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   165
            TabIndex        =   49
            Top             =   720
            Width           =   480
         End
      End
      Begin VB.Frame fraIntereses 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   1410
         Left            =   90
         TabIndex        =   17
         Top             =   3360
         Width           =   10425
         Begin VB.CommandButton cmbAyudaPeriodoLetra 
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
            Left            =   4470
            Picture         =   "frmLetras.frx":0118
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   225
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaTipoInteres 
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
            Left            =   9975
            Picture         =   "frmLetras.frx":04A2
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   225
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaTipoCuotaLetra 
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
            Left            =   4470
            Picture         =   "frmLetras.frx":082C
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   600
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaTipoCapitalizacion 
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
            Left            =   9975
            Picture         =   "frmLetras.frx":0BB6
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   600
            Width           =   390
         End
         Begin VB.CommandButton cmbActualizarLetras 
            Caption         =   "&Actualizar Intereses"
            Height          =   360
            Left            =   8325
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   975
            Width           =   2025
         End
         Begin CATControls.CATTextBox txtCod_PeriodoLetra 
            Height          =   315
            Left            =   870
            TabIndex        =   23
            Tag             =   "TidPerCliente"
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
            Container       =   "frmLetras.frx":0F40
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_PeriodoLetra 
            Height          =   315
            Left            =   1845
            TabIndex        =   24
            Tag             =   "TGlsCliente"
            Top             =   225
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
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
            Container       =   "frmLetras.frx":0F5C
            Estilo          =   1
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_TipoInteres 
            Height          =   315
            Left            =   6375
            TabIndex        =   25
            Tag             =   "TidPerCliente"
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
            Container       =   "frmLetras.frx":0F78
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_TipoInteres 
            Height          =   315
            Left            =   7350
            TabIndex        =   26
            Tag             =   "TGlsCliente"
            Top             =   225
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
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
            Container       =   "frmLetras.frx":0F94
            Estilo          =   1
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_TipoCuotaLetra 
            Height          =   315
            Left            =   870
            TabIndex        =   27
            Tag             =   "TidPerCliente"
            Top             =   600
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
            Container       =   "frmLetras.frx":0FB0
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_TipoCuotaLetra 
            Height          =   315
            Left            =   1845
            TabIndex        =   28
            Tag             =   "TGlsCliente"
            Top             =   600
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
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
            Container       =   "frmLetras.frx":0FCC
            Estilo          =   1
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_TipoCapitalizacion 
            Height          =   315
            Left            =   6375
            TabIndex        =   29
            Tag             =   "TidPerCliente"
            Top             =   600
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
            Container       =   "frmLetras.frx":0FE8
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_TipoCapitalizacion 
            Height          =   315
            Left            =   7350
            TabIndex        =   30
            Tag             =   "TGlsCliente"
            Top             =   600
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
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
            Container       =   "frmLetras.frx":1004
            Estilo          =   1
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtVal_Tasa 
            Height          =   315
            Left            =   870
            TabIndex        =   31
            Top             =   975
            Width           =   900
            _ExtentX        =   1588
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
            Container       =   "frmLetras.frx":1020
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Periodo"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   150
            TabIndex        =   36
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Interés"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   5325
            TabIndex        =   35
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cuota"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   150
            TabIndex        =   34
            Top             =   675
            Width           =   420
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Capitalización"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   5325
            TabIndex        =   33
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Tasa"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   150
            TabIndex        =   32
            Top             =   1050
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   " Letras "
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   105
         TabIndex        =   9
         Top             =   2085
         Width           =   10425
         Begin VB.CommandButton cmbGenerarLetras 
            Caption         =   "&Generar Letras"
            Height          =   360
            Left            =   8685
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   660
            Width           =   1665
         End
         Begin CATControls.CATTextBox txt_MontoLetras 
            Height          =   330
            Left            =   4335
            TabIndex        =   11
            Top             =   250
            Width           =   1665
            _ExtentX        =   2937
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmLetras.frx":103C
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_CantidadLetras 
            Height          =   330
            Left            =   8685
            TabIndex        =   12
            Top             =   250
            Width           =   1650
            _ExtentX        =   2910
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmLetras.frx":1058
            Decimales       =   2
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp_EmisionLetra 
            Height          =   315
            Left            =   1305
            TabIndex        =   13
            Tag             =   "FFecEmision"
            Top             =   250
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
            Format          =   103940097
            CurrentDate     =   38955
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   3750
            TabIndex        =   16
            Top             =   300
            Width           =   435
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   7860
            TabIndex        =   15
            Top             =   300
            Width           =   630
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   165
            TabIndex        =   14
            Top             =   300
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   " Otros Datos "
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   90
         TabIndex        =   4
         Top             =   4800
         Width           =   10425
         Begin VB.CommandButton cmbAyudaAval 
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
            Left            =   5160
            Picture         =   "frmLetras.frx":1074
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Aval 
            Height          =   315
            Left            =   870
            TabIndex        =   6
            Tag             =   "TidPerVendedor"
            Top             =   240
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
            Container       =   "frmLetras.frx":13FE
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Aval 
            Height          =   315
            Left            =   1845
            TabIndex        =   7
            Top             =   240
            Width           =   3300
            _ExtentX        =   5821
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
            Container       =   "frmLetras.frx":141A
            Vacio           =   -1  'True
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Aval"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   165
            TabIndex        =   8
            Top             =   300
            Width           =   330
         End
      End
   End
   Begin VB.Frame fraLetras 
      Caption         =   " Letras "
      Height          =   3585
      Left            =   45
      TabIndex        =   0
      Top             =   6270
      Width           =   10650
      Begin DXDBGRIDLibCtl.dxDBGrid gLetras 
         Height          =   3225
         Left            =   60
         OleObjectBlob   =   "frmLetras.frx":1436
         TabIndex        =   1
         Top             =   225
         Width           =   10455
      End
      Begin MSComctlLib.ImageList imgDocVentas 
         Left            =   1050
         Top             =   2250
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
               Picture         =   "frmLetras.frx":565A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":59F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":5E46
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":61E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":657A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":6914
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":6CAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":7048
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":73E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":777C
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":7B16
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLetras.frx":87D8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
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
Attribute VB_Name = "frmLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strTD As String
Private strNumDoc As String
Private strSerie As String
Private strEstDoc As String
Private indInserta As Boolean
Private cdirCliente As String
Dim strEstadoLetra As String
Dim streval        As Integer
Dim sw As String

Private Sub chkIntereses_Click()

    fraIntereses.Enabled = chkIntereses.Value
    If chkIntereses.Value Then
        txtCod_PeriodoLetra.Text = "09001"
        txtCod_TipoInteres.Text = "10001"
        txtCod_TipoCuotaLetra.Text = "11001"
        txtCod_TipoCapitalizacion.Text = "12001"
    Else
        txtCod_PeriodoLetra.Text = ""
        txtCod_TipoInteres.Text = ""
        txtCod_TipoCuotaLetra.Text = ""
        txtCod_TipoCapitalizacion.Text = ""
    End If

End Sub

Private Sub cmbActualizarLetras_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim dblInteresTotal As Double
Dim dblInteresIGV As Double
Dim intPlazo As Integer

    If chkIntereses.Value Then
        gLetras.Dataset.First
        Do While Not gLetras.Dataset.EOF
            intPlazo = Val(gLetras.Columns.ColumnByFieldName("Plazo").Value & "")
            If intPlazo <> 0 Then
                gLetras.Dataset.Edit
                gLetras.Columns.ColumnByFieldName("FecVcto").Value = Format(DateAdd("d", intPlazo, dtp_Emision.Value), "dd/mm/yyyy")
    
                If Val(gLetras.Columns.ColumnByFieldName("Item").Value) = Val(txt_CantidadLetras.Value) Then     'si esta en la ultima linea
                    CalculaPorcentaje StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                End If
            End If
            CalculaIntereses dblInteresTotal, dblInteresIGV, StrMsgError
            gLetras.Columns.ColumnByFieldName("Interes").Value = Format(dblInteresTotal, "###,##0.00")
            gLetras.Columns.ColumnByFieldName("IGV").Value = Format(dblInteresIGV, "###,##0.00")
            CalculoTotalLetraFila
            gLetras.Dataset.Next
        Loop
        gLetras.Dataset.Edit
        If gLetras.Dataset.State = dsEdit Or gLetras.Dataset.State = dsInsert Then
            gLetras.Dataset.Post
        End If
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaAval_Click()
    
    mostrarAyuda "AVAL", txtCod_Aval, txtGls_Aval

End Sub

Private Sub cmbAyudaPeriodoLetra_Click()
    
    mostrarAyuda "PERIODOLETRA", txtCod_PeriodoLetra, txtGls_PeriodoLetra

End Sub

Private Sub cmbAyudaTipoCapitalizacion_Click()
    
    mostrarAyuda "TIPOCAPITALIZACION", txtCod_TipoCapitalizacion, txtGls_TipoCapitalizacion

End Sub

Private Sub cmbAyudaTipoCuotaLetra_Click()
    
    mostrarAyuda "TIPOCUOTALETRA", txtCod_TipoCuotaLetra, txtGls_TipoCuotaLetra

End Sub

Private Sub cmbAyudaTipoInteres_Click()
    
    mostrarAyuda "TIPOINTERES", txtCod_TipoInteres, txtGls_TipoInteres

End Sub

Private Sub cmbGenerarLetras_Click()
On Error GoTo Err
Dim StrMsgError As String

    CrearLetras StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    
    indInserta = False
    Txt_TipoCambio.Decimales = glsDecimalesTC
    txt_TotalBruto.Decimales = glsDecimalesCaja
    txt_TotalIGV.Decimales = glsDecimalesCaja
    txt_TotalNeto.Decimales = glsDecimalesCaja
    Toolbar1.Buttons(1).Visible = True
    ConfGrid gLetras, True, True, False, False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub MostrarForm(ByVal strVarTipoDoc As String, ByVal strVarNumDoc As String, ByVal strVarSerie As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rstletras As New ADODB.Recordset

    strTD = strVarTipoDoc
    strNumDoc = strVarNumDoc
    strSerie = strVarSerie
    
    txt_serie.Text = strSerie
    txt_numdoc.Text = strNumDoc
    lblDoc.Caption = traerCampo("documentos", "GlsDocumento", "idDocumento", strTD, False)
    
    csql = "SELECT d.idPerCliente, d.FecEmision, d.idMoneda, d.TipoCambio, d.TotalValorVenta, d.TotalIGVVenta, " & _
             "d.TotalPrecioVenta, d.estDocVentas, d.dirCliente, c.indAgenteRetencion " & _
             "FROM docventas d, clientes c " & _
             "WHERE d.idPerCliente = c.idCliente " & _
             "AND d.idEmpresa = c.idEmpresa " & _
             "AND d.idDocumento = '" & strTD & "' " & _
             "AND d.idDocVentas = '" & strNumDoc & "' " & _
             "AND d.idSerie = '" & strSerie & "' " & _
             "AND d.idEmpresa = '" & glsEmpresa & "' "
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly

    If Not rst.EOF Then
        cdirCliente = "" & rst.Fields("dirCliente")
        txtCod_Cliente.Text = "" & rst.Fields("idPerCliente")
        dtp_Emision.Value = Format("" & rst.Fields("FecEmision"), "dd/mm/yyyy")
        dtp_EmisionLetra.Value = Format(dtp_Emision.Value, "dd/mm/yyyy")
        txtCod_Moneda.Text = "" & rst.Fields("idMoneda")
        Txt_TipoCambio.Text = "" & rst.Fields("TipoCambio")
        strEstDoc = "" & rst.Fields("estDocVentas")
        txt_TotalBruto.Text = "" & rst.Fields("TotalValorVenta")
        txt_TotalIGV.Text = "" & rst.Fields("TotalIGVVenta")
        txt_TotalNeto.Text = "" & rst.Fields("TotalPrecioVenta")
        
        If traerCampo("Empresas", "indRetencion", "idEmpresa", glsEmpresa, False) = "0" Then
            If Val("" & rst.Fields("indAgenteRetencion")) = 1 Then
                txt_MontoLetras.Text = Val(Format(rst.Fields("TotalPrecioVenta") - (rst.Fields("TotalPrecioVenta") * (glsPorcentajeRetencion / 100)), "0.00"))
            Else
                txt_MontoLetras.Text = "" & rst.Fields("TotalPrecioVenta")
            End If
        Else
            txt_MontoLetras.Text = "" & rst.Fields("TotalPrecioVenta")
        End If
        
        txt_CantidadLetras.Text = 1
    End If
    
    nuevo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Me.Show 1
    Unload Me
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo(ByRef StrMsgError As String)
On Error GoTo Err

    listaLetras StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gLetras.Columns.FocusedIndex = gLetras.Columns.ColumnByFieldName("idFormadePago").Index
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub listaLetras(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim cant_Letras As Integer
Dim numLetra As String
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset

    strEstadoLetra = ""
    csql = "SELECT Item,idLetra,Plazo, FecVencimiento, Porcentaje,Capital,Portes,Interes,IGVIntereses,TotalLetra,idLetra,situacion " & _
           "FROM letras " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' " & _
             "AND left(glsDocReferencia, 2) = '" & strTD & "' AND  right(glsDocReferencia, 8) = '" & strNumDoc & "' AND substring(glsDocReferencia,4 ,3) = '" & strSerie & "' and Situacion <> 'A' ORDER BY item"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly

    rsg.Fields.Append "item", adInteger, , adFldRowID
    rsg.Fields.Append "idLetra", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "Plazo", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "FecVcto", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "Porcentaje", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Capital", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Portes", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Interes", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Total", adDouble, 14, adFldIsNullable
    rsg.Open
    
    If rst.EOF Then
        cmbGenerarLetras.Enabled = True
        txt_CantidadLetras.Enabled = True
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idLetra") = ""
        rsg.Fields("Plazo") = 0
        rsg.Fields("FecVcto") = ""
        rsg.Fields("Porcentaje") = 0
        rsg.Fields("Capital") = 0
        rsg.Fields("Portes") = 0
        rsg.Fields("Interes") = 0
        rsg.Fields("IGV") = 0
        rsg.Fields("Total") = 0
        Toolbar1.Buttons(1).Visible = True
        sw = "0"
    Else
        cmbGenerarLetras.Enabled = False
        txt_CantidadLetras.Enabled = False
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idLetra") = "" & rst.Fields("idLetra")
            rsg.Fields("Plazo") = "" & rst.Fields("Plazo")
            rsg.Fields("FecVcto") = "" & rst.Fields("FecVencimiento")
            rsg.Fields("Porcentaje") = "" & rst.Fields("Porcentaje")
            rsg.Fields("Capital") = "" & rst.Fields("Capital")
            rsg.Fields("Portes") = "" & rst.Fields("Portes")
            rsg.Fields("Interes") = "" & rst.Fields("Interes")
            rsg.Fields("IGV") = "" & rst.Fields("IGVIntereses")
            rsg.Fields("Total") = "" & rst.Fields("TotalLetra")
            strEstadoLetra = "" & rst.Fields("Situacion")
            cant_Letras = cant_Letras + 1
            numLetra = "" & rst.Fields("idLetra")
            rst.MoveNext
        Loop
        sw = "1"
        If Not rst.EOF Then
            txtCod_PeriodoLetra.Text = traerCampo("letras", "idPeriodoLetra", "idLetra", numLetra, True, " left(glsDocReferencia, 2) = '" & strTD & "' AND  right(glsDocReferencia, 8) = '" & strNumDoc & "' AND substring(glsDocReferencia,4 ,3) = '" & strSerie & "' ")
            txtCod_TipoInteres.Text = traerCampo("letras", "idTipoInteres", "idLetra", numLetra, True, " left(glsDocReferencia, 2) = '" & strTD & "' AND  right(glsDocReferencia, 8) = '" & strNumDoc & "' AND substring(glsDocReferencia,4 ,3) = '" & strSerie & "' ")
            txtCod_TipoCuotaLetra.Text = traerCampo("letras", "idTipoCuotaLetra", "idLetra", numLetra, True, " left(glsDocReferencia, 2) = '" & strTD & "' AND  right(glsDocReferencia, 8) = '" & strNumDoc & "' AND substring(glsDocReferencia,4 ,3) = '" & strSerie & "' ")
            txtCod_TipoCapitalizacion.Text = traerCampo("letras", "idTipoCapitalizacion", "idLetra", numLetra, True, " left(glsDocReferencia, 2) = '" & strTD & "' AND  right(glsDocReferencia, 8) = '" & strNumDoc & "' AND substring(glsDocReferencia,4 ,3) = '" & strSerie & "' ")
            txtVal_Tasa.Text = traerCampo("letras", "TasaLetra", "idLetra", numLetra, True, " left(glsDocReferencia, 2) = '" & strTD & "' AND  right(glsDocReferencia, 8) = '" & strNumDoc & "' AND substring(glsDocReferencia,4 ,3) = '" & strSerie & "' ")
            txtCod_Aval.Text = traerCampo("letras", "idAval", "idLetra", numLetra, True, " left(glsDocReferencia, 2) = '" & strTD & "' AND  right(glsDocReferencia, 8) = '" & strNumDoc & "' AND substring(glsDocReferencia,4 ,3) = '" & strSerie & "' ")
            txt_CantidadLetras.Text = cant_Letras
        End If
        
        If txtVal_Tasa.Text = "0.00" Then
            chkIntereses.Value = 0
        Else
            chkIntereses.Value = 0
        End If
        txt_MontoLetras.Text = Format(Val(Format(gLetras.Columns.ColumnByFieldName("PREBRUTO").SummaryFooterValue, "0.00")), "###,##0.00")
    End If
    
    If strEstadoLetra = "E" Then
        fraLetras.Enabled = True
        FraGeneral.Enabled = False
        streval = 1
        BloqueaColumnas streval
        Toolbar1.Buttons(1).Visible = False
    Else
        FraGeneral.Enabled = True
        fraLetras.Enabled = True
        streval = 0
        Toolbar1.Buttons(1).Visible = True
    End If
    mostrarDatosGridSQL gLetras, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Me.Refresh
    If rst.State = 1 Then rst.Close: Set rst = Nothing

    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gLetras_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError     As String
Dim strAno          As String
Dim strNumLetra     As String
Dim strSerieLetra   As String

    If Column.Index = gLetras.Columns.ColumnByName("Imprime").Index Then
        If gLetras.Columns.ColumnByFieldName("idLetra").Value <> "" Then
            strAno = Mid(gLetras.Columns.ColumnByFieldName("idANO").Value, 1, 4)
            strNumLetra = gLetras.Columns.ColumnByFieldName("idLetra").Value
            strSerieLetra = "0"
            If MsgBox("Seguro de Imprimir la Letra N° " & strAno & "-" & strNumLetra, vbInformation + vbYesNo, App.Title) = vbYes Then
                Select Case glsFormatoImpLetra
                    Case "GENERAL": IMPRIME_LETRA_GENERAL strAno, strNumLetra, strSerieLetra
                    Case "SINCHI": IMPRIME_LETRA_SINCHI strAno, strNumLetra, strSerieLetra
                    Case "PIC": IMPRIME_LETRA_PIC strAno, strNumLetra, strSerieLetra
                    Case "LATINPLAST": IMPRIME_LETRA_LATINPLAST strAno, strNumLetra, strSerieLetra
                    Case "ITS": IMPRIME_LETRA_ITS strAno, strNumLetra, strSerieLetra
                    Case "HAC": IMPRIME_LETRA_HAC strAno, strNumLetra, strSerieLetra
                    Case "ALSISAC": IMPRIME_LETRA_ALSISAC strAno, strNumLetra, strSerieLetra
                    Case "APIMAS": IMPRIME_LETRA_APIMAS strAno, strNumLetra, strSerieLetra
                End Select
            End If
        Else
            StrMsgError = "Tiene que Grabar las Letras."
            GoTo Err
        End If
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub IMPRIME_LETRA_HAC(strAno As String, strNumLetra As String, strSerieLetra As String)
Dim csql As String
Dim rstletra As New ADODB.Recordset
Dim strimporteletras As String
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strdirec  As String
Dim strtelef As String
Dim StrMsgError As String
Dim straval As String
Dim strdirecaval As String
Dim strTelefonoAval As String
Dim strRUCAval As String
Dim strabrev As String
Dim strcoddis As String
Dim strnomdis As String
Dim strcoddisaval As String
Dim strnomdisaval As String
Dim impresoraletra  As String
Dim localidad As String
Dim cDocReferencia As String

    strAno = Year(dtp_Emision.Value)
    impresoraletra = traerCampo("usuarios", "ImpresoraLetras", "idUsuario", glsUser, True)
    
    If impresoraletra <> "" Then
        For Each p In Printers
           If UCase(p.DeviceName) = impresoraletra Then
              Set Printer = p
              indPrinter = True
              Exit For
           End If
        Next p
    End If
    
    strNumLetra = Format(strNumLetra, "000000")
    intScale = 6
    Printer.ScaleMode = intScale
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    
    csql = "SELECT idletra,glsDocreferencia,FecEmision,FecVencimiento,TotalLetra," & _
           "CASE idMoneda WHEN 'PEN' THEN 'S/.' ELSE 'US$' END AS idMoneda,idCliente,glsCliente,RucCliente,idaval " & _
           "FROM LETRAS  " & _
           "WHERE IDEMPRESA = '" & glsEmpresa & "' " & _
           "AND IDLETRA = '" & strNumLetra & "' " & _
           "AND IDSERIE = '" & strSerieLetra & "' and idAno = '" & Trim("" & strAno) & "' "
    rstletra.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstletra.EOF Then
        '--- NUMERO LETRA
        ImprimeXY "HAC - " & rstletra.Fields("idLetra") & "", "T", 20, 27, 25, 0, 0, StrMsgError
        strabrev = traerCampo("documentos", "abredocumento", "iddocumento", strTD, False)
        cDocReferencia = left(strabrev, 1) & "."
        cDocReferencia = "" & Val(strNumDoc)
        
        '--- REFERENCIA
        ImprimeXY left(cDocReferencia, Len(cDocReferencia) - 1), "T", 19, 22, 49, 0, 3, StrMsgError
        '--- LUGAR DE GIRO
        localidad = traerCampo("personas", "glsPersona", "idPersona", glsSucursal, False)
        ImprimeXY localidad, "T", 20, 24, 111, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO EMISION
        ImprimeXY Format(Year(rstletra.Fields("FecEmision")), "0000") & "", "T", 4, 25, 70 + 21, 0, 0, StrMsgError
        '--- IMPRIME EL MES EMISION
        ImprimeXY Format(Month(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 25, 70 + 14, 0, 0, StrMsgError
        '--- IMPRIME EL DIA EMISION
        ImprimeXY Format(Day(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 25, 76, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO VENCIMIENTO
        ImprimeXY Format(Year(rstletra.Fields("FecVencimiento")), "0000") & "", "T", 4, 25, 130 + 21, 0, 0, StrMsgError
        '--- IMPRIME EL MES VENCIMIENTO
        ImprimeXY Format(Month(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 25, 130 + 14, 0, 0, StrMsgError
        '--- IMPRIME EL DIA VENCIMIENTO
        ImprimeXY Format(Day(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 25, 136, 0, 0, StrMsgError
        '--- Signo de la modeda
        ImprimeXY rstletra.Fields("idMoneda") & "", "T", 3, 24, 163, 2, 0, StrMsgError
        '--- IMPORTE
        ImprimeXY rstletra.Fields("TotalLetra") & "", "N", 18, 24, 163, 2, 0, StrMsgError
        '--- IMPORTE EN LETRAS
        strimporteletras = "SON:" & Cadenanum(Format(rstletra.Fields("TotalLetra"), "0.00"), IIf(rstletra.Fields("idMoneda") = "S/.", "NUEVOS SOLES", "DOLARES AMERICANOS"))
        ImprimeXY Mid(strimporteletras, 5, Len(strimporteletras) - 3) & "", "T", 250, 43, 24, 0, 0, StrMsgError
        '--- CLIENTE
        ImprimeXY rstletra.Fields("glsCliente") & "", "T", 50, 52, 34, 0, 0, StrMsgError
        '--- DIRECCION
        strdirec = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idCliente"), False)
        ImprimeXY left(strdirec, 120) & "", "T", 145, 61, 34, 0, 0, StrMsgError
        '--- LOCALIDAD
        strcoddis = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idCliente"), False)
        strcoddis = left(strcoddis, 4) & "00"
        strnomdis = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddis, False)
        ImprimeXY strnomdis & "", "T", 150, 64, 80, 0, 0, StrMsgError
        '--- RUC
        ImprimeXY rstletra.Fields("RucCliente") & "", "T", 153, 69, 34, 0, 0, StrMsgError
        '--- TELEFONO
        strtelef = traerCampo("personas", "telefonos", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 8
        ImprimeXY left(strtelef, 12) & "", "T", 153, 69, 80, 0, 0, StrMsgError
        
        If rstletra.Fields("idAval") <> "" Then
            straval = traerCampo("personas", "glspersona", "idpersona", rstletra.Fields("idaval"), False)
            strdirecaval = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY straval & "", "T", 150, 72, 50, 0, 0, StrMsgError
            ImprimeXY left(strdirecaval, 50) & "", "T", 150, 78, 50, 0, 0, StrMsgError
            
            strcoddisaval = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idaval"), False)
            strnomdisaval = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddisaval, False)
            ImprimeXY strnomdisaval & "", "T", 150, 85, 50, 0, 0, StrMsgError
            
            strTelefonoAval = traerCampo("personas", "Telefonos", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY left(strTelefonoAval, 12) & "", "T", 153, 84, 90, 0, 0, StrMsgError
            
            strRUCAval = traerCampo("personas", "RUC", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY strRUCAval, "T", 153, 91, 39, 0, 0, StrMsgError
        End If
        
        Printer.Print Chr$(149)
        Printer.Print ""
        Printer.Print ""
        Printer.EndDoc
    End If

End Sub

Private Sub IMPRIME_LETRA_ITS(strAno As String, strNumLetra As String, strSerieLetra As String)
Dim csql As String
Dim rstletra As New ADODB.Recordset
Dim strimporteletras As String
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strdirec  As String
Dim strtelef As String
Dim StrMsgError As String
Dim straval As String
Dim strdirecaval As String
Dim strTelefonoAval As String
Dim strRUCAval As String
Dim strabrev As String
Dim strcoddis As String
Dim strnomdis As String
Dim strcoddisaval As String
Dim strnomdisaval As String
Dim impresoraletra  As String
Dim localidad As String

    strAno = Year(dtp_Emision.Value)
    impresoraletra = traerCampo("usuarios", "ImpresoraLetras", "idUsuario", glsUser, True)
    If impresoraletra <> "" Then
        For Each p In Printers
           If UCase(p.DeviceName) = impresoraletra Then
              Set Printer = p
              indPrinter = True
              Exit For
           End If
        Next p
    End If
    
    strNumLetra = Format(strNumLetra, "000000")
    intScale = 6
    Printer.ScaleMode = intScale
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    
    csql = "SELECT idletra,glsDocreferencia,FecEmision,FecVencimiento,TotalLetra," & _
           "CASE idMoneda WHEN 'PEN' THEN 'S/.' ELSE 'US$' END AS idMoneda,idCliente,glsCliente,RucCliente,idaval " & _
           "FROM LETRAS  " & _
           "WHERE IDEMPRESA = '" & glsEmpresa & "' " & _
           "AND IDLETRA = '" & strNumLetra & "' " & _
           "AND IDSERIE = '" & strSerieLetra & "' and idAno = '" & Trim("" & strAno) & "' "
    rstletra.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstletra.EOF Then
        '--- NUMERO LETRA
        ImprimeXY rstletra.Fields("idLetra") & "", "T", 20, 22, 48, 0, 0, StrMsgError
        strabrev = traerCampo("documentos", "abredocumento", "iddocumento", left(rstletra.Fields("glsDocreferencia"), 2), False)
        '--- REFERENCIA
        ImprimeXY strabrev & " " & right(rstletra.Fields("glsDocreferencia"), 12) & "", "T", 20, 22, 63, 0, 0, StrMsgError
        '--- LUGAR DE GIRO
        localidad = traerCampo("personas", "glsPersona", "idPersona", glsSucursal, False)
        ImprimeXY localidad, "T", 20, 22, 95, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO EMISION
        ImprimeXY Format(Year(rstletra.Fields("FecEmision")), "0000") & "", "T", 4, 22, 91 + 37, 0, 0, StrMsgError
        '--- IMPRIME EL MES EMISION
        ImprimeXY Format(Month(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 22, 91 + 32, 0, 0, StrMsgError
        '--- IMPRIME EL DIA EMISION
        ImprimeXY Format(Day(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 22, 91 + 27, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO VENCIMIENTO
        ImprimeXY Format(Year(rstletra.Fields("FecVencimiento")), "0000") & "", "T", 4, 22, 91 + 63, 0, 0, StrMsgError
        '--- IMPRIME EL MES VENCIMIENTO
        ImprimeXY Format(Month(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 22, 91 + 58, 0, 0, StrMsgError
        '--- IMPRIME EL DIA VENCIMIENTO
        ImprimeXY Format(Day(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 22, 91 + 53, 0, 0, StrMsgError
        '--- Signo de la modeda
        ImprimeXY rstletra.Fields("idMoneda") & "", "T", 3, 22, 175, 2, 0, StrMsgError
        '--- IMPORTE
        ImprimeXY rstletra.Fields("TotalLetra") & "", "N", 18, 22, 175, 2, 0, StrMsgError
        '--- IMPORTE EN LETRAS
        strimporteletras = "SON:" & Cadenanum(Format(rstletra.Fields("TotalLetra"), "0.00"), IIf(rstletra.Fields("idMoneda") = "S/.", "NUEVOS SOLES", "DOLARES AMERICANOS"))
        ImprimeXY Mid(strimporteletras, 5, Len(strimporteletras) - 3) & "", "T", 250, 40, 58, 0, 0, StrMsgError
        '--- CLIENTE
        ImprimeXY rstletra.Fields("glsCliente") & "", "T", 153, 52, 54, 0, 0, StrMsgError
        '--- RUC
        ImprimeXY rstletra.Fields("RucCliente") & "", "T", 153, 64, 55, 0, 0, StrMsgError
        '--- LOCALIDAD
        strcoddis = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idCliente"), False)
        strcoddis = left(strcoddis, 4) & "00"
        strnomdis = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddis, False)
        ImprimeXY strnomdis & "", "T", 150, 60, 95, 0, 0, StrMsgError
        
        strdirec = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idCliente"), False)
        strtelef = traerCampo("personas", "telefonos", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 6
        ImprimeXY left(strdirec, 120) & "", "T", 145, 60, 56, 0, 0, StrMsgError
        Printer.FontSize = 8
        ImprimeXY left(strtelef, 12) & "", "T", 153, 64, 95, 0, 0, StrMsgError
        
        If rstletra.Fields("idAval") <> "" Then
            straval = traerCampo("personas", "glspersona", "idpersona", rstletra.Fields("idaval"), False)
            strdirecaval = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY straval & "", "T", 150, 71, 63, 0, 0, StrMsgError
            ImprimeXY left(strdirecaval, 50) & "", "T", 150, 77, 53, 0, 0, StrMsgError
            
            strcoddisaval = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idaval"), False)
            strnomdisaval = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddisaval, False)
            ImprimeXY strnomdisaval & "", "T", 150, 84, 53, 0, 0, StrMsgError
            
            strTelefonoAval = traerCampo("personas", "Telefonos", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY left(strTelefonoAval, 12) & "", "T", 153, 83, 102, 0, 0, StrMsgError
            
            strRUCAval = traerCampo("personas", "RUC", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY strRUCAval, "T", 153, 90, 52, 0, 0, StrMsgError
        End If
        
        Printer.Print Chr$(149)
        Printer.Print ""
        Printer.Print ""
        Printer.EndDoc
    End If

End Sub

Private Sub IMPRIME_LETRA_LATINPLAST(strAno As String, strNumLetra As String, strSerieLetra As String)
Dim csql As String
Dim rstletra As New ADODB.Recordset
Dim strimporteletras As String
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strdirec  As String
Dim strtelef As String
Dim StrMsgError As String
Dim straval As String
Dim strdirecaval As String
Dim strTelefonoAval As String
Dim strRUCAval As String
Dim strabrev As String
Dim strcoddis As String
Dim strnomdis As String
Dim strcoddisaval As String
Dim strnomdisaval As String
Dim impresoraletra  As String
Dim localidad As String

    strAno = Year(dtp_Emision.Value)
    impresoraletra = traerCampo("usuarios", "ImpresoraLetras", "idUsuario", glsUser, True)
    
    If impresoraletra <> "" Then
        For Each p In Printers
           If UCase(p.DeviceName) = impresoraletra Then
              Set Printer = p
              indPrinter = True
              Exit For
           End If
        Next p
    End If
    
    strNumLetra = Format(strNumLetra, "000000")
    intScale = 6
    Printer.ScaleMode = intScale
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    
    csql = "SELECT idletra,glsDocreferencia,FecEmision,FecVencimiento,TotalLetra," & _
           "CASE idMoneda WHEN 'PEN' THEN 'S/.' ELSE 'US$' END AS idMoneda,idCliente,glsCliente,RucCliente,idaval " & _
           "FROM LETRAS  " & _
           "WHERE IDEMPRESA = '" & glsEmpresa & "' " & _
           "AND IDLETRA = '" & strNumLetra & "' " & _
           "AND IDSERIE = '" & strSerieLetra & "' and idAno = '" & Trim("" & strAno) & "' "
    rstletra.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstletra.EOF Then
        '--- NUMERO LETRA
        ImprimeXY rstletra.Fields("idLetra") & "", "T", 20, 25, 25, 0, 0, StrMsgError
        strabrev = traerCampo("documentos", "abredocumento", "iddocumento", left(rstletra.Fields("glsDocreferencia"), 2), False)
        '--- REFERENCIA
        ImprimeXY strabrev & " " & right(rstletra.Fields("glsDocreferencia"), 12) & "", "T", 20, 25, 47, 0, 0, StrMsgError
        '--- LUGAR DE GIRO
        localidad = traerCampo("personas", "glsPersona", "idPersona", glsSucursal, False)
        ImprimeXY localidad, "T", 20, 24, 111, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO EMISION
        ImprimeXY Format(Year(rstletra.Fields("FecEmision")), "0000") & "", "T", 4, 25, 70 + 21, 0, 0, StrMsgError
        '--- IMPRIME EL MES EMISION
        ImprimeXY Format(Month(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 25, 70 + 14, 0, 0, StrMsgError
        '--- IMPRIME EL DIA EMISION
        ImprimeXY Format(Day(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 25, 75, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO VENCIMIENTO
        ImprimeXY Format(Year(rstletra.Fields("FecVencimiento")), "0000") & "", "T", 4, 25, 130 + 21, 0, 0, StrMsgError
        '--- IMPRIME EL MES VENCIMIENTO
        ImprimeXY Format(Month(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 25, 130 + 14, 0, 0, StrMsgError
        ''IMPRIME EL DIA VENCIMIENTO
        ImprimeXY Format(Day(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 25, 136, 0, 0, StrMsgError
        '--- Signo de la modeda
        ImprimeXY rstletra.Fields("idMoneda") & "", "T", 3, 25, 163, 2, 0, StrMsgError
        '--- IMPORTE
        ImprimeXY rstletra.Fields("TotalLetra") & "", "N", 18, 25, 163, 2, 0, StrMsgError
        '--- IMPORTE EN LETRAS
        strimporteletras = "SON:" & Cadenanum(Format(rstletra.Fields("TotalLetra"), "0.00"), IIf(rstletra.Fields("idMoneda") = "S/.", "NUEVOS SOLES", "DOLARES AMERICANOS"))
        ImprimeXY Mid(strimporteletras, 5, Len(strimporteletras) - 3) & "", "T", 250, 38, 24, 0, 0, StrMsgError
        '--- CLIENTE
        ImprimeXY rstletra.Fields("glsCliente") & "", "T", 50, 51, 34, 0, 0, StrMsgError
        '--- DIRECCION
        strdirec = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 6
        ImprimeXY left(strdirec, 120) & "", "T", 145, 58, 36, 0, 0, StrMsgError
        '--- LOCALIDAD
        strcoddis = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idCliente"), False)
        strcoddis = left(strcoddis, 4) & "00"
        strnomdis = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddis, False)
        ImprimeXY strnomdis & "", "T", 150, 62, 73, 0, 0, StrMsgError
        '--- RUC
        ImprimeXY rstletra.Fields("RucCliente") & "", "T", 153, 67, 38, 0, 0, StrMsgError
        '--- TELEFONO
        strtelef = traerCampo("personas", "telefonos", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 8
        ImprimeXY left(strtelef, 12) & "", "T", 153, 67, 75, 0, 0, StrMsgError
        
        If rstletra.Fields("idAval") <> "" Then
            straval = traerCampo("personas", "glspersona", "idpersona", rstletra.Fields("idaval"), False)
            strdirecaval = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY straval & "", "T", 150, 72, 50, 0, 0, StrMsgError
            ImprimeXY left(strdirecaval, 50) & "", "T", 150, 78, 50, 0, 0, StrMsgError
            
            strcoddisaval = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idaval"), False)
            strnomdisaval = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddisaval, False)
            ImprimeXY strnomdisaval & "", "T", 150, 85, 50, 0, 0, StrMsgError
            
            strTelefonoAval = traerCampo("personas", "Telefonos", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY left(strTelefonoAval, 12) & "", "T", 153, 84, 90, 0, 0, StrMsgError
            
            strRUCAval = traerCampo("personas", "RUC", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY strRUCAval, "T", 153, 91, 39, 0, 0, StrMsgError
        End If
        Printer.Print Chr$(149)
        Printer.Print ""
        Printer.Print ""
        Printer.EndDoc
    End If
    
End Sub

Private Sub gLetras_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError     As String
Dim dblInteresTotal As Double
Dim dblInteresIGV   As Double
Dim intPlazo        As Integer
Dim intDia          As Integer
Dim dblCapital      As Double
Dim dblPorcentaje   As Double

    If gLetras.Dataset.State = dsEdit Then
        Select Case gLetras.Columns.FocusedColumn.Index
            Case gLetras.Columns.ColumnByFieldName("Plazo").Index
                intPlazo = Val(gLetras.Columns.ColumnByFieldName("Plazo").Value & "")
                If intPlazo <> 0 Then
                    gLetras.Dataset.Edit
                    gLetras.Columns.ColumnByFieldName("FecVcto").Value = Format(DateAdd("d", intPlazo, dtp_Emision.Value), "dd/mm/yyyy")
                End If
            
            Case gLetras.Columns.ColumnByFieldName("FecVcto").Index
                If IsDate(gLetras.Columns.ColumnByFieldName("FecVcto").Value) Then
                    gLetras.Columns.ColumnByFieldName("Plazo").Value = CDate(gLetras.Columns.ColumnByFieldName("FecVcto").Value) - CDate(dtp_Emision.Value)
                    If Val(gLetras.Columns.ColumnByFieldName("Plazo").Value & "") > 0 Then
                        If chkIntereses.Value Then
                            CalculaIntereses dblInteresTotal, dblInteresIGV, StrMsgError
                            gLetras.Columns.ColumnByFieldName("Interes").Value = Format(dblInteresTotal, "###,##0.00")
                        End If
                    Else
                        gLetras.Columns.ColumnByFieldName("Plazo").Value = 0
                        gLetras.Columns.ColumnByFieldName("FecVcto").Value = ""
                        StrMsgError = "Verifique, Fecha incorrecta."
                        GoTo Err
                    End If
                End If
            
            Case gLetras.Columns.ColumnByFieldName("Porcentaje").Index
                dblPorcentaje = gLetras.Columns.ColumnByFieldName("Porcentaje").Value
                dblCapital = Val(txt_MontoLetras.Value) * (dblPorcentaje / 100)
                gLetras.Columns.ColumnByFieldName("Capital").Value = Format(dblCapital, "###,##0.00")
                CalculoTotalLetraFila
            
            Case gLetras.Columns.ColumnByFieldName("Capital").Index
                If Val(Format(gLetras.Columns.ColumnByFieldName("Capital").Value, "0.00")) > Format(txt_MontoLetras.Value, "0.00") Then
                    StrMsgError = "El capital es mayor al monto, ingrese un nuevo capital"
                    gLetras.Columns.ColumnByFieldName("Capital").Value = 0
                    If gLetras.Dataset.State = dsEdit Then gLetras.Dataset.Post
                    GoTo Err
                Else
                    dblCapital = gLetras.Columns.ColumnByFieldName("Capital").Value
                    dblPorcentaje = (dblCapital * 100) / Val(txt_MontoLetras.Value)
                    gLetras.Columns.ColumnByFieldName("Porcentaje").Value = Format(dblPorcentaje, "###,##0.00")
                End If
                CalculoTotalLetraFila
            
            Case gLetras.Columns.ColumnByFieldName("Portes").Index
                If chkIntereses.Value Then   'si es con interes
                    CalculaIntereses dblInteresTotal, dblInteresIGV, StrMsgError
                    gLetras.Columns.ColumnByFieldName("Interes").Value = Format(dblInteresTotal, "###,##0.00")
                    CalculoTotalLetraFila 'RECIEN
                End If
                gLetras.Columns.ColumnByFieldName("IGV").Value = Format(Val(Format(gLetras.Columns.ColumnByFieldName("Portes").Value, "0.00")) * (glsIGV / 100), "0.00")
                CalculoTotalLetraFila
            
            Case gLetras.Columns.ColumnByFieldName("Interes").Index
                gLetras.Columns.ColumnByFieldName("IGV").Value = Format(Val(Format(gLetras.Columns.ColumnByFieldName("Interes").Value, "0.00")) * (glsIGV / 100), "0.00")
                CalculoTotalLetraFila
        End Select
        
        If gLetras.Dataset.State = dsEdit Then gLetras.Dataset.Post
    End If
   
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
            If sw = "0" Then
                Grabar StrMsgError
                If StrMsgError <> "" Then GoTo Err
            ElseIf sw = "1" Then
                Modificar StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
        Case 2 'Cancelar
            Unload Me
        Case 4 'Salir
            Unload Me
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Aval_Change()
    
    txtGls_Aval.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Aval.Text, False)

End Sub

Private Sub txtCod_Aval_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "AVAL", txtCod_Aval, txtGls_Aval
        KeyAscii = 0
        If txtCod_Aval.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub txtCod_Cliente_Change()
    
    txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Public Sub CrearLetras(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg As New ADODB.Recordset
Dim strNumLetra As String
Dim strAno As String
Dim strFecVenc As String
Dim strFormaPago As String
Dim dblCapital As Double
Dim dblMonto As Double
Dim dblPorcentaje As Double
Dim intNumLetras As Integer
Dim intPlazo As Integer
Dim i As Integer
    
    rsg.Fields.Append "item", adInteger, , adFldRowID
    rsg.Fields.Append "idLetra", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "Plazo", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "FecVcto", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "Porcentaje", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Capital", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Portes", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Interes", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "IGV", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Total", adDouble, 14, adFldIsNullable
    rsg.Open
    
    dblMonto = txt_MontoLetras.Value
    intNumLetras = txt_CantidadLetras.Value
    strAno = Year(dtp_EmisionLetra.Value)
    
    If dblMonto > 0 Then
        If intNumLetras > 0 Then
            strFormaPago = traerCampo("movcajasdet", "idFormadePago", "iddocumento", strTD, True, " idserie = '" & strSerie & "' and iddocventas = '" & strNumDoc & "' ")
            intPlazo = Val(traerCampo("formaspagos", "diasVcto", "idFormaPago", strFormaPago, True))
            dblPorcentaje = 100 / intNumLetras
            dblCapital = Format(dblMonto / intNumLetras, "0.00")
            strNumLetra = ""
        
            If Trim(traerCampo("Parametros", "ValParametro", "Glsparametro", "GENERA_CORRELATIVO_LETRA_MES", True) & "") = "S" Then
                For i = 0 To intNumLetras - 1
                    strNumLetra = right(generaCorrelativo("letras", "idLetra", 6, , True, "month(FecEmision)='" & Format(Month(dtp_EmisionLetra.Value), "00") & "'"), 4)
                    strFecVenc = CStr(dtp_EmisionLetra.Value + intPlazo)
                    rsg.AddNew
                    rsg.Fields("Item") = i
                    rsg.Fields("idLetra") = Format(Month(dtp_EmisionLetra.Value), "00") & Format(CInt(Val(strNumLetra)) + i, "0000")
                    rsg.Fields("Plazo") = intPlazo
                    rsg.Fields("FecVcto") = strFecVenc
                    rsg.Fields("Porcentaje") = dblPorcentaje
                    rsg.Fields("Capital") = dblCapital
                    rsg.Fields("Portes") = 0
                    rsg.Fields("Interes") = 0
                    rsg.Fields("IGV") = 0
                    rsg.Fields("Total") = dblCapital
                    intPlazo = intPlazo + 15
                Next i
                
            Else
                For i = 1 To intNumLetras
                    strFecVenc = CStr(dtp_EmisionLetra.Value + intPlazo)
                    strNumLetra = traerCampo("Letras", "max(idLetra)", "1", "1", True)
                    rsg.AddNew
                    rsg.Fields("Item") = i
                    rsg.Fields("idLetra") = Format(CInt(Val(strNumLetra)) + i, "000000") 'strAno & "-" & strNumLetra
                    rsg.Fields("Plazo") = intPlazo
                    rsg.Fields("FecVcto") = strFecVenc
                    rsg.Fields("Porcentaje") = dblPorcentaje
                    rsg.Fields("Capital") = dblCapital
                    rsg.Fields("Portes") = 0
                    rsg.Fields("Interes") = 0
                    rsg.Fields("IGV") = 0
                    rsg.Fields("Total") = dblCapital
                    intPlazo = intPlazo + intPlazo
                Next i
            End If
        End If

        mostrarDatosGridSQL gLetras, rsg, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Else
        StrMsgError = "Ingrese monto mayor a 0"
        GoTo Err
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub CalculaTotalesLetras(ByRef StrMsgError As String)
    
    If gLetras.Dataset.State = dsEdit Then
        gLetras.Dataset.Post
    End If

End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strAno        As String
Dim strMoneda     As String
Dim strNumLetra   As String
Dim strRUCCliente   As String
Dim indTrans As Boolean
Dim cnn_empresa     As New ADODB.Connection
Dim rsbusca         As New ADODB.Recordset
Dim cconex_empresa  As String, cdirectorio As String, cruta As String, ccliente As String, cbusca As String
Dim cinsert         As String, cmoneda As String, cupdate  As String

    indTrans = False
    strAno = Year(dtp_Emision.Value)
    strMoneda = txtCod_Moneda.Text
    strRUCCliente = traerCampo("personas", "RUC", "idPersona", txtCod_Cliente.Text, False)
    
    If gLetras.Dataset.State = dsEdit Or gLetras.Dataset.State = dsInsert Then gLetras.Dataset.Post
    indTrans = True
    
    Cn.BeginTrans
    
    strNumLetra = generaCorrelativo("letras", "idLetra", 6, , True)
    gLetras.Dataset.First
    Do While Not gLetras.Dataset.EOF
        csql = "INSERT INTO letras(idEmpresa,idSucursal,idAno,idLetra,idSerie,GlsDocReferencia,Plazo,FecVencimiento," & _
            " Porcentaje,Capital,Portes,Interes,TotalLetra,MontoOriginal,idMoneda, " & _
            " IGVIntereses,FecEmision,idCliente,GlsCliente,RUCCliente,TipoCambio,TasaLetra, " & _
            " TotalDocumento,idUsuarioReg,Item,idPeriodoLetra,idTipoInteres,idTipoCuotaLetra,idTipoCapitalizacion,FecRegistro,idAval,Situacion)"
            
        csql = csql & " VALUES('" & glsEmpresa & "','" & glsSucursal & "','" & strAno & "','" & strNumLetra & "',0,'" & (strTD & "-" & strSerie & "-" & strNumDoc) & "'," & _
             gLetras.Columns.ColumnByFieldName("Plazo").Value & ",'" & Format(gLetras.Columns.ColumnByFieldName("FecVcto").Value, "yyyy-mm-dd") & "'," & gLetras.Columns.ColumnByFieldName("Porcentaje").Value & "," & gLetras.Columns.ColumnByFieldName("Capital").Value & "," & _
             gLetras.Columns.ColumnByFieldName("Portes").Value & "," & gLetras.Columns.ColumnByFieldName("Interes").Value & "," & gLetras.Columns.ColumnByFieldName("Total").Value & "," & gLetras.Columns.ColumnByFieldName("Total").Value & ",'" & txtCod_Moneda.Text & "'," & _
             gLetras.Columns.ColumnByFieldName("IGV").Value & ",'" & Format(dtp_EmisionLetra.Value, "yyyy-mm-dd") & "','" & txtCod_Cliente.Text & "','" & txtGls_Cliente.Text & "','" & strRUCCliente & "'," & _
             Txt_TipoCambio.Value & "," & txtVal_Tasa.Value & "," & txt_TotalNeto.Value & ",'" & glsUser & "'," & gLetras.Columns.ColumnByFieldName("Item").Value & ",'" & _
             txtCod_PeriodoLetra.Text & "','" & txtCod_TipoInteres.Text & "','" & txtCod_TipoCuotaLetra.Text & "','" & txtCod_TipoCapitalizacion.Text & "',sysdate(),'" & txtCod_Aval.Text & "','X')"
        
        Cn.Execute csql
        
        '---------------------------------------------------------------------------------------------------------------
        cdirectorio = traerCampo("empresas", "Carpeta", "idEmpresa", glsEmpresa, False)
        If cdirectorio <> "" And glsSistemaAccess = "S" Then 'Grabamos en la Version de Access
            cruta = glsRuta_Access & cdirectorio
            If cnn_empresa.State = adStateOpen Then cnn_empresa.Close
            cconex_empresa = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cruta & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
            cnn_empresa.Open cconex_empresa
    
            '--------------------------------
            '--- CLIENTE
            ccliente = ""
            cbusca = "SELECT F2CODCLI FROM EF2CLIENTES WHERE F2NEWRUC = '" & strRUCCliente & "'"
            If rsbusca.State = adStateOpen Then rsbusca.Close
            rsbusca.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
            If Not rsbusca.EOF Then
                ccliente = Trim(rsbusca.Fields("F2CODCLI") & "")
            End If
            rsbusca.Close: Set rsbusca = Nothing
            '--------------------------------
            '---- AGREGA CLIENTE
            If Len(ccliente) = 0 Then
                ccliente = ""
                cbusca = "SELECT F2CODCLI FROM EF2CLIENTES ORDER BY F2CODCLI DESC"
                If rsbusca.State = adStateOpen Then rsbusca.Close
                rsbusca.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
                If Not rsbusca.EOF Then
                    ccliente = Format(Val(rsbusca.Fields("F2CODCLI") & "") + 1, "0000")
                End If
                rsbusca.Close: Set rsbusca = Nothing
                cinsert = "INSERT INTO EF2CLIENTES " & _
                          "(F2CODCLI,F2NOMCLI,F2NEWRUC,F2DIRCLI,F2TIPDOC) " & _
                          "VALUES ('" & ccliente & "','" & txtGls_Cliente.Text & "','" & strRUCCliente & "','" & cdirCliente & "','J')"
                cnn_empresa.Execute (cinsert)
            End If
            '--------------------------------
            cmoneda = IIf(txtCod_Moneda.Text = "PEN", "S", "D")
            cinsert = "INSERT INTO LETRAS " & _
                      "(ANO_LETRA,NRO_LETRA,SER_LETRA,NRO_REF,PLAZO,FCH_VENC," & _
                      "PORCENTAJE,CAPITAL,PORTES,INTERES,TOT_LETRA,MONTO_ORIGINAL," & _
                      "TIP_DCTO,TIP_MONE,IGV_INT,FCH_EMIS,CLIENTE,NOMCLI,ruccli,TIP_CAMB,TASA_INT,SITUACION,tip_inte,NRO_DCTO,TOT_DCTO,usuario) " & _
                      "VALUES ('" & strAno & "','" & strNumLetra & "','0',''," & gLetras.Columns.ColumnByFieldName("Plazo").Value & _
                      ",CVDate('" & Format(gLetras.Columns.ColumnByFieldName("FecVcto").Value, "dd/mm/yyyy") & "')," & _
                      gLetras.Columns.ColumnByFieldName("Porcentaje").Value & "," & gLetras.Columns.ColumnByFieldName("Capital").Value & "," & _
                      gLetras.Columns.ColumnByFieldName("Portes").Value & "," & gLetras.Columns.ColumnByFieldName("Interes").Value & "," & _
                      gLetras.Columns.ColumnByFieldName("Total").Value & "," & gLetras.Columns.ColumnByFieldName("Total").Value & _
                      ",'','" & cmoneda & "'," & gLetras.Columns.ColumnByFieldName("IGV").Value & _
                      ",CVDATE('" & Format(dtp_Emision.Value, "DD/MM/YYYY") & "'),'" & ccliente & "','" & txtGls_Cliente.Text & "','" & _
                      strRUCCliente & "'," & Txt_TipoCambio.Value & "," & txtVal_Tasa.Value & ",'X','',''," & _
                      gLetras.Columns.ColumnByFieldName("Total").Value & ",'')"
            cnn_empresa.Execute (cinsert)
            
            cupdate = "UPDATE AUTOGENERA SET NRO_LET = '" & strNumLetra & "'"
            cnn_empresa.Execute (cupdate)
            
        End If
        '---------------------------------------------------------------------------------------------------------------
        gLetras.Dataset.Edit
        gLetras.Columns.ColumnByFieldName("idLetra").Value = strNumLetra
        strNumLetra = Format(Val(strNumLetra + 1), "000000")
        gLetras.Dataset.Next
    Loop
    
    csql = "UPDATE docventas SET indLetra = 'S' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idSerie = '" & strSerie & "' AND idDocVentas = '" & strNumDoc & "'"
    Cn.Execute csql
    sw = "1"
    Cn.CommitTrans
    
    MsgBox "Letras Registradas Satisfactoriamente", vbInformation, App.Title
    
    Exit Sub

Err:
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Modificar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strAno        As String
Dim strMoneda     As String
Dim strNumLetra   As String
Dim strRUCCliente   As String
Dim indTrans As Boolean
Dim cnn_empresa     As New ADODB.Connection
Dim rsbusca         As New ADODB.Recordset
Dim cconex_empresa  As String, cdirectorio As String, cruta As String, ccliente As String, cbusca As String
Dim cinsert         As String, cmoneda As String, cupdate  As String

    indTrans = False
    strAno = Year(dtp_Emision.Value)
    strMoneda = txtCod_Moneda.Text
    strRUCCliente = traerCampo("personas", "RUC", "idPersona", txtCod_Cliente.Text, False)
    
    If gLetras.Dataset.State = dsEdit Or gLetras.Dataset.State = dsInsert Then gLetras.Dataset.Post
    
    indTrans = True
    Cn.BeginTrans
    
    gLetras.Dataset.First
    Do While Not gLetras.Dataset.EOF
        strNumLetra = gLetras.Columns.ColumnByFieldName("idLetra").Value
        
        csql = "delete from Letras where idLetra = '" & strNumLetra & "' and idEmpresa = '" & glsEmpresa & "'  and idSucursal = '" & glsSucursal & "' and idSerie = '0' and idCliente = '" & txtCod_Cliente.Text & "' and Item = " & gLetras.Columns.ColumnByFieldName("Item").Value & " and idAno = '" & strAno & "' "
        Cn.Execute csql
        
        csql = "INSERT INTO letras(idEmpresa,idSucursal,idAno,idLetra,idSerie,GlsDocReferencia,Plazo,FecVencimiento," & _
            " Porcentaje,Capital,Portes,Interes,TotalLetra,MontoOriginal,idMoneda, " & _
            " IGVIntereses,FecEmision,idCliente,GlsCliente,RUCCliente,TipoCambio,TasaLetra, " & _
            " TotalDocumento,idUsuarioReg,Item,idPeriodoLetra,idTipoInteres,idTipoCuotaLetra,idTipoCapitalizacion,FecRegistro,idAval,Situacion)"
            
        csql = csql & " VALUES('" & glsEmpresa & "','" & glsSucursal & "','" & strAno & "','" & strNumLetra & "',0,'" & (strTD & "-" & strSerie & "-" & strNumDoc) & "'," & _
             gLetras.Columns.ColumnByFieldName("Plazo").Value & ",'" & Format(gLetras.Columns.ColumnByFieldName("FecVcto").Value, "yyyy-mm-dd") & "'," & gLetras.Columns.ColumnByFieldName("Porcentaje").Value & "," & gLetras.Columns.ColumnByFieldName("Capital").Value & "," & _
             gLetras.Columns.ColumnByFieldName("Portes").Value & "," & gLetras.Columns.ColumnByFieldName("Interes").Value & "," & gLetras.Columns.ColumnByFieldName("Total").Value & "," & gLetras.Columns.ColumnByFieldName("Total").Value & ",'" & txtCod_Moneda.Text & "'," & _
             gLetras.Columns.ColumnByFieldName("IGV").Value & ",'" & Format(dtp_EmisionLetra.Value, "yyyy-mm-dd") & "','" & txtCod_Cliente.Text & "','" & txtGls_Cliente.Text & "','" & strRUCCliente & "'," & _
             Txt_TipoCambio.Value & "," & txtVal_Tasa.Value & "," & txt_TotalNeto.Value & ",'" & glsUser & "'," & gLetras.Columns.ColumnByFieldName("Item").Value & ",'" & _
             txtCod_PeriodoLetra.Text & "','" & txtCod_TipoInteres.Text & "','" & txtCod_TipoCuotaLetra.Text & "','" & txtCod_TipoCapitalizacion.Text & "',sysdate(),'" & txtCod_Aval.Text & "','X')"
             
        Cn.Execute csql
        
        '---------------------------------------------------------------------------------------------------------------
        cdirectorio = traerCampo("empresas", "Carpeta", "idEmpresa", glsEmpresa, False)
        If cdirectorio <> "" And glsSistemaAccess = "S" Then 'Grabamos en la Version de Access
            cruta = glsRuta_Access & cdirectorio
            If cnn_empresa.State = adStateOpen Then cnn_empresa.Close
            cconex_empresa = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cruta & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
            cnn_empresa.Open cconex_empresa
    
            '--------------------------------
            '--- CLIENTE
            ccliente = ""
            cbusca = "SELECT F2CODCLI FROM EF2CLIENTES WHERE F2NEWRUC = '" & strRUCCliente & "'"
            If rsbusca.State = adStateOpen Then rsbusca.Close
            rsbusca.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
            If Not rsbusca.EOF Then
                ccliente = Trim(rsbusca.Fields("F2CODCLI") & "")
            End If
            rsbusca.Close: Set rsbusca = Nothing
            '--------------------------------
            '---- AGREGA CLIENTE
            If Len(ccliente) = 0 Then
                ccliente = ""
                cbusca = "SELECT F2CODCLI FROM EF2CLIENTES ORDER BY F2CODCLI DESC"
                If rsbusca.State = adStateOpen Then rsbusca.Close
                rsbusca.Open cbusca, cnn_empresa, adOpenKeyset, adLockOptimistic
                If Not rsbusca.EOF Then
                    ccliente = Format(Val(rsbusca.Fields("F2CODCLI") & "") + 1, "0000")
                End If
                rsbusca.Close: Set rsbusca = Nothing
                cinsert = "INSERT INTO EF2CLIENTES " & _
                          "(F2CODCLI,F2NOMCLI,F2NEWRUC,F2DIRCLI,F2TIPDOC) " & _
                          "VALUES ('" & ccliente & "','" & txtGls_Cliente.Text & "','" & strRUCCliente & "','" & cdirCliente & "','J')"
                cnn_empresa.Execute (cinsert)
            End If
            '--------------------------------
            cmoneda = IIf(txtCod_Moneda.Text = "PEN", "S", "D")
            
            cinsert = "delete from Letras where NRO_LETRA = '" & strNumLetra & "' and SER_LETRA = '0' and CLIENTE = '" & ccliente & "' and ANO_LETRA = '" & strAno & "' "
            cnn_empresa.Execute (cinsert)
            
            cinsert = "INSERT INTO LETRAS " & _
                      "(ANO_LETRA,NRO_LETRA,SER_LETRA,NRO_REF,PLAZO,FCH_VENC," & _
                      "PORCENTAJE,CAPITAL,PORTES,INTERES,TOT_LETRA,MONTO_ORIGINAL," & _
                      "TIP_DCTO,TIP_MONE,IGV_INT,FCH_EMIS,CLIENTE,NOMCLI,ruccli,TIP_CAMB,TASA_INT,SITUACION,tip_inte,NRO_DCTO,TOT_DCTO,usuario) " & _
                      "VALUES ('" & strAno & "','" & strNumLetra & "','0',''," & gLetras.Columns.ColumnByFieldName("Plazo").Value & _
                      ",CVDate('" & Format(gLetras.Columns.ColumnByFieldName("FecVcto").Value, "dd/mm/yyyy") & "')," & _
                      gLetras.Columns.ColumnByFieldName("Porcentaje").Value & "," & gLetras.Columns.ColumnByFieldName("Capital").Value & "," & _
                      gLetras.Columns.ColumnByFieldName("Portes").Value & "," & gLetras.Columns.ColumnByFieldName("Interes").Value & "," & _
                      gLetras.Columns.ColumnByFieldName("Total").Value & "," & gLetras.Columns.ColumnByFieldName("Total").Value & _
                      ",'','" & cmoneda & "'," & gLetras.Columns.ColumnByFieldName("IGV").Value & _
                      ",CVDATE('" & Format(dtp_Emision.Value, "DD/MM/YYYY") & "'),'" & ccliente & "','" & txtGls_Cliente.Text & "','" & _
                      strRUCCliente & "'," & Txt_TipoCambio.Value & "," & txtVal_Tasa.Value & ",'X','',''," & _
                      gLetras.Columns.ColumnByFieldName("Total").Value & ",'')"
            cnn_empresa.Execute (cinsert)
            
            cupdate = "UPDATE AUTOGENERA SET NRO_LET = '" & strNumLetra & "'"
            cnn_empresa.Execute (cupdate)
        End If
        '---------------------------------------------------------------------------------------------------------------
        
        gLetras.Dataset.Edit
        gLetras.Columns.ColumnByFieldName("idLetra").Value = strNumLetra
        strNumLetra = Format(Val(strNumLetra + 1), "000000")
        gLetras.Dataset.Next
    Loop
    
    csql = "UPDATE docventas SET indLetra = 'S' WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTD & "' AND idSerie = '" & strSerie & "' AND idDocVentas = '" & strNumDoc & "'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    MsgBox "Letras Modificada Satisfactoriamente", vbInformation, App.Title
    
    Exit Sub

Err:
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_PeriodoLetra_Change()
    
    txtGls_PeriodoLetra.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_PeriodoLetra.Text, False)

End Sub

Private Sub txtCod_PeriodoLetra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PERIODOLETRA", txtCod_PeriodoLetra, txtGls_PeriodoLetra
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_TipoInteres_Change()
    
    txtGls_TipoInteres.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_TipoInteres.Text, False)

End Sub

Private Sub txtCod_TipoInteres_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "TIPOINTERES", txtCod_TipoInteres, txtGls_TipoInteres
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_TipoCuotaLetra_Change()
    
    txtGls_TipoCuotaLetra.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_TipoCuotaLetra.Text, False)

End Sub

Private Sub txtCod_TipoCuotaLetra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "TIPOCUOTALETRA", txtCod_TipoCuotaLetra, txtGls_TipoCuotaLetra
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_TipoCapitalizacion_Change()
    
    txtGls_TipoCapitalizacion.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_TipoCapitalizacion.Text, False)

End Sub

Private Sub txtCod_TipoCapitalizacion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "TIPOCAPITALIZACION", txtCod_TipoCapitalizacion, txtGls_TipoCapitalizacion
        KeyAscii = 0
    End If

End Sub

Public Sub CalculaIntereses(ByRef dblInteresTotal As Double, ByRef dblInteresIGV As Double, ByRef StrMsgError As String)
On Error GoTo Err
Dim intDias     As Integer
Dim dblInteres  As Double

'PERIODO LETRA
'09001 - ANUAL
'09002 - MENSUAL

'TIPO INTERES
'10001 - EFECTIVO
'10002 - NOMINAL
'10003 - SIMPLE

'TIPO CUOTA LETRA
'11001 - CRECIENTE

'TIPO CAPITALIZACION
'12001 - MENSUAL
'12002 - DIARIO

    If txtCod_TipoCapitalizacion.Text = "12001" Then  'si  capitalización es mensual
        If txtCod_PeriodoLetra.Text = "09001" Then  ' periodo anual
            Select Case txtCod_TipoInteres.Text    'tipo de interes
                Case "10002"    ' nominal
                   dblInteres = Val(txtVal_Tasa.Value) / 1200
                Case "10001"       'efectivo
                   dblInteres = (((Val(txtVal_Tasa.Value) / 100) + 1) ^ (1 / 12)) - 1
                Case Else
                   dblInteres = Val(txtVal_Tasa.Value) / 1200
            End Select
        Else
            Select Case txtCod_TipoInteres.Text  'tipo de interes
                Case "10002"    ' nominal
                   dblInteres = Val(txtVal_Tasa.Value) / 100
                Case "10001"    'efectivo
                   dblInteres = (Val(txtVal_Tasa.Value)) / 100
                Case Else
                   dblInteres = Val(txtVal_Tasa.Value) / 100
            End Select
        End If
        
        Select Case txtCod_TipoCuotaLetra.Text 'cuota
            Case "11001"        'creciente
                If txtCod_TipoInteres.Text = "10003" Then      ' tipo de interes simple
                    dblInteresTotal = ((Val(Format(gLetras.Columns.ColumnByFieldName("Capital").Value, "0.00")) * (1 + (dblInteres * (Val(gLetras.Columns.ColumnByFieldName("Plazo").Value) / IIf(txtCod_PeriodoLetra.Text = "09001", 360, 30)))))) - Val(Format(gLetras.Columns.ColumnByFieldName("Capital").Value, "0.00"))
                Else    'efectivo o nominal
                    dblInteresTotal = Val(Format(gLetras.Columns.ColumnByFieldName("Capital").Value, "0.00")) * (1 + dblInteres) ^ (Val(gLetras.Columns.ColumnByFieldName("Plazo").Value) / 30) - Val(Format(gLetras.Columns.ColumnByFieldName("Capital").Value, "0.00"))
                End If
            Case "11002"    'rebatir
'                If SALREBAT = 0 Then
'                    SALREBAT = monto
'                Else
'                    SALREBAT = SALREBAT - ACUMLET
'                End If
'                If txtCod_TipoInteres.Text = "10003" Then    'tipo de int. simple
'                   dblInteresTotal = SALREBAT * (1 + (dblInteres * ((Val(gLetras.Columns.ColumnByFieldName("Plazo").Value) - PLAZOANT) / IIf(peri = "A", 360, 30)))) - SALREBAT
'                Else      'efectivo o nominal
'                   dblInteresTotal = SALREBAT * (1 + dblInteres) ^ ((gLetras.Columns.ColumnByFieldName("Plazo").Value - PLAZOANT) / 30) - SALREBAT
'                End If
'                ACUMLET = Val(gLetras.Columns.ColumnByFieldName("CAPITAL").Value)
'                PLAZOANT = Val(gLetras.Columns.ColumnByFieldName("PLAZO").Value)
        End Select
    Else 'si  capitalización es diaria
        If txtCod_PeriodoLetra.Text = "09001" And txtCod_TipoInteres.Text = "10002" Then    'si periodo es Anual nominal
            dblInteres = Val(txtVal_Tasa.Value) / 36000
        Else
            dblInteres = Val(txtVal_Tasa.Value) / 3000
        End If
        If txtCod_TipoCuotaLetra.Text = "11001" Then    ' si cuota es creciente
            dblInteresTotal = Val(gLetras.Columns.ColumnByFieldName("Capital").Value) * (1 + dblInteres) ^ (gLetras.Columns.ColumnByFieldName("Plazo").Value) - Val(gLetras.Columns.ColumnByFieldName("Capital").Value)
        Else
            dblInteresTotal = Val(gLetras.Columns.ColumnByFieldName("Capital").Value) * (1 + dblInteres) ^ (gLetras.Columns.ColumnByFieldName("Plazo").Value - gLetras.Columns.ColumnByFieldName("Plazo").Value) - Val(gLetras.Columns.ColumnByFieldName("Capital").Value)
        End If
    End If
    
    dblInteresIGV = dblInteresTotal * (glsTC / 100)

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub CalculoTotalLetraFila()
    
    gLetras.Columns.ColumnByFieldName("Total").Value = gLetras.Columns.ColumnByFieldName("Capital").Value + gLetras.Columns.ColumnByFieldName("Portes").Value + gLetras.Columns.ColumnByFieldName("Interes").Value + Val(Format(gLetras.Columns.ColumnByFieldName("IGV").Value, "0.00"))

End Sub

Public Sub CalculaPorcentaje(ByRef StrMsgError As String)    'calculo de porcentaje de la ultima fila
On Error GoTo Err
Dim dblTotalPorcentaje   As Double
Dim dblTotal            As Double

    gLetras.Dataset.Edit
    gLetras.Columns.ColumnByFieldName("Porcentaje").Value = "0.00"
    gLetras.Columns.ColumnByFieldName("Total").Value = "0.00"
    gLetras.Columns.ColumnByFieldName("IGV").Value = "0.00"
    gLetras.Dataset.Post
    
    dblTotalPorcentaje = Val(gLetras.Columns.ColumnByFieldName("Porcentaje").SummaryFooterValue & "")
    dblTotal = Val(gLetras.Columns.ColumnByFieldName("Total").SummaryFooterValue & "")

    gLetras.Dataset.Edit
    gLetras.Columns.ColumnByFieldName("Porcentaje").Value = Format(100 - (dblTotalPorcentaje), "###,##0.00")
    gLetras.Columns.ColumnByFieldName("Capital").Value = Format(txt_MontoLetras.Value - dblTotal, "###,##0.00")

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub IMPRIME_LETRA_GENERAL(strAno As String, strNumLetra As String, strSerieLetra As String)
Dim csql As String
Dim rstletra As New ADODB.Recordset
Dim strimporteletras As String
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strdirec  As String
Dim strtelef As String
Dim StrMsgError As String
Dim straval As String
Dim strdirecaval As String
Dim strabrev As String
Dim strcoddis As String
Dim strnomdis As String
Dim strcoddisaval As String
Dim strnomdisaval As String
Dim impresoraletra  As String
    
    strAno = Year(dtp_Emision.Value)
    impresoraletra = traerCampo("usuarios", "ImpresoraLetras", "idUsuario", glsUser, True)
    
    If impresoraletra <> "" Then
        For Each p In Printers
           If UCase(p.DeviceName) = impresoraletra Then
              Set Printer = p
              indPrinter = True
              Exit For
           End If
        Next p
    End If
    
    strNumLetra = Format(strNumLetra, "000000")
    intScale = 6
    Printer.ScaleMode = intScale
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    
    csql = "SELECT idletra,glsDocreferencia,FecEmision,FecVencimiento,TotalLetra,CASE idMoneda WHEN 'PEN' THEN 'S/.' ELSE 'US$' END AS idMoneda,idCliente,glsCliente,RucCliente,idaval FROM LETRAS  WHERE IDEMPRESA= '" & glsEmpresa & "' AND IDLETRA = '" & strNumLetra & "' AND IDSERIE = '" & strSerieLetra & "' and idAno = '" & Trim("" & strAno) & "' "
    rstletra.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstletra.EOF Then
        '--- NUMERO LETRA
        ImprimeXY rstletra.Fields("idLetra") & "", "T", 20, 15, 46, 0, 0, StrMsgError
        strabrev = traerCampo("documentos", "abredocumento", "iddocumento", left(rstletra.Fields("glsDocreferencia"), 2), False)
        '--- REFERENCIA
        ImprimeXY strabrev & " " & right(rstletra.Fields("glsDocreferencia"), 12) & "", "T", 20, 15, 65, 0, 0, StrMsgError
        '--- LUGAR DE GIRO
        ImprimeXY "LIMA", "T", 20, 15, 128, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO EMISION
        ImprimeXY Format(Year(rstletra.Fields("FecEmision")), "0000") & "", "T", 4, 15, 90 + 18, 0, 0, StrMsgError
        '--- IMPRIME EL MES EMISION
        ImprimeXY Format(Month(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 15, 90 + 11, 0, 0, StrMsgError
        '--- IMPRIME EL DIA EMISION
        ImprimeXY Format(Day(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 15, 93, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO VENCIMIENTO
        ImprimeXY Format(Year(rstletra.Fields("FecVencimiento")), "0000") & "", "T", 4, 15, 148 + 18, 0, 0, StrMsgError
        '--- IMPRIME EL MES VENCIMIENTO
        ImprimeXY Format(Month(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 15, 148 + 11, 0, 0, StrMsgError
        '--- IMPRIME EL DIA VENCIMIENTO
        ImprimeXY Format(Day(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 15, 151, 0, 0, StrMsgError
        '--- Signo de la modeda
        ImprimeXY rstletra.Fields("idMoneda") & "", "T", 3, 15, 177, 2, 0, StrMsgError
        '--- IMPORTE
        ImprimeXY rstletra.Fields("TotalLetra") & "", "N", 20, 15, 182, 2, 0, StrMsgError
        '--- IMPORTE EN LETRAS
        strimporteletras = "SON:" & Cadenanum(Format(rstletra.Fields("TotalLetra"), "0.00"), IIf(rstletra.Fields("idMoneda") = "S/.", "NUEVOS SOLES", "DOLARES AMERICANOS"))
        ImprimeXY strimporteletras & "", "T", 250, 36, 41, 0, 0, StrMsgError
        '--- CLIENTE
        ImprimeXY rstletra.Fields("glsCliente") & "", "T", 36, 46, 52, 0, 0, StrMsgError
        '--- RUC
        ImprimeXY rstletra.Fields("RucCliente") & "", "T", 153, 64, 58, 0, 0, StrMsgError
        '--- LOCALIDAD
        strcoddis = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idCliente"), False)
        strnomdis = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddis, False)
        ImprimeXY strnomdis & "", "T", 150, 59, 86, 0, 0, StrMsgError
        
        strdirec = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idCliente"), False)
        strtelef = traerCampo("personas", "telefonos", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 6
        ImprimeXY left(strdirec, 120) & "", "T", 145, 55, 52, 0, 0, StrMsgError
        Printer.FontSize = 8
        ImprimeXY left(strtelef, 12) & "", "T", 153, 64, 93, 0, 0, StrMsgError
        
        If rstletra.Fields("idAval") <> "" Then
            straval = traerCampo("personas", "glspersona", "idpersona", rstletra.Fields("idaval"), False)
            strdirecaval = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY straval & "", "T", 150, 72, 60, 0, 0, StrMsgError
            ImprimeXY left(strdirecaval, 50) & "", "T", 150, 78, 54, 0, 0, StrMsgError
            
            strcoddisaval = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idaval"), False)
            strnomdisaval = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddisaval, False)
            ImprimeXY strnomdisaval & "", "T", 150, 80, 55, 0, 0, StrMsgError
        End If
        Printer.Print Chr$(149)
        Printer.Print ""
        Printer.Print ""
        Printer.EndDoc
    End If

End Sub

Private Sub IMPRIME_LETRA_SINCHI(strAno As String, strNumLetra As String, strSerieLetra As String)
Dim csql As String
Dim rstletra As New ADODB.Recordset
Dim strimporteletras As String
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strdirec  As String
Dim strtelef As String
Dim StrMsgError As String
Dim straval As String
Dim strdirecaval As String
Dim strTelefonoAval As String
Dim strRUCAval As String
Dim strabrev As String
Dim strcoddis As String
Dim strnomdis As String
Dim strcoddisaval As String
Dim strnomdisaval As String
Dim impresoraletra  As String
Dim localidad As String

    strAno = Year(dtp_Emision.Value)
    impresoraletra = traerCampo("usuarios", "ImpresoraLetras", "idUsuario", glsUser, True)
    
    If impresoraletra <> "" Then
        For Each p In Printers
           If UCase(p.DeviceName) = impresoraletra Then
              Set Printer = p
              indPrinter = True
              Exit For
           End If
        Next p
    End If
    
    strNumLetra = Format(strNumLetra, "000000")
    intScale = 6
    Printer.ScaleMode = intScale
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    
    csql = "SELECT idletra,glsDocreferencia,FecEmision,FecVencimiento,TotalLetra," & _
           "CASE idMoneda WHEN 'PEN' THEN 'S/.' ELSE 'US$' END AS idMoneda,idCliente,glsCliente,RucCliente,idaval " & _
           "FROM LETRAS  " & _
           "WHERE IDEMPRESA = '" & glsEmpresa & "' " & _
           "AND IDLETRA = '" & strNumLetra & "' " & _
           "AND IDSERIE = '" & strSerieLetra & "' and idAno = '" & Trim("" & strAno) & "' "
    rstletra.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstletra.EOF Then
        '--- NUMERO LETRA
        ImprimeXY rstletra.Fields("idLetra") & "", "T", 20, 18, 48, 0, 0, StrMsgError
        strabrev = traerCampo("documentos", "abredocumento", "iddocumento", left(rstletra.Fields("glsDocreferencia"), 2), False)
        '--- REFERENCIA
        ImprimeXY strabrev & " " & right(rstletra.Fields("glsDocreferencia"), 12) & "", "T", 20, 18, 67, 0, 0, StrMsgError
        '--- LUGAR DE GIRO
        localidad = traerCampo("personas", "glsPersona", "idPersona", glsSucursal, False)
        ImprimeXY localidad, "T", 20, 18, 128, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO EMISION
        ImprimeXY Format(Year(rstletra.Fields("FecEmision")), "0000") & "", "T", 4, 18, 91 + 18, 0, 0, StrMsgError
        '--- IMPRIME EL MES EMISION
        ImprimeXY Format(Month(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 18, 91 + 11, 0, 0, StrMsgError
        '--- IMPRIME EL DIA EMISION
        ImprimeXY Format(Day(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 18, 94, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO VENCIMIENTO
        ImprimeXY Format(Year(rstletra.Fields("FecVencimiento")), "0000") & "", "T", 4, 18, 145 + 18, 0, 0, StrMsgError
        '--- IMPRIME EL MES VENCIMIENTO
        ImprimeXY Format(Month(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 18, 145 + 11, 0, 0, StrMsgError
        '--- IMPRIME EL DIA VENCIMIENTO
        ImprimeXY Format(Day(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 18, 148, 0, 0, StrMsgError
        '--- Signo de la modeda
        ImprimeXY rstletra.Fields("idMoneda") & "", "T", 3, 18, 175, 2, 0, StrMsgError
        '--- IMPORTE
        ImprimeXY rstletra.Fields("TotalLetra") & "", "N", 20, 18, 175, 2, 0, StrMsgError
        '--- IMPORTE EN LETRAS
        strimporteletras = "SON:" & Cadenanum(Format(rstletra.Fields("TotalLetra"), "0.00"), IIf(rstletra.Fields("idMoneda") = "S/.", "NUEVOS SOLES", "DOLARES AMERICANOS"))
        ImprimeXY Mid(strimporteletras, 5, Len(strimporteletras) - 3) & "", "T", 250, 32, 58, 0, 0, StrMsgError
        '--- CLIENTE
        ImprimeXY rstletra.Fields("glsCliente") & "", "T", 50, 45, 54, 0, 0, StrMsgError
        '--- RUC
        ImprimeXY rstletra.Fields("RucCliente") & "", "T", 153, 62, 55, 0, 0, StrMsgError
        '--- LOCALIDAD
        strcoddis = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idCliente"), False)
        strcoddis = left(strcoddis, 4) & "00"
        strnomdis = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddis, False)
        ImprimeXY strnomdis & "", "T", 150, 57, 102, 0, 0, StrMsgError
        
        strdirec = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idCliente"), False)
        strtelef = traerCampo("personas", "telefonos", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 6
        ImprimeXY left(strdirec, 120) & "", "T", 145, 53, 52, 0, 0, StrMsgError
        Printer.FontSize = 8
        ImprimeXY left(strtelef, 12) & "", "T", 153, 62, 102, 0, 0, StrMsgError
        
        If rstletra.Fields("idAval") <> "" Then
            straval = traerCampo("personas", "glspersona", "idpersona", rstletra.Fields("idaval"), False)
            strdirecaval = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY straval & "", "T", 150, 71, 63, 0, 0, StrMsgError
            ImprimeXY left(strdirecaval, 50) & "", "T", 150, 77, 53, 0, 0, StrMsgError
            
            strcoddisaval = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idaval"), False)
            strnomdisaval = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddisaval, False)
            ImprimeXY strnomdisaval & "", "T", 150, 84, 53, 0, 0, StrMsgError
            
            strTelefonoAval = traerCampo("personas", "Telefonos", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY left(strTelefonoAval, 12) & "", "T", 153, 83, 102, 0, 0, StrMsgError
            
            strRUCAval = traerCampo("personas", "RUC", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY strRUCAval, "T", 153, 90, 52, 0, 0, StrMsgError
        End If
        
        Printer.Print Chr$(149)
        Printer.Print ""
        Printer.Print ""
        Printer.EndDoc
    End If

End Sub

Private Sub IMPRIME_LETRA_PIC(strAno As String, strNumLetra As String, strSerieLetra As String)
Dim csql As String
Dim rstletra As New ADODB.Recordset
Dim strimporteletras As String
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strdirec  As String
Dim strtelef As String
Dim StrMsgError As String
Dim straval As String
Dim strdirecaval As String
Dim strTelefonoAval As String
Dim strRUCAval As String
Dim strabrev As String
Dim strcoddis As String
Dim strnomdis As String
Dim strcoddisaval As String
Dim strnomdisaval As String
Dim impresoraletra  As String
Dim localidad As String

    strAno = Year(dtp_Emision.Value)
    impresoraletra = traerCampo("usuarios", "ImpresoraLetras", "idUsuario", glsUser, True)
    
    If impresoraletra <> "" Then
        For Each p In Printers
           If UCase(p.DeviceName) = impresoraletra Then
              Set Printer = p
              indPrinter = True
              Exit For
           End If
        Next p
    End If
    
    strNumLetra = Format(strNumLetra, "000000")
    intScale = 6
    Printer.ScaleMode = intScale
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    
    csql = "SELECT idletra,glsDocreferencia,FecEmision,FecVencimiento,TotalLetra," & _
           "CASE idMoneda WHEN 'PEN' THEN 'S/.' ELSE 'US$' END AS idMoneda,idCliente,glsCliente,RucCliente,idaval " & _
           "FROM LETRAS  " & _
           "WHERE IDEMPRESA = '" & glsEmpresa & "' " & _
           "AND IDLETRA = '" & strNumLetra & "' " & _
           "AND IDSERIE = '" & strSerieLetra & "' and idAno = '" & Trim("" & strAno) & "' "
    rstletra.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstletra.EOF Then
        '--- NUMERO LETRA
        ImprimeXY rstletra.Fields("idLetra") & "", "T", 20, 18, 48, 0, 0, StrMsgError
        strabrev = traerCampo("documentos", "abredocumento", "iddocumento", left(rstletra.Fields("glsDocreferencia"), 2), False)
        '--- REFERENCIA
        ImprimeXY strabrev & " " & right(rstletra.Fields("glsDocreferencia"), 12) & "", "T", 20, 18, 67, 0, 0, StrMsgError
        '--- LUGAR DE GIRO
        localidad = traerCampo("personas", "glsPersona", "idPersona", glsSucursal, False)
        ImprimeXY localidad, "T", 20, 18, 128, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO EMISION
        ImprimeXY Format(Year(rstletra.Fields("FecEmision")), "0000") & "", "T", 4, 18, 91 + 18, 0, 0, StrMsgError
        '--- IMPRIME EL MES EMISION
        ImprimeXY Format(Month(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 18, 91 + 11, 0, 0, StrMsgError
        '--- IMPRIME EL DIA EMISION
        ImprimeXY Format(Day(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 18, 94, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO VENCIMIENTO
        ImprimeXY Format(Year(rstletra.Fields("FecVencimiento")), "0000") & "", "T", 4, 18, 145 + 18, 0, 0, StrMsgError
        '--- IMPRIME EL MES VENCIMIENTO
        ImprimeXY Format(Month(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 18, 145 + 11, 0, 0, StrMsgError
        '--- IMPRIME EL DIA VENCIMIENTO
        ImprimeXY Format(Day(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 18, 148, 0, 0, StrMsgError
        '--- Signo de la modeda
        ImprimeXY rstletra.Fields("idMoneda") & "", "T", 3, 18, 175, 2, 0, StrMsgError
        '--- IMPORTE
        ImprimeXY rstletra.Fields("TotalLetra") & "", "N", 20, 18, 175, 2, 0, StrMsgError
        '--- IMPORTE EN LETRAS
        strimporteletras = "SON:" & Cadenanum(Format(rstletra.Fields("TotalLetra"), "0.00"), IIf(rstletra.Fields("idMoneda") = "S/.", "NUEVOS SOLES", "DOLARES AMERICANOS"))
        ImprimeXY Mid(strimporteletras, 5, Len(strimporteletras) - 3) & "", "T", 250, 32, 58, 0, 0, StrMsgError
        '--- CLIENTE
        ImprimeXY rstletra.Fields("glsCliente") & "", "T", 50, 44, 54, 0, 0, StrMsgError
        '--- RUC
        ImprimeXY rstletra.Fields("RucCliente") & "", "T", 153, 62, 55, 0, 0, StrMsgError
        '--- LOCALIDAD
        strcoddis = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idCliente"), False)
        strcoddis = left(strcoddis, 4) & "00"
        strnomdis = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddis, False)
        ImprimeXY strnomdis & "", "T", 150, 56, 102, 0, 0, StrMsgError
        
        strdirec = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idCliente"), False)
        strtelef = traerCampo("personas", "telefonos", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 6
        ImprimeXY left(strdirec, 120) & "", "T", 145, 52, 52, 0, 0, StrMsgError
        Printer.FontSize = 8
        ImprimeXY left(strtelef, 12) & "", "T", 153, 61, 102, 0, 0, StrMsgError
        
        If rstletra.Fields("idAval") <> "" Then
            straval = traerCampo("personas", "glspersona", "idpersona", rstletra.Fields("idaval"), False)
            strdirecaval = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY straval & "", "T", 150, 71, 63, 0, 0, StrMsgError
            ImprimeXY left(strdirecaval, 50) & "", "T", 150, 77, 53, 0, 0, StrMsgError
            
            strcoddisaval = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idaval"), False)
            strnomdisaval = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddisaval, False)
            ImprimeXY strnomdisaval & "", "T", 150, 84, 53, 0, 0, StrMsgError
            
            strTelefonoAval = traerCampo("personas", "Telefonos", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY left(strTelefonoAval, 12) & "", "T", 153, 83, 102, 0, 0, StrMsgError
            
            strRUCAval = traerCampo("personas", "RUC", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY strRUCAval, "T", 153, 90, 52, 0, 0, StrMsgError
        End If
        Printer.Print Chr$(149)
        Printer.Print ""
        Printer.Print ""
        Printer.EndDoc
    End If

End Sub

Private Sub IMPRIME_LETRA_ALSISAC(strAno As String, strNumLetra As String, strSerieLetra As String)
Dim csql                As String
Dim rstletra            As New ADODB.Recordset
Dim strimporteletras    As String
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strdirec            As String
Dim strtelef            As String
Dim StrMsgError         As String
Dim straval             As String
Dim strdirecaval        As String
Dim strTelefonoAval     As String
Dim strRUCAval          As String
Dim strabrev            As String
Dim strcoddis           As String
Dim strnomdis           As String
Dim strcoddisaval       As String
Dim strnomdisaval       As String
Dim impresoraletra      As String
Dim localidad           As String
Dim strcadena           As String

    strAno = Year(dtp_Emision.Value)
    impresoraletra = traerCampo("usuarios", "ImpresoraLetras", "idUsuario", glsUser, True)
    
    If impresoraletra <> "" Then
        For Each p In Printers
            If UCase(p.DeviceName) = impresoraletra Then
                Set Printer = p
                indPrinter = True
                Exit For
            End If
        Next p
    End If
    
    strNumLetra = Format(strNumLetra, "000000")
    intScale = 6
    Printer.ScaleMode = intScale
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    
    csql = "SELECT idletra,glsDocreferencia,FecEmision,FecVencimiento,TotalLetra," & _
           "CASE idMoneda WHEN 'PEN' THEN 'S/.' ELSE 'US$' END AS idMoneda,idCliente,glsCliente,RucCliente,idaval " & _
           "FROM LETRAS  " & _
           "WHERE IDEMPRESA = '" & glsEmpresa & "' " & _
           "AND IDLETRA = '" & strNumLetra & "' " & _
           "AND IDSERIE = '" & strSerieLetra & "' and idAno = '" & Trim("" & strAno) & "' "
    rstletra.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstletra.EOF Then
        '--- NUMERO LETRA
        ImprimeXY rstletra.Fields("idLetra") & "", "T", 20, 22, 48, 0, 0, StrMsgError
        '--- REFERENCIA
        ImprimeXY right(rstletra.Fields("glsDocreferencia"), 12) & "", "T", 20, 22, 63, 0, 0, StrMsgError
        '--- LUGAR DE GIRO
        ImprimeXY "LIMA", "T", 20, 22, 126, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO EMISION
        ImprimeXY Format(Year(rstletra.Fields("FecEmision")), "0000") & "", "T", 4, 22, 64 + 46, 0, 0, StrMsgError
        '--- IMPRIME EL MES EMISION
        ImprimeXY Format(Month(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 22, 64 + 38, 0, 0, StrMsgError
        '--- IMPRIME EL DIA EMISION
        ImprimeXY Format(Day(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 22, 64 + 30, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO VENCIMIENTO
        ImprimeXY Format(Year(rstletra.Fields("FecVencimiento")), "0000") & "", "T", 4, 22, 93 + 71, 0, 0, StrMsgError
        '--- IMPRIME EL MES VENCIMIENTO
        ImprimeXY Format(Month(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 22, 93 + 64, 0, 0, StrMsgError
        '--- IMPRIME EL DIA VENCIMIENTO
        ImprimeXY Format(Day(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 22, 93 + 56, 0, 0, StrMsgError
        '--- Signo de la modeda
        ImprimeXY rstletra.Fields("idMoneda") & "", "T", 3, 22, 176, 2, 0, StrMsgError
        '--- IMPORTE
        ImprimeXY rstletra.Fields("TotalLetra") & "", "N", 18, 22, 175, 2, 0, StrMsgError
        '--- IMPORTE EN LETRAS
        strimporteletras = MonedaTexto(Format(rstletra.Fields("TotalLetra"), "0.00"), IIf(rstletra.Fields("idMoneda") = "S/.", "0", "1"))
        ImprimeXY UCase(strimporteletras) & "", "T", 250, 35, 40, 0, 0, StrMsgError
        '--- CLIENTE
        ImprimeXY rstletra.Fields("glsCliente") & "", "T", 153, 45, 55, 0, 0, StrMsgError
        '--- RUC
        ImprimeXY rstletra.Fields("RucCliente") & "", "T", 153, 59, 53, 0, 0, StrMsgError
        '--- LOCALIDAD
        strcoddis = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idCliente"), False)
        strcoddis = left(strcoddis, 4) & "00"
        strnomdis = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddis, False)
        strdirec = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idCliente"), False)
        
        strtelef = traerCampo("personas", "telefonos", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 6
        strcadena = UCase(left(strdirec, 120) & Space(2) & strnomdis) & ""
        
        If left(strcadena, 45) >= 45 Then
            ImprimeXY left(strcadena, 44), "T", 145, 50, 55, 0, 0, StrMsgError
        End If
        
        strcoddis = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idCliente"), False)
        strcoddis = left(strcoddis, 4) & "00"
        strnomdis = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddis, False)
        strdirec = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idCliente"), False)
        ImprimeXY Mid(strcadena, 45, Len(strcadena)), "T", 145, 54, 55, 0, 0, StrMsgError
                
        Printer.FontSize = 8
        ImprimeXY left(strtelef, 12) & "", "T", 153, 59, 95, 0, 0, StrMsgError
        
        '--- ETIQUETAS CLI
        ImprimeXY "Girado a:", "T", 10, 45, 40, 0, 0, StrMsgError
        ImprimeXY "Direccion:", "T", 10, 50, 40, 0, 0, StrMsgError
        ImprimeXY "RUC:", "T", 5, 59, 40, 0, 0, StrMsgError
        ImprimeXY "Tel./Fax:", "T", 10, 59, 83, 0, 0, StrMsgError
        
        '--- ETIQUETAS AVAL
        ImprimeXY "Aval Permanente:" & "", "T", 30, 67, 40, 0, 0, StrMsgError
        ImprimeXY "Domicilio:" & "", "T", 15, 72, 40, 0, 0, StrMsgError
        ImprimeXY "RUC:" & "", "T", 5, 78, 40, 0, 0, StrMsgError
        ImprimeXY "Tel./Fax:" & "", "T", 30, 78, 85, 0, 0, StrMsgError
        
        If rstletra.Fields("idAval") <> "" Then
            straval = traerCampo("personas", "glspersona", "idpersona", rstletra.Fields("idaval"), False)
            strdirecaval = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY straval & "", "T", 150, 67, 63, 0, 0, StrMsgError
            ImprimeXY left(strdirecaval, 50) & "", "T", 150, 72, 53, 0, 0, StrMsgError
            
            strTelefonoAval = traerCampo("personas", "Telefonos", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY left(strTelefonoAval, 12) & "", "T", 153, 78, 102, 0, 0, StrMsgError
            
            strRUCAval = traerCampo("personas", "RUC", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY strRUCAval, "T", 153, 78, 52, 0, 0, StrMsgError
        End If
        
        Printer.Print Chr$(149)
        Printer.Print ""
        Printer.Print ""
        Printer.EndDoc
    End If
    
End Sub

Private Sub IMPRIME_LETRA_APIMAS(strAno As String, strNumLetra As String, strSerieLetra As String)
Dim csql As String
Dim rstletra As New ADODB.Recordset
Dim strimporteletras As String
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer
Dim strdirec  As String
Dim strtelef As String
Dim StrMsgError As String
Dim straval As String
Dim strdirecaval As String
Dim strTelefonoAval As String
Dim strRUCAval As String
Dim strabrev As String
Dim strcoddis As String
Dim strnomdis As String
Dim strcoddisaval As String
Dim strnomdisaval As String
Dim impresoraletra  As String
Dim localidad As String
    
    impresoraletra = traerCampo("usuarios", "ImpresoraLetras", "idUsuario", glsUser, True)
    If impresoraletra <> "" Then
        For Each p In Printers
           If UCase(p.DeviceName) = impresoraletra Then
              Set Printer = p
              indPrinter = True
              Exit For
           End If
        Next p
    End If
    
    strNumLetra = Format(strNumLetra, "000000")
    intScale = 6
    'Printer.ScaleMode = intScale
    Printer.FontName = "Draft 17cpi"
    'Printer.FontSize = 8
    Printer.FontBold = False
    
    csql = "SELECT idletra,glsDocreferencia,FecEmision,FecVencimiento,TotalLetra," & _
           "CASE idMoneda WHEN 'PEN' THEN 'S/.' ELSE 'US$' END AS idMoneda,idCliente,glsCliente,RucCliente,idaval " & _
           "FROM LETRAS  " & _
           "WHERE IDEMPRESA = '" & glsEmpresa & "' " & _
           "AND IDLETRA = '" & strNumLetra & "' " & _
           "AND IDSERIE = '" & strSerieLetra & "' "
    rstletra.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstletra.EOF Then
        '--- NUMERO LETRA
        ImprimeXY rstletra.Fields("idLetra") & "", "T", 20, 25, 47, 0, 0, StrMsgError
        strabrev = traerCampo("documentos", "abredocumento", "iddocumento", left(rstletra.Fields("glsDocreferencia"), 2), False)
        '--- REFERENCIA
        ImprimeXY strabrev & " " & right(rstletra.Fields("glsDocreferencia"), 12) & "", "T", 20, 25, 71, 0, 0, StrMsgError
        '--- IMPRIME EL DIA EMISION
        ImprimeXY Format(Day(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 25, 104, 0, 0, StrMsgError
        '--- IMPRIME EL MES EMISION
        ImprimeXY Format(Month(rstletra.Fields("FecEmision")), "00") & "", "T", 2, 25, 110, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO EMISION
        ImprimeXY Format(Year(rstletra.Fields("FecEmision")), "0000") & "", "T", 4, 25, 118, 0, 0, StrMsgError
        '--- LUGAR DE GIRO
        localidad = traerCampo("personas", "glsPersona", "idPersona", glsSucursal, False)
        ImprimeXY localidad, "T", 20, 25, 130, 0, 0, StrMsgError
        '--- IMPRIME EL DIA VENCIMIENTO
        ImprimeXY Format(Day(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 25, 158, 0, 0, StrMsgError
        '--- IMPRIME EL MES VENCIMIENTO
        ImprimeXY Format(Month(rstletra.Fields("FecVencimiento")), "00") & "", "T", 2, 25, 166, 0, 0, StrMsgError
        '--- IMPRIME EL AÑO VENCIMIENTO
        ImprimeXY Format(Year(rstletra.Fields("FecVencimiento")), "0000") & "", "T", 4, 25, 172, 0, 0, StrMsgError
        '--- Signo de la modeda
        ImprimeXY rstletra.Fields("idMoneda") & "", "T", 3, 25, 182, 2, 0, StrMsgError
        '--- IMPORTE
        ImprimeXY rstletra.Fields("TotalLetra") & "", "N", 18, 25, 176, 2, 0, StrMsgError
        
        '--- IMPORTE EN LETRAS
        strimporteletras = "SON:" & Cadenanum(Format(rstletra.Fields("TotalLetra"), "0.00"), IIf(rstletra.Fields("idMoneda") = "S/.", "NUEVOS SOLES", "DOLARES AMERICANOS"))
        ImprimeXY Mid(strimporteletras, 5, Len(strimporteletras) - 3) & "", "T", 250, 38, 50, 0, 0, StrMsgError
        
        '--- CLIENTE
        ImprimeXY rstletra.Fields("glsCliente") & "", "T", 50, 49, 58, 0, 0, StrMsgError
        '--- DIRECCION
        strdirec = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 6
        ImprimeXY left(strdirec, 120) & "", "T", 145, 55, 58, 0, 0, StrMsgError
        '--- LOCALIDAD
        strcoddis = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idCliente"), False)
        strcoddis = left(strcoddis, 4) & "00"
        strnomdis = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddis, False)
        ImprimeXY strnomdis & "", "T", 150, 59, 95, 0, 0, StrMsgError
        '--- RUC
        ImprimeXY rstletra.Fields("RucCliente") & "", "T", 153, 64, 58, 0, 0, StrMsgError
        '--- TELEFONO
        strtelef = traerCampo("personas", "telefonos", "idpersona", rstletra.Fields("idCliente"), False)
        Printer.FontSize = 8
        ImprimeXY left(strtelef, 12) & "", "T", 153, 64, 98, 0, 0, StrMsgError
        
        If rstletra.Fields("idAval") <> "" Then
            straval = traerCampo("personas", "glspersona", "idpersona", rstletra.Fields("idaval"), False)
            strdirecaval = traerCampo("personas", "direccion", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY straval & "", "T", 150, 72, 50, 0, 0, StrMsgError
            ImprimeXY left(strdirecaval, 50) & "", "T", 150, 78, 50, 0, 0, StrMsgError
            
            strcoddisaval = traerCampo("personas", "iddistrito", "idpersona", rstletra.Fields("idaval"), False)
            strnomdisaval = traerCampo("ubigeo", "Glsubigeo", "iddistrito", strcoddisaval, False)
            ImprimeXY strnomdisaval & "", "T", 150, 85, 50, 0, 0, StrMsgError
            
            strTelefonoAval = traerCampo("personas", "Telefonos", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY left(strTelefonoAval, 12) & "", "T", 153, 84, 90, 0, 0, StrMsgError
            
            strRUCAval = traerCampo("personas", "RUC", "idpersona", rstletra.Fields("idaval"), False)
            ImprimeXY strRUCAval, "T", 153, 91, 39, 0, 0, StrMsgError
        End If
        
        Printer.Print Chr$(149)
        Printer.Print ""
        Printer.Print ""
        Printer.EndDoc
    End If
End Sub

Private Sub ImprimeXY(varData As Variant, strTipoDato As String, intTamanoCampo As Integer, intFila As Integer, intColu As Integer, intDecimales As Integer, intFilas As Integer, ByRef StrMsgError As String)
On Error GoTo Err
Dim i As Integer
Dim strDec  As String
Dim indFinWhile As Boolean
Dim intFilaImp As Integer
Dim intIndiceInicio As Integer
    
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
                    
            '--- Asigna la cantidad de decimales
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
                Printer.Print "(" & right((Space(intTamanoCampo) & Format(varData, "#,###,##0." & strDec)), intTamanoCampo - 2) & ")"
            End If
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub BloqueaColumnas(ByVal streval As Integer)

    If streval = 1 Then
        gLetras.Columns.ColumnByFieldName("idLetra").DisableEditor = True
        gLetras.Columns.ColumnByFieldName("Plazo").DisableEditor = True
        gLetras.Columns.ColumnByFieldName("FecVcto").DisableEditor = True
        gLetras.Columns.ColumnByFieldName("Porcentaje").DisableEditor = True
        gLetras.Columns.ColumnByFieldName("Porcentaje").DisableEditor = True
        gLetras.Columns.ColumnByFieldName("Capital").DisableEditor = True
        gLetras.Columns.ColumnByFieldName("Portes").DisableEditor = True
        gLetras.Columns.ColumnByFieldName("Interes").DisableEditor = True
        gLetras.Columns.ColumnByFieldName("IGV").DisableEditor = True
        gLetras.Columns.ColumnByFieldName("TOTAL").DisableEditor = True
    Else
        gLetras.Columns.ColumnByFieldName("idLetra").DisableEditor = False
        gLetras.Columns.ColumnByFieldName("Plazo").DisableEditor = False
        gLetras.Columns.ColumnByFieldName("FecVcto").DisableEditor = False
        gLetras.Columns.ColumnByFieldName("Porcentaje").DisableEditor = False
        gLetras.Columns.ColumnByFieldName("Porcentaje").DisableEditor = False
        gLetras.Columns.ColumnByFieldName("Capital").DisableEditor = False
        gLetras.Columns.ColumnByFieldName("Portes").DisableEditor = False
        gLetras.Columns.ColumnByFieldName("Interes").DisableEditor = False
        gLetras.Columns.ColumnByFieldName("IGV").DisableEditor = False
        gLetras.Columns.ColumnByFieldName("TOTAL").DisableEditor = False
    End If

End Sub
