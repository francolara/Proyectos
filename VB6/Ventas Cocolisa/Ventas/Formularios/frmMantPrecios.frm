VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantPrecios 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de  Precios"
   ClientHeight    =   5130
   ClientLeft      =   2415
   ClientTop       =   3330
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   9045
      Begin VB.Frame fraNivel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   60
         TabIndex        =   18
         Top             =   600
         Width           =   8925
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   0
            Left            =   8400
            Picture         =   "frmMantPrecios.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   90
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   1
            Left            =   8400
            Picture         =   "frmMantPrecios.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   360
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   2
            Left            =   8400
            Picture         =   "frmMantPrecios.frx":0714
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   720
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   3
            Left            =   8400
            Picture         =   "frmMantPrecios.frx":0A9E
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1080
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaNivel 
            Height          =   315
            Index           =   4
            Left            =   8400
            Picture         =   "frmMantPrecios.frx":0E28
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1440
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   0
            Left            =   1305
            TabIndex        =   24
            Tag             =   "TidNivelPred"
            Top             =   75
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
            Container       =   "frmMantPrecios.frx":11B2
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   25
            Top             =   75
            Width           =   6090
            _ExtentX        =   10742
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
            Container       =   "frmMantPrecios.frx":11CE
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   1
            Left            =   1305
            TabIndex        =   26
            Tag             =   "TidNivelPred"
            Top             =   390
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
            Container       =   "frmMantPrecios.frx":11EA
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   27
            Top             =   390
            Width           =   6090
            _ExtentX        =   10742
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
            Container       =   "frmMantPrecios.frx":1206
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   2
            Left            =   1305
            TabIndex        =   28
            Tag             =   "TidNivelPred"
            Top             =   750
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
            Container       =   "frmMantPrecios.frx":1222
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   2
            Left            =   2280
            TabIndex        =   29
            Top             =   750
            Width           =   6090
            _ExtentX        =   10742
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
            Container       =   "frmMantPrecios.frx":123E
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   3
            Left            =   1305
            TabIndex        =   30
            Tag             =   "TidNivelPred"
            Top             =   1110
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
            Container       =   "frmMantPrecios.frx":125A
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   3
            Left            =   2280
            TabIndex        =   31
            Top             =   1110
            Width           =   6090
            _ExtentX        =   10742
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
            Container       =   "frmMantPrecios.frx":1276
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Nivel 
            Height          =   315
            Index           =   4
            Left            =   1305
            TabIndex        =   32
            Tag             =   "TidNivelPred"
            Top             =   1470
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
            Container       =   "frmMantPrecios.frx":1292
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Nivel 
            Height          =   315
            Index           =   4
            Left            =   2280
            TabIndex        =   33
            Top             =   1470
            Width           =   6090
            _ExtentX        =   10742
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
            Container       =   "frmMantPrecios.frx":12AE
            Vacio           =   -1  'True
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   38
            Top             =   90
            Width           =   405
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   37
            Top             =   405
            Width           =   405
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   36
            Top             =   765
            Width           =   405
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   35
            Top             =   1125
            Width           =   405
         End
         Begin VB.Label lblNivel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   34
            Top             =   1485
            Width           =   405
         End
      End
      Begin VB.Frame fraContenido 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2910
         Left            =   30
         TabIndex        =   1
         Top             =   570
         Width           =   8790
         Begin VB.CheckBox chkfactor 
            Caption         =   "Habilitar Factor"
            Height          =   285
            Left            =   90
            TabIndex        =   54
            Top             =   1620
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.Frame frafactor 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   780
            Left            =   90
            TabIndex        =   47
            Top             =   3315
            Visible         =   0   'False
            Width           =   8295
            Begin CATControls.CATTextBox txtCosto 
               Height          =   315
               Left            =   870
               TabIndex        =   48
               Tag             =   "NFactor"
               Top             =   330
               Width           =   1110
               _ExtentX        =   1958
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
               Alignment       =   1
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantPrecios.frx":12CA
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtFactor_Costo 
               Height          =   315
               Left            =   3825
               TabIndex        =   49
               Tag             =   "NFactor"
               Top             =   315
               Width           =   1020
               _ExtentX        =   1799
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
               Alignment       =   1
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantPrecios.frx":12E6
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtFactor2_Costo 
               Height          =   315
               Left            =   6660
               TabIndex        =   50
               Tag             =   "NFactor"
               Top             =   270
               Width           =   1020
               _ExtentX        =   1799
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
               Alignment       =   1
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   -2147483640
               Container       =   "frmMantPrecios.frx":1302
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               EnterTab        =   -1  'True
            End
            Begin VB.Label Label7 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Costo : "
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   270
               TabIndex        =   53
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label10 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Factor 1 :"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   3015
               TabIndex        =   52
               Top             =   330
               Width           =   675
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Factor 2 :"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   5895
               TabIndex        =   51
               Top             =   315
               Width           =   675
            End
         End
         Begin VB.CheckBox chkAfecto 
            Appearance      =   0  'Flat
            Caption         =   "Afecto al IGV"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   7020
            TabIndex        =   46
            Tag             =   "NafectoIGV"
            Top             =   1200
            Width           =   1440
         End
         Begin VB.CommandButton cmbAyudaUMVenta 
            Height          =   315
            Left            =   8400
            Picture         =   "frmMantPrecios.frx":131E
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   840
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaProducto 
            Height          =   315
            Left            =   8400
            Picture         =   "frmMantPrecios.frx":16A8
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Producto 
            Height          =   285
            Left            =   1290
            TabIndex        =   4
            Tag             =   "TidMarca"
            Top             =   135
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
            Container       =   "frmMantPrecios.frx":1A32
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Producto 
            Height          =   285
            Left            =   2250
            TabIndex        =   5
            Top             =   135
            Width           =   6090
            _ExtentX        =   10742
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
            Container       =   "frmMantPrecios.frx":1A4E
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_UMVenta 
            Height          =   315
            Left            =   1290
            TabIndex        =   6
            Tag             =   "TidUMVenta"
            Top             =   870
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
            Container       =   "frmMantPrecios.frx":1A6A
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_UMVenta 
            Height          =   315
            Left            =   2250
            TabIndex        =   7
            Top             =   870
            Width           =   6090
            _ExtentX        =   10742
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
            Container       =   "frmMantPrecios.frx":1A86
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Factor 
            Height          =   315
            Left            =   1290
            TabIndex        =   8
            Tag             =   "NFactor"
            Top             =   1245
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantPrecios.frx":1AA2
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtVal_VV 
            Height          =   315
            Left            =   1050
            TabIndex        =   9
            Tag             =   "NFactor"
            Top             =   2205
            Width           =   1290
            _ExtentX        =   2275
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantPrecios.frx":1ABE
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtVal_IGV 
            Height          =   315
            Left            =   3300
            TabIndex        =   10
            Tag             =   "NFactor"
            Top             =   2205
            Width           =   1290
            _ExtentX        =   2275
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantPrecios.frx":1ADA
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtVal_PV 
            Height          =   315
            Left            =   5505
            TabIndex        =   11
            Tag             =   "NFactor"
            Top             =   2205
            Width           =   1290
            _ExtentX        =   2275
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantPrecios.frx":1AF6
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_UMCompra 
            Height          =   315
            Left            =   1290
            TabIndex        =   43
            Tag             =   "TidTipoProducto"
            Top             =   510
            Width           =   915
            _ExtentX        =   1614
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
            MaxLength       =   8
            Container       =   "frmMantPrecios.frx":1B12
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_UMCompra 
            Height          =   315
            Left            =   2250
            TabIndex        =   44
            Top             =   510
            Width           =   6090
            _ExtentX        =   10742
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
            Container       =   "frmMantPrecios.frx":1B2E
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtDctoListaPrec 
            Height          =   315
            Left            =   7890
            TabIndex        =   55
            Top             =   2220
            Width           =   465
            _ExtentX        =   820
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
            Container       =   "frmMantPrecios.frx":1B4A
            Text            =   "0"
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Max % Dcto:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6945
            TabIndex        =   56
            Top             =   2250
            Width           =   900
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "U.M. Compra:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   45
            Top             =   525
            Width           =   975
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "U.M. Venta:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   17
            Top             =   915
            Width           =   855
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Producto:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Factor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   15
            Top             =   1290
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "V.V. Unit.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   2280
            Width           =   720
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "I.G.V. Unit.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2400
            TabIndex        =   13
            Top             =   2250
            Width           =   825
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "P.V. Unit.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4695
            TabIndex        =   12
            Top             =   2295
            Width           =   720
         End
      End
      Begin CATControls.CATTextBox txtCod_Lista 
         Height          =   315
         Left            =   1380
         TabIndex        =   39
         Tag             =   "TidTipoProducto"
         Top             =   240
         Width           =   915
         _ExtentX        =   1614
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
         MaxLength       =   8
         Container       =   "frmMantPrecios.frx":1B66
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Lista 
         Height          =   315
         Left            =   2340
         TabIndex        =   40
         Top             =   240
         Width           =   6090
         _ExtentX        =   10742
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
         Container       =   "frmMantPrecios.frx":1B82
         Vacio           =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Lista de Precios:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   285
         Width           =   1170
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   3960
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
            Picture         =   "frmMantPrecios.frx":1B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":1F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":238A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":2724
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":2ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":2E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":31F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":358C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":3926
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":3CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":405A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPrecios.frx":4D1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1164
      ButtonWidth     =   2381
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "        Nuevo        "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "        Grabar        "
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
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMantPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NumNiveles As Integer
Dim indMovNivel As Boolean
Dim indCalculando As Boolean
Dim intTipoOpe As Integer


Private Sub chkfactor_Click()
If indCalculando Then Exit Sub
If chkfactor.Value = 1 Then
 frafactor.Enabled = True
 txtCosto.SetFocus
Else
 frafactor.Enabled = False
 txtVal_VV.SetFocus
End If
indCalculando = False
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

Private Sub cmbAyudaProducto_Click()
          
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto, " AND idNivel = '" & txtCod_Nivel(NumNiveles - 1).Text & "'"
    
End Sub

Private Sub cmbAyudaUMVenta_Click()

    mostrarAyuda "PRESENTACIONES", txtCod_UMVenta, txtGls_UMVenta, " AND idProducto = '" & txtCod_Producto.Text & "'"

End Sub

Private Sub Form_Load()

txtVal_VV.Decimales = glsDecimalesPrecios
txtVal_IGV.Decimales = glsDecimalesPrecios
txtVal_PV.Decimales = glsDecimalesPrecios
txtCosto.Decimales = glsDecimalesPrecios
txtFactor_Costo.Decimales = glsDecimalesPrecios
txtFactor2_Costo.Decimales = glsDecimalesPrecios

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
On Error GoTo Err
Select Case Button.Index
    Case 1 'Nuevo
        nuevo
        fraGeneral.Enabled = True
    Case 2 'Grabar
        Grabar StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        If intTipoOpe = 1 Then
            Unload Me
            Exit Sub
        End If
        
        nuevo
        
        habilitaBotones 3
    Case 3 'Modificar
        fraGeneral.Enabled = True
'    Case 4, 6 'Cancelar
'        fraGeneral.Enabled = False
'        Unload Me
    Case 5 'Imprimir
    
'    Case 6 'Lista
'        fraListado.Visible = True
'        fraCabecera.Visible = False
    Case 4, 6, 7 'Salir
        Unload Me
        Exit Sub
End Select

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Lista_Change()
    txtGls_Lista.Text = traerCampo("listaprecios", "GlsLista", "idLista", txtCod_Lista.Text, True)
End Sub

Private Sub txtCod_Nivel_Change(Index As Integer)
Dim StrMsgError As String
Dim i As Integer

On Error GoTo Err

If indMovNivel Then Exit Sub

txtGls_Nivel(Index).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_Nivel(Index).Text, True)

indMovNivel = True
For i = Index + 1 To txtCod_Nivel.Count - 1
    txtCod_Nivel(i).Text = ""
    txtGls_Nivel(i).Text = ""
Next
indMovNivel = False

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Nivel_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim peso As Integer
    Dim strCodTipoNivel As String
    Dim strCondPred As String
    Dim strCodJerarquia As String
    
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
Dim RsP As New ADODB.Recordset
Dim StrMsgError As String

On Error GoTo Err

    csql = "SELECT GlsProducto,idUMCompra,afectoIGV,idTipoProducto " & _
           "FROM productos " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idProducto = '" & txtCod_Producto.Text & "'"
    
    RsP.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If Not RsP.EOF Then
        txtGls_Producto.Text = "" & RsP.Fields("GlsProducto")
        txtCod_UMCompra.Text = "" & RsP.Fields("idUMCompra")
        chkAfecto.Value = Val("" & RsP.Fields("afectoIGV"))
        
        If ("" & RsP.Fields("idTipoProducto")) = "06002" Then
            txtCod_UMVenta.Enabled = False
            txtCod_UMVenta.BackColor = &HFFF9F0
            txtCod_UMVenta.Vacio = True
        Else
            txtCod_UMVenta.Enabled = True
            txtCod_UMVenta.BackColor = &HFFFFFF
            txtCod_UMVenta.Vacio = False
        End If
    Else
        txtGls_Producto.Text = ""
        txtCod_UMCompra.Text = ""
        chkAfecto.Value = 0
    End If

If RsP.State = 1 Then RsP.Close
Set RsP = Nothing
Exit Sub
Err:
If RsP.State = 1 Then RsP.Close
Set RsP = Nothing
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_UMCompra_Change()
    txtGls_UMCompra.Text = traerCampo("unidadmedida", "abreUM", "idUM", txtCod_UMCompra.Text, False)
End Sub

Private Sub txtCod_UMVenta_Change()
    txtGls_UMVenta.Text = traerCampo("unidadmedida", "abreUM", "idUM", txtCod_UMVenta.Text, False)
    
    If txtGls_UMVenta.Text <> "" Then
        txt_Factor.Text = traerCampo("presentaciones", "Factor", "idUM", txtCod_UMVenta.Text, True, " idProducto = '" & txtCod_Producto.Text & "'")
    End If
    
End Sub

Public Sub MostrarForm(ByVal strCodLista As String, ByVal strProd As String, ByVal strUM As String, ByVal dblFactor As Double, ByRef dblVV As Double, ByRef dblIGV As Double, ByRef dblPV As Double, ByRef dblCosto As Double, ByRef dblFactorUnit As Double, ByRef dblFactor2Unit As Double, ByRef StrMsgError As String)
Dim i  As Integer
On Error GoTo Err

intTipoOpe = 1 'Modificar

mostrarNiveles StrMsgError
If StrMsgError <> "" Then GoTo Err

txtCod_Lista.Text = strCodLista
txtCod_Producto.Text = strProd
txtCod_UMVenta.Text = strUM
txt_Factor.Text = dblFactor

indMovNivel = True
For i = 1 To NumNiveles
    If i > 1 Then
        txtCod_Nivel((NumNiveles - i)).Text = traerCampo("niveles", "idNivelPred", "idNivel", txtCod_Nivel((NumNiveles - i) + 1).Text, True)
    Else
        txtCod_Nivel((NumNiveles - i)).Text = traerCampo("productos", "idNivel", "idProducto", txtCod_Producto.Text, True)
    End If
    
    txtGls_Nivel((NumNiveles - i)).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_Nivel(NumNiveles - i).Text, True)
Next
indMovNivel = False


indCalculando = True

txtVal_VV.Text = dblVV
txtVal_IGV.Text = dblIGV
txtVal_PV.Text = dblPV
If dblCosto <> 0 Or dblFactorUnit > 0 Or dblFactor2Unit > 0 Then
    chkfactor.Value = 1
    frafactor.Enabled = True
  
    txtCosto.Text = dblCosto
    txtFactor_Costo.Text = dblFactorUnit
    txtFactor2_Costo.Text = dblFactor2Unit
Else
    chkfactor.Value = 0
    frafactor.Enabled = False
    
End If
indCalculando = False


TxtDctoListaPrec.Text = traerCampo("PRECIOSVENTA", "MaxDcto", "idLista", strCodLista, True, "idProducto = '" & strProd & "' and idUM = '" & strUM & "' ")

habilitaBotones 3

frmMantPrecios.Show 1

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub


Private Sub mostrarNiveles(ByRef StrMsgError As String)
Dim rsj As New ADODB.Recordset
Dim i As Integer

On Error GoTo Err

'Limpiando Tag
For i = 0 To 4
    txtCod_Nivel(i).Tag = ""
Next

'jalamos tipos nivel
rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
 
NumNiveles = Val("" & rsj.RecordCount)

fraNivel.Height = 355 * NumNiveles

i = 0

Do While Not rsj.EOF
        
    lblNivel(i).Caption = "" & rsj.Fields("GlsTipoNivel")
    
    rsj.MoveNext
    
    i = i + 1

Loop

fraContenido.top = fraNivel.top + fraNivel.Height - 70

fraGeneral.Height = fraNivel.top + fraNivel.Height + fraContenido.Height + 200

Me.Height = Toolbar1.Height + fraGeneral.Height + 500


If rsj.State = 1 Then rsj.Close
Set rsj = Nothing
Exit Sub
Err:
If rsj.State = 1 Then rsj.Close
Set rsj = Nothing
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()
Dim C As Object

For Each C In Me.Controls

    If TypeOf C Is CATTextBox And C.Name <> "txtCod_Lista" And C.Name <> "txtGls_Lista" And C.Name <> "txtCod_Nivel" And C.Name <> "txtGls_Nivel" Then
        C.Text = ""
    End If
    
Next
End Sub

Private Sub habilitaBotones(indexBoton As Integer)
Dim indHabilitar As Boolean

Select Case indexBoton
    Case 1, 2, 3 'Nuevo, Grabar, Modificar
        If indexBoton = 2 Then indHabilitar = True
        Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
        Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
        Toolbar1.Buttons(3).Visible = False 'Modificar
        Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
        Toolbar1.Buttons(5).Visible = False 'Imprimir
        Toolbar1.Buttons(6).Visible = False 'Lista
    Case 4, 6 'Cancelar, Lista
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(3).Visible = False
        Toolbar1.Buttons(4).Visible = False
        Toolbar1.Buttons(5).Visible = False
        Toolbar1.Buttons(6).Visible = False
End Select

End Sub

Private Sub Grabar(ByRef StrMsgError As String)
Dim indIniTrans As Boolean
Dim strMsg As String
On Error GoTo Err

If Trim(txtCod_Producto.Text) = "" Then
    StrMsgError = "Faltan datos"
    txtCod_Producto.OnError = True
    GoTo Err
End If

If Trim(txtCod_UMVenta.Text) = "" And txtCod_UMVenta.Vacio = False Then
    StrMsgError = "Faltan datos"
    txtCod_UMVenta.OnError = True
    GoTo Err
End If

If intTipoOpe = 0 Then 'Grabo

    'Validamos si ya existe el resgistro
    If traerCampo("preciosventa", "idLista", "idLista", txtCod_Lista.Text, True, " idUM = '" & txtCod_UMVenta.Text & "' AND idProducto = '" & txtCod_Producto.Text & "'") <> "" Then
        StrMsgError = "El registro ya existe"
        GoTo Err
    End If
    
    csql = "INSERT INTO preciosventa (idEmpresa,idLista,idProducto,idUM,VVUnit,IGVUnit,PVUnit,CostoUnit,FactorUnit,Factor2Unit,MaxDcto) VALUES(" & _
           "'" & glsEmpresa & "','" & txtCod_Lista.Text & "','" & txtCod_Producto.Text & "','" & txtCod_UMVenta.Text & "'," & _
           txtVal_VV.Value & "," & txtVal_IGV.Value & "," & txtVal_PV.Value & "," & txtCosto.Value & "," & txtFactor_Costo.Value & "," & txtFactor2_Costo.Value & "," & TxtDctoListaPrec.Text & ")"
           
    strMsg = "Grabó"
Else 'Modifico


    csql = "UPDATE preciosventa SET MaxDcto = " & TxtDctoListaPrec.Text & ",VVUnit = " & txtVal_VV.Value & ",IGVUnit = " & txtVal_IGV.Value & ", PVUnit = " & txtVal_PV.Value & ",CostoUnit=" & txtCosto.Value & ",FactorUnit=" & txtFactor_Costo.Value & ",Factor2Unit=" & txtFactor2_Costo.Value & "  " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idLista = '" & txtCod_Lista.Text & "' AND idProducto = '" & txtCod_Producto.Text & "' AND idUM = '" & txtCod_UMVenta.Text & "'"
           
    strMsg = "Modificó"
    
End If

Cn.BeginTrans
indIniTrans = True

Cn.Execute csql

If Val(txtVal_VV.Text) > 0 Then
    csql = "Update Productos Set  indInsertaPrecioLista = '1' Where idProducto =  '" & txtCod_Producto.Text & "' And idEmpresa = '" & glsEmpresa & "' "
    Cn.Execute (csql)
End If

Cn.CommitTrans

MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title

Exit Sub
Err:
If indIniTrans Then Cn.RollbackTrans
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCosto_Change()
If indCalculando Then Exit Sub
 txtVal_VV.Text = Format(Val(Format(txtCosto.Value, "0.00")) * Val(Format(txtFactor_Costo.Text, "0.00")) * Val(Format(txtFactor2_Costo.Text, "0.00")), "0.00")
 indCalculando = False
End Sub

Private Sub txtFactor_Costo_Change()
If indCalculando Then Exit Sub
 txtVal_VV.Text = Format(Val(Format(txtCosto.Value, "0.00")) * Val(Format(txtFactor_Costo.Text, "0.00")) * Val(Format(txtFactor2_Costo.Text, "0.00")), "0.00")
 indCalculando = False
End Sub

Private Sub txtFactor_Costo_KeyPress(KeyAscii As Integer)
'txtVal_VV.Text = Format(Val(Format(txtCosto.Value, "0.00")) * Val(Format(txtFactor_Costo.Text, "0.00")), "0.00")

End Sub

Private Sub txtFactor2_Costo_Change()
If indCalculando Then Exit Sub
 txtVal_VV.Text = Format(Val(Format(txtCosto.Value, "0.00")) * Val(Format(txtFactor_Costo.Text, "0.00")) * Val(Format(txtFactor2_Costo.Text, "0.00")), "0.00")
 indCalculando = False
End Sub

Private Sub txtVal_PV_Change()
    If indCalculando Then Exit Sub
    
    indCalculando = True
    If chkAfecto.Value Then
        txtVal_VV.Text = Val(txtVal_PV.Value) / (glsIGV + 1)
        txtVal_IGV.Text = Val(txtVal_PV.Value) - Val(txtVal_VV.Value)
    Else
        txtVal_IGV.Text = 0#
        txtVal_VV.Text = txtVal_PV.Value
    End If
    indCalculando = False
End Sub

Private Sub txtVal_VV_Change()
    If indCalculando Then Exit Sub
    
    indCalculando = True
    If chkAfecto.Value Then
        txtVal_IGV.Text = Val(txtVal_VV.Value) * glsIGV
        txtVal_PV.Text = Val(txtVal_VV.Value) + Val(txtVal_IGV.Value)
    Else
        txtVal_IGV.Text = 0#
        txtVal_PV.Text = txtVal_VV.Value
    End If
    indCalculando = False
End Sub

Public Sub mostrarFormNuevo(ByVal strCodLista As String, ByVal strCodNivel As String, ByRef StrMsgError As String)
Dim i  As Integer
On Error GoTo Err

intTipoOpe = 0 'Nuevo

mostrarNiveles StrMsgError
If StrMsgError <> "" Then GoTo Err

txtCod_Lista.Text = strCodLista

txtCod_Nivel((NumNiveles - 1)).Text = strCodNivel

indMovNivel = True

For i = 2 To NumNiveles
    txtCod_Nivel((NumNiveles - i)).Text = traerCampo("niveles", "idNivelPred", "idNivel", txtCod_Nivel((NumNiveles - i) + 1).Text, True)
    
    txtGls_Nivel((NumNiveles - i)).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_Nivel(NumNiveles - i).Text, True)
Next
indMovNivel = False

txtVal_VV.Decimales = glsDecimalesPrecios
txtVal_IGV.Decimales = glsDecimalesPrecios
txtVal_PV.Decimales = glsDecimalesPrecios

habilitaBotones 3

frmMantPrecios.Show 1

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
