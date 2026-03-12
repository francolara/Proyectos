VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmRptInventarioValorizado_Lotes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario Valorizado Por Tallas"
   ClientHeight    =   8460
   ClientLeft      =   4350
   ClientTop       =   3615
   ClientWidth     =   9285
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
   ScaleHeight     =   8460
   ScaleWidth      =   9285
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7890
      Width           =   1260
   End
   Begin VB.CommandButton btnSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7890
      Width           =   1260
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
      Height          =   7770
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.Frame fraReportes 
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
         ForeColor       =   &H00C00000&
         Height          =   765
         Index           =   7
         Left            =   120
         TabIndex        =   49
         Top             =   4545
         Width           =   8850
         Begin VB.CommandButton CmdAyudaCodigoRapido 
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
            Picture         =   "FrmRptInventarioValorizado_Lotes.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   295
            Width           =   390
         End
         Begin CATControls.CATTextBox TxtCodigoRapido 
            Height          =   315
            Left            =   3870
            TabIndex        =   51
            Tag             =   "TidMoneda"
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
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
            Container       =   "FrmRptInventarioValorizado_Lotes.frx":038A
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Código Rápido"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   2700
            TabIndex        =   52
            Top             =   330
            Width           =   1035
         End
      End
      Begin VB.Frame fraReportes 
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
         Height          =   765
         Index           =   6
         Left            =   135
         TabIndex        =   46
         Top             =   6120
         Width           =   8835
         Begin VB.OptionButton OptFamilia 
            Caption         =   "Con Familias"
            Height          =   285
            Index           =   0
            Left            =   1935
            TabIndex        =   48
            Top             =   270
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton OptFamilia 
            Caption         =   "Sin Familias"
            Height          =   285
            Index           =   1
            Left            =   5760
            TabIndex        =   47
            Top             =   270
            Width           =   1275
         End
      End
      Begin VB.Frame fraReportes 
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
         Height          =   765
         Index           =   2
         Left            =   135
         TabIndex        =   43
         Top             =   6885
         Width           =   8835
         Begin VB.OptionButton OptOrdena 
            Caption         =   "Ordenar por Descripción"
            Height          =   285
            Index           =   1
            Left            =   5760
            TabIndex        =   45
            Top             =   270
            Width           =   2220
         End
         Begin VB.OptionButton OptOrdena 
            Caption         =   "Ordenar por Código"
            Height          =   285
            Index           =   0
            Left            =   1935
            TabIndex        =   44
            Top             =   270
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame fraReportes 
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
         ForeColor       =   &H00C00000&
         Height          =   765
         Index           =   5
         Left            =   135
         TabIndex        =   38
         Top             =   1755
         Width           =   8805
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
            Left            =   8175
            Picture         =   "FrmRptInventarioValorizado_Lotes.frx":03A6
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   270
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Moneda 
            Height          =   315
            Left            =   1395
            TabIndex        =   40
            Tag             =   "TidMoneda"
            Top             =   300
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
            Container       =   "FrmRptInventarioValorizado_Lotes.frx":0730
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Moneda 
            Height          =   315
            Left            =   2370
            TabIndex        =   41
            Top             =   300
            Width           =   5760
            _ExtentX        =   10160
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
            Container       =   "FrmRptInventarioValorizado_Lotes.frx":074C
            Vacio           =   -1  'True
         End
         Begin VB.Label lbl_Moneda 
            Appearance      =   0  'Flat
            Caption         =   "Moneda"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   225
            TabIndex        =   42
            Top             =   330
            Width           =   765
         End
      End
      Begin VB.Frame fraReportes 
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
         Height          =   795
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   5325
         Width           =   8835
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   3600
            TabIndex        =   28
            Top             =   300
            Width           =   1320
            _ExtentX        =   2328
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
            Format          =   121962497
            CurrentDate     =   38667
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            Caption         =   "Al"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3225
            TabIndex        =   29
            Top             =   375
            Width           =   165
         End
      End
      Begin VB.Frame fraReportes 
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
         ForeColor       =   &H00C00000&
         Height          =   1995
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   8835
         Begin VB.Frame fraNivel 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Height          =   390
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   8625
            Begin VB.CommandButton cmbAyudaNivel 
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
               Index           =   0
               Left            =   8100
               Picture         =   "FrmRptInventarioValorizado_Lotes.frx":0768
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   45
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaNivel 
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
               Index           =   1
               Left            =   8100
               Picture         =   "FrmRptInventarioValorizado_Lotes.frx":0AF2
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   390
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaNivel 
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
               Index           =   2
               Left            =   8100
               Picture         =   "FrmRptInventarioValorizado_Lotes.frx":0E7C
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   750
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaNivel 
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
               Index           =   3
               Left            =   8100
               Picture         =   "FrmRptInventarioValorizado_Lotes.frx":1206
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   1110
               Width           =   390
            End
            Begin VB.CommandButton cmbAyudaNivel 
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
               Index           =   4
               Left            =   8100
               Picture         =   "FrmRptInventarioValorizado_Lotes.frx":1590
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   1470
               Width           =   390
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   0
               Left            =   1305
               TabIndex        =   12
               Tag             =   "TidNivelPred"
               Top             =   45
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":191A
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   0
               Left            =   2280
               TabIndex        =   13
               Top             =   45
               Width           =   5790
               _ExtentX        =   10213
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":1936
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   1
               Left            =   1305
               TabIndex        =   14
               Tag             =   "TidNivelPred"
               Top             =   390
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":1952
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   1
               Left            =   2280
               TabIndex        =   15
               Top             =   390
               Width           =   5790
               _ExtentX        =   10213
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":196E
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   2
               Left            =   1305
               TabIndex        =   16
               Tag             =   "TidNivelPred"
               Top             =   750
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":198A
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   2
               Left            =   2280
               TabIndex        =   17
               Top             =   750
               Width           =   5790
               _ExtentX        =   10213
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":19A6
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   3
               Left            =   1305
               TabIndex        =   18
               Tag             =   "TidNivelPred"
               Top             =   1110
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":19C2
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   3
               Left            =   2280
               TabIndex        =   19
               Top             =   1110
               Width           =   5790
               _ExtentX        =   10213
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":19DE
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   4
               Left            =   1305
               TabIndex        =   20
               Tag             =   "TidNivelPred"
               Top             =   1470
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":19FA
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   4
               Left            =   2280
               TabIndex        =   21
               Top             =   1470
               Width           =   5790
               _ExtentX        =   10213
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
               Container       =   "FrmRptInventarioValorizado_Lotes.frx":1A16
               Vacio           =   -1  'True
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   135
               TabIndex        =   26
               Top             =   45
               Width           =   345
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel:"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   135
               TabIndex        =   25
               Top             =   390
               Width           =   390
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel:"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   135
               TabIndex        =   24
               Top             =   750
               Width           =   390
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel:"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   135
               TabIndex        =   23
               Top             =   1110
               Width           =   390
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel:"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   135
               TabIndex        =   22
               Top             =   1470
               Width           =   390
            End
         End
      End
      Begin VB.Frame fraReportes 
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
         ForeColor       =   &H00C00000&
         Height          =   765
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   8835
         Begin VB.CommandButton cmbAyudaSucursal 
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
            Left            =   8160
            Picture         =   "FrmRptInventarioValorizado_Lotes.frx":1A32
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   315
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Sucursal 
            Height          =   315
            Left            =   1395
            TabIndex        =   31
            Tag             =   "TidMoneda"
            Top             =   330
            Width           =   960
            _ExtentX        =   1693
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
            Container       =   "FrmRptInventarioValorizado_Lotes.frx":1DBC
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Sucursal 
            Height          =   315
            Left            =   2415
            TabIndex        =   32
            Top             =   330
            Width           =   5730
            _ExtentX        =   10107
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
            Container       =   "FrmRptInventarioValorizado_Lotes.frx":1DD8
            Vacio           =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            Caption         =   "Sucursal"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   225
            TabIndex        =   33
            Top             =   375
            Width           =   765
         End
      End
      Begin VB.Frame fraReportes 
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
         ForeColor       =   &H00C00000&
         Height          =   765
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   8835
         Begin VB.CommandButton cmbAyudaAlmacen 
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
            Left            =   8160
            Picture         =   "FrmRptInventarioValorizado_Lotes.frx":1DF4
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   225
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Almacen 
            Height          =   315
            Left            =   1395
            TabIndex        =   35
            Tag             =   "TidAlmacen"
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
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
            Container       =   "FrmRptInventarioValorizado_Lotes.frx":217E
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Almacen 
            Height          =   315
            Left            =   2415
            TabIndex        =   36
            Top             =   240
            Width           =   5730
            _ExtentX        =   10107
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
            Container       =   "FrmRptInventarioValorizado_Lotes.frx":219A
            Vacio           =   -1  'True
         End
         Begin VB.Label lbl_Almacen 
            Appearance      =   0  'Flat
            Caption         =   "Almacén"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   225
            TabIndex        =   37
            Top             =   270
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "FrmRptInventarioValorizado_Lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnSalir_Click()
    
    Unload Me

End Sub

Private Sub cmbAyudaAlmacen_Click()
Dim strCondicion As String
    
    If txtCod_Sucursal.Text = "" Then
        MsgBox "Seleccione una sucursal", vbInformation, App.Title
        txtCod_Sucursal.SetFocus
        Exit Sub
    End If

    strCondicion = " AND idSucursal = '" & txtCod_Sucursal.Text & "'"
    mostrarAyuda "ALMACENVTA", TxtCod_Almacen, TxtGls_Almacen, strCondicion
    
End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"

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

Private Sub cmbAyudaSucursal_Click()

    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub CmdAyudaCodigoRapido_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim TxtGlsCodigoRapido                  As Object

    mostrarAyuda "PRODUCTOSCODIGORAPIDO", TxtCodigoRapido, TxtGlsCodigoRapido
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    
    Me.top = 0
    Me.left = 0
    
    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    TxtGls_Almacen.Text = "TODOS LOS ALMACENES"
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    dtpFecha.Value = Format(Date, "dd/mm/yyyy")
    txtCod_Moneda.Text = "PEN"
    
    If leeParametro("FILTRO_POR_CODIGORAPIDO") = "S" Then
        
        fraReportes(7).Visible = True
        Me.Height = 9030
        Frame1.Height = 7770
        fraReportes(3).top = 5325
        fraReportes(6).top = 6120
        fraReportes(2).top = 6885
        btnAceptar.top = 7890
        btnSalir.top = 7890
        
    Else
        
        fraReportes(7).Visible = False
        Me.Height = 8295
        Frame1.Height = 7070
        fraReportes(3).top = 4560
        fraReportes(6).top = 5355
        fraReportes(2).top = 6120
        btnAceptar.top = 7170
        btnSalir.top = 7170
        
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarNiveles(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsj      As New ADODB.Recordset

    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    fraNivel.Height = 355 * glsNumNiveles
    i = 0
    Do While Not rsj.EOF
        lblNivel(i).Caption = "" & rsj.Fields("GlsTipoNivel")
        txtGls_Nivel(i).Text = "TODO(A) LOS(AS) " & rsj.Fields("GlsTipoNivel")
        rsj.MoveNext
        i = i + 1
    Loop
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    
    Exit Sub

Err:
    If rsj.State = 1 Then rsj.Close: Set rsj = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Almacen_Change()

    If TxtCod_Almacen.Text <> "" Then
        TxtGls_Almacen.Text = traerCampo("Almacenes", "GlsAlmacen", "idAlmacen", TxtCod_Almacen.Text, True)
    Else
        TxtGls_Almacen.Text = "TODOS LOS ALMACENES"
    End If

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If

End Sub

Private Sub BtnAceptar_Click()
On Error GoTo Err
Dim StrMsgError             As String
Dim GlsReporte              As String
Dim cWhereNiveles           As String
Dim fCorte                  As String
Dim X                       As Integer
Dim COrdena                 As String
Dim CFamilia                As String

    GlsReporte = "rptInventarioValorizadoLotes.rpt"
    fCorte = Format(dtpFecha.Value, "yyyy-mm-dd")
    If Len(Trim(txtCod_Nivel(0).Text)) > 0 Then
        cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles, "00") & " = ''" & txtCod_Nivel(0).Text & "'' "
        If Len(Trim(txtCod_Nivel(1).Text)) > 0 Then
            cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles - 1, "00") & "   = ''" & txtCod_Nivel(1).Text & "'' "
            If Len(Trim(txtCod_Nivel(2).Text)) > 0 Then
                cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles - 2, "00") & "  = ''" & txtCod_Nivel(2).Text & "'' "
            End If
        End If
    End If
    
    For X = 1 To glsNumNiveles
        cNiveles = cNiveles & "vn.idNivel" & Format(X, "00") & ", vn.GlsNivel" & Format(X, "00") & ","
    Next X
    
    If OptOrdena(0).Value Then
    
        COrdena = "C"
    
    Else
        
        COrdena = "D"
        
    End If
    
    CFamilia = ""
    
    If OptFamilia(1).Value Then
        
        CFamilia = "_Listado"
    
    End If
    

    mostrarReporte left(GlsReporte, Len(GlsReporte) - 4) & Format(glsNumNiveles, "00") & CFamilia & ".rpt", "varEmpresa|varSucursal|varAlmacen|varMoneda|varFecha|varNiveles|varGlsNiveles|varOrdena|varCodigoRapido|varNivel01|varNivel02", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & TxtCod_Almacen.Text & "|" & Trim(txtCod_Moneda.Text) & "|" & fCorte & "|" & cWhereNiveles & "|" & cNiveles & "|" & COrdena & "|" & TxtCodigoRapido.Text & "|" & txtCod_Nivel(0).Text & "|" & txtCod_Nivel(1).Text, GlsForm, StrMsgError
    
    If StrMsgError <> "" Then GoTo Err

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
