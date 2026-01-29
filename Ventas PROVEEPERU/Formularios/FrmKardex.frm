VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmKardex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kardex"
   ClientHeight    =   8010
   ClientLeft      =   7065
   ClientTop       =   3750
   ClientWidth     =   7260
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
   ScaleHeight     =   8010
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkDetallado 
      Caption         =   "Detallado"
      Height          =   285
      Left            =   5580
      TabIndex        =   53
      Top             =   7470
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7380
      Width           =   1185
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3735
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7380
      Width           =   1185
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
      Height          =   7170
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   7170
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
         Left            =   180
         TabIndex        =   49
         Top             =   4500
         Width           =   6825
         Begin VB.CheckBox ChkAgrupadoCR 
            Caption         =   "Agrupado por C.R."
            Height          =   285
            Left            =   5085
            TabIndex        =   54
            Top             =   315
            Width           =   1680
         End
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
            Left            =   4530
            Picture         =   "FrmKardex.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   295
            Width           =   390
         End
         Begin CATControls.CATTextBox TxtCodigoRapido 
            Height          =   315
            Left            =   3240
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
            Container       =   "FrmKardex.frx":038A
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Código Rápido"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   2070
            TabIndex        =   52
            Top             =   330
            Width           =   1035
         End
      End
      Begin VB.Frame fraReportes 
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
         ForeColor       =   &H00C00000&
         Height          =   1935
         Index           =   2
         Left            =   180
         TabIndex        =   27
         Top             =   1665
         Width           =   6820
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
            TabIndex        =   28
            Top             =   120
            Width           =   6645
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
               Left            =   6255
               Picture         =   "FrmKardex.frx":03A6
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   1470
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
               Left            =   6255
               Picture         =   "FrmKardex.frx":0730
               Style           =   1  'Graphical
               TabIndex        =   32
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
               Index           =   2
               Left            =   6255
               Picture         =   "FrmKardex.frx":0ABA
               Style           =   1  'Graphical
               TabIndex        =   31
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
               Index           =   1
               Left            =   6255
               Picture         =   "FrmKardex.frx":0E44
               Style           =   1  'Graphical
               TabIndex        =   30
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
               Index           =   0
               Left            =   6255
               Picture         =   "FrmKardex.frx":11CE
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   45
               Width           =   390
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   0
               Left            =   945
               TabIndex        =   34
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
               Container       =   "FrmKardex.frx":1558
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   0
               Left            =   1920
               TabIndex        =   35
               Top             =   45
               Width           =   4305
               _ExtentX        =   7594
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
               Container       =   "FrmKardex.frx":1574
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   1
               Left            =   945
               TabIndex        =   36
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
               Container       =   "FrmKardex.frx":1590
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   1
               Left            =   1920
               TabIndex        =   37
               Top             =   390
               Width           =   4305
               _ExtentX        =   7594
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
               Container       =   "FrmKardex.frx":15AC
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   2
               Left            =   945
               TabIndex        =   38
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
               Container       =   "FrmKardex.frx":15C8
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   2
               Left            =   1920
               TabIndex        =   39
               Top             =   750
               Width           =   4305
               _ExtentX        =   7594
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
               Container       =   "FrmKardex.frx":15E4
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   3
               Left            =   945
               TabIndex        =   40
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
               Container       =   "FrmKardex.frx":1600
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   3
               Left            =   1920
               TabIndex        =   41
               Top             =   1110
               Width           =   4305
               _ExtentX        =   7594
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
               Container       =   "FrmKardex.frx":161C
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   4
               Left            =   945
               TabIndex        =   42
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
               Container       =   "FrmKardex.frx":1638
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   4
               Left            =   1920
               TabIndex        =   43
               Top             =   1500
               Width           =   4305
               _ExtentX        =   7594
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
               Container       =   "FrmKardex.frx":1654
               Vacio           =   -1  'True
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel:"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   45
               TabIndex        =   48
               Top             =   1470
               Width           =   390
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel:"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   45
               TabIndex        =   47
               Top             =   1110
               Width           =   390
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel:"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   45
               TabIndex        =   46
               Top             =   750
               Width           =   390
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel:"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   45
               TabIndex        =   45
               Top             =   390
               Width           =   390
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   45
               TabIndex        =   44
               Top             =   45
               Width           =   345
            End
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Formato"
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   6210
         Width           =   6825
         Begin VB.OptionButton OptResumido 
            Caption         =   "Resumido"
            Height          =   210
            Left            =   4680
            TabIndex        =   26
            Top             =   315
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.OptionButton Opt2 
            Caption         =   "Por Familia"
            Height          =   210
            Left            =   3720
            TabIndex        =   25
            Top             =   315
            Width           =   1680
         End
         Begin VB.OptionButton Opt1 
            Caption         =   "General"
            Height          =   210
            Left            =   1800
            TabIndex        =   24
            Top             =   315
            Width           =   1680
         End
      End
      Begin VB.CheckBox ChkValorizado 
         Caption         =   "Valorizado"
         Height          =   240
         Left            =   5895
         TabIndex        =   22
         Top             =   6435
         Value           =   1  'Checked
         Width           =   1095
      End
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
         Left            =   6600
         Picture         =   "FrmKardex.frx":1670
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4095
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaProducto 
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
         Left            =   6600
         Picture         =   "FrmKardex.frx":19FA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3735
         Width           =   390
      End
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
         Left            =   6600
         Picture         =   "FrmKardex.frx":1D84
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   370
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   5355
         Width           =   6825
         Begin VB.ComboBox cbodatos 
            Height          =   330
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   270
            Width           =   3390
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   855
         Width           =   6825
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   1
            Top             =   300
            Width           =   1250
            _ExtentX        =   2196
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
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
            TabIndex        =   2
            Top             =   300
            Width           =   1250
            _ExtentX        =   2196
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
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
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
            Height          =   210
            Left            =   855
            TabIndex        =   10
            Top             =   375
            Width           =   465
         End
      End
      Begin CATControls.CATTextBox txtCod_Almacen 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Tag             =   "TidAlmacen"
         Top             =   375
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
         Container       =   "FrmKardex.frx":210E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Almacen 
         Height          =   315
         Left            =   2205
         TabIndex        =   14
         Top             =   375
         Width           =   4365
         _ExtentX        =   7699
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
         Container       =   "FrmKardex.frx":212A
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   1155
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   3735
         Width           =   1020
         _ExtentX        =   1799
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
         Container       =   "FrmKardex.frx":2146
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   2205
         TabIndex        =   16
         Top             =   3735
         Width           =   4365
         _ExtentX        =   7699
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
         Container       =   "FrmKardex.frx":2162
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1155
         TabIndex        =   4
         Tag             =   "TidMoneda"
         Top             =   4095
         Width           =   1020
         _ExtentX        =   1799
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
         Container       =   "FrmKardex.frx":217E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2205
         TabIndex        =   20
         Top             =   4095
         Width           =   4365
         _ExtentX        =   7699
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
         Container       =   "FrmKardex.frx":219A
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   225
         TabIndex        =   21
         Top             =   4125
         Width           =   765
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Producto"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   225
         TabIndex        =   17
         Top             =   3810
         Width           =   765
      End
      Begin VB.Label lbl_Almacen 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   225
         TabIndex        =   15
         Top             =   450
         Width           =   630
      End
   End
End
Attribute VB_Name = "FrmKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CCodProducto                    As String

Private Sub cmbAyudaAlmacen_Click()
Dim strCondicion As String
    
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

Private Sub cmbAyudaProducto_Click()
    
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim Fini        As String, Ffin As String, strIni As String, strFin As String
Dim rsReporte   As New ADODB.Recordset
Dim reporte     As CRAXDRT.Report
Dim StrMsgError As String
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim GlsReporte      As String
Dim cWhereNiveles   As String
Dim CIdProducto     As String

    Screen.MousePointer = 11
    Fini = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strIni = Format(dtpfInicio.Value, "dd/mm/yyyy")
    strFin = Format(dtpFFinal.Value, "dd/mm/yyyy")
    
    If OptResumido.Value = True Then
    
        If traerCampo("Parametros", "ValParametro", "GlsParametro", "FORMATO_KARDEX_RESUMIDO", True) = 1 Then
            mostrarReporte "rptKardexValorizadoResumxConcepto.rpt", "parEmpresa|parSucursal|parAlmacen|parMoneda|parFechaIni|parFechaFin|parProducto", glsEmpresa & "|" & glsSucursal & "|" & Trim(TxtCod_Almacen.Text) & "|" & txtCod_Moneda.Text & "|" & Fini & "|" & Ffin & "|" & txtCod_Producto.Text, "Kardex Valorizado Resumido", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Else
            mostrarReporte "rptKardexValorizadoResum" & IIf(ChkDetallado.Value = 1, "Detallado", "") & ".rpt", "parEmpresa|parAlmacen|parMoneda|parFechaIni|parFechaFin|parProducto", glsEmpresa & "|" & Trim(TxtCod_Almacen.Text) & "|" & txtCod_Moneda.Text & "|" & Fini & "|" & Ffin & "|" & txtCod_Producto.Text, "Kardex Valorizado Resumido" & IIf(ChkDetallado.Value = 1, " - Detallado", ""), StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
                
    Else
        If CCodProducto = "CodigoRapido" Then
            If Len(Trim("" & txtCod_Producto.Text)) > 0 Then
                CIdProducto = traerCampo("Productos", "IdProducto", "CodigoRapido", txtCod_Producto.Text, True)
            Else
                CIdProducto = txtCod_Producto.Text
            End If
        Else
            CIdProducto = txtCod_Producto.Text
        End If
        
        Set rsReporte = kardex(StrMsgError, Trim(CIdProducto))
        If StrMsgError <> "" Then GoTo Err
         
        If Opt1.Value = True Then
            If ChkValorizado.Value = 1 Then
                If leeParametro("FORMATOKARDEX") = "2" Then
                    
                    Set reporte = aplicacion.OpenReport(gStrRutaRpts & "rptKardex_Formato2.rpt")
                
                Else
                    
                    If leeParametro("ORDENA_KARDEX_DESCRIPCION") = "S" Then
                        Set reporte = aplicacion.OpenReport(gStrRutaRpts & "RptKardexPorDescripcion.rpt")
                    Else
                        If ChkAgrupadoCR.Value = 1 Then
                            Set reporte = aplicacion.OpenReport(gStrRutaRpts & "RptKardexPorCodigoRapido.rpt")
                        Else
                            Set reporte = aplicacion.OpenReport(gStrRutaRpts & "rptKardex.rpt")
                        End If
                    End If
                    
                End If
            Else
                Set reporte = aplicacion.OpenReport(gStrRutaRpts & "rptKardexProducto.rpt")
            End If
        ElseIf Opt2.Value = True Then
            Set reporte = aplicacion.OpenReport(gStrRutaRpts & "rptKardex" & Format(glsNumNiveles, "00") & ".rpt")
        End If
        
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
    End If

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

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError                 As String

    Me.top = 0
    Me.left = 0
    
    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If leeParametro("VIZUALIZA_CODIGO_RAPIDO") = "S" Then
        txtCod_Producto.MaxLength = 20
        CCodProducto = "CodigoRapido"
    Else
        txtCod_Producto.MaxLength = 8
        CCodProducto = "IdProducto"
    End If
    
    txtCod_Moneda.Text = "PEN"
    txtGls_Moneda.Text = "NUEVOS SOLES"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    Opt1.Value = True
    
    Tipo_datos
    
    If leeParametro("FILTRO_POR_CODIGORAPIDO") = "S" Then
        
        fraReportes(7).Visible = True
        Me.Height = 8480
        Frame1.Height = 7170
        fraReportes(3).top = 5355
        fraReportes(0).top = 6210
        cmdaceptar.top = 7380
        cmdsalir.top = 7380
        ChkDetallado.top = 7470
        
    Else
        
        fraReportes(7).Visible = False
        Me.Height = 7415
        Frame1.Height = 6225
        fraReportes(3).top = 4455
        fraReportes(0).top = 5310
        cmdaceptar.top = 6390
        cmdsalir.top = 6390
        ChkDetallado.top = 6480
        
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Tipo_datos()
On Error GoTo Err
Dim SQL2 As String
Dim rs1 As New ADODB.Recordset

    SQL2 = "SELECT iddato,glsdato From Datos where idtipodatoS = '20' order by iddato desc "
    If rs1.State = 1 Then rs1.Close
    rs1.Open SQL2, Cn, adOpenStatic, adLockOptimistic
    If Not rs1.EOF Then
        Do While Not rs1.EOF
            cbodatos.AddItem rs1.Fields("glsdato") & Space(100) & rs1.Fields("iddato")
            rs1.MoveNext
        Loop
    End If
    rs1.Close: Set rs1 = Nothing
    If cbodatos.ListCount > 0 Then cbodatos.ListIndex = 0
            
    Exit Sub
        
Err:
    MsgBox "Se ha producido el sgt. error: " & Err.Description, vbInformation, App.Title
    Exit Sub
End Sub

Private Sub Opt1_Click()
    
    If Opt1.Value Then
        
        fraReportes(2).Enabled = False
        ChkDetallado.Visible = False
        
    End If
    
End Sub

Private Sub Opt2_Click()
    
    If Opt2.Value Then
        
        fraReportes(2).Enabled = True
        ChkDetallado.Visible = False
        
    End If
    
End Sub

Private Sub OptResumido_Click()
    
    If OptResumido.Value Then
        
        fraReportes(2).Enabled = False
        ChkDetallado.Visible = True
    
    Else
        
        fraReportes(2).Enabled = True
        ChkDetallado.Visible = False
        
    End If
    
End Sub

Private Sub txtCod_Almacen_Change()
    
    If TxtCod_Almacen.Text <> "" Then
        TxtGls_Almacen.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", TxtCod_Almacen.Text, True, "idSucursal =  '" & glsSucursal & "' ")
    Else
        TxtGls_Almacen.Text = "TODOS LOS ALMACENES"
    End If

End Sub

Private Sub txtCod_Almacen_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        TxtCod_Almacen.Text = ""
    End If

End Sub

Private Sub txtCod_Almacen_KeyPress(KeyAscii As Integer)
Dim strCondicion As String

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "ALMACENVTA", TxtCod_Almacen, TxtGls_Almacen, strCondicion
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

Private Sub txtCod_Producto_Change()
    
    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", CCodProducto, txtCod_Producto.Text, True)
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

Private Function kardex(ByRef StrMsgError As String, Optional codproducto As String) As Recordset
On Error GoTo Err
Dim rstemp                      As New ADODB.Recordset
Dim RsD                         As New ADODB.Recordset
Dim rsx                         As New ADODB.Recordset
Dim strSQL                      As String
Dim i                           As Integer
Dim dblSaldo                    As Double
Dim dblValSaldo                 As Double
Dim CodProd                     As String
Dim CodProdAnt                  As String
Dim strFecIni                   As String
Dim strFecFin                   As String
Dim cadenaniveles               As String
Dim strSimboloMoneda            As String
Dim cNiveles                    As String
Dim X                           As Integer
Dim visCodRapido                As String
Dim dblValSaldoIni              As Double
Dim strProAnt                   As String
Dim rstempAnt                   As New ADODB.Recordset
Dim indVerifica                 As Boolean
Dim StrCodalmacenx              As String
Dim Codalmax                    As String
Dim DblCostoPAnt                As Double
Dim CArray(8)                   As String
Dim RsTempAntClone              As ADODB.Recordset
Dim cWhereNiveles               As String
Dim cWhereNivelesSp             As String
Dim CTipoNivel                  As String
Dim IndPasa                     As Boolean

    If Len(Trim(txtCod_Nivel(0).Text)) > 0 Then
        cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles, "00") & " = '" & txtCod_Nivel(0).Text & "' "
        cWhereNivelesSp = cWhereNivelesSp & "And vn.idNivel" & Format(glsNumNiveles, "00") & " = ''" & txtCod_Nivel(0).Text & "'' "
        If Len(Trim(txtCod_Nivel(1).Text)) > 0 Then
            cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles - 1, "00") & " = '" & txtCod_Nivel(1).Text & "' "
            cWhereNivelesSp = cWhereNivelesSp & "And vn.idNivel" & Format(glsNumNiveles - 1, "00") & " = ''" & txtCod_Nivel(1).Text & "'' "
            If Len(Trim(txtCod_Nivel(2).Text)) > 0 Then
                cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles - 2, "00") & " = '" & txtCod_Nivel(2).Text & "' "
                cWhereNivelesSp = cWhereNivelesSp & "And vn.idNivel" & Format(glsNumNiveles - 2, "00") & " = ''" & txtCod_Nivel(2).Text & "'' "
            End If
        End If
    End If
    
    StrCodalmacenx = ""
    indVerifica = False
    DblCostoPAnt = 0
    
    visCodRapido = Trim("" & traerCampo("parametros", "ValParametro", "GlsParametro", "VIZUALIZA_CODIGO_RAPIDO", True))
    'If ChkAgrupadoCR.Value = 1 Then
    '    visCodRapido = "S"
    'End If
    
    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    cadenaniveles = right(cbodatos.Text, 5)
    If cadenaniveles = "20004" Then
        cadenaniveles = " "
        CTipoNivel = ""
    Else
        cadenaniveles = " and n.tipo =  '" & cadenaniveles & "' "
        CTipoNivel = right(cbodatos.Text, 5)
    End If
    cNiveles = ""
    strSimboloMoneda = txtGls_Moneda.Text
    
    If Opt1.Value = True Then
    
        strSQL = "SELECT vc.idValesCab,vc.idConcepto, " & _
                    "vc.tipoVale," & _
                    "vc.fechaEmision,iif('" & visCodRapido & "' = 'S',pr.Codigorapido,vd.IdProducto) xIdProducto, " & _
                    "pr.IdProducto," & _
                    "pr.GlsProducto," & _
                    "vc.IdAlmacen," & _
                    "vc.idProvCliente,pe.ruc,pe.GlsPersona," & _
                    "vc.GlsDocReferencia," & _
                    "um.abreUM," & _
                    "vd.Cantidad, " & _
                    "CASE '" & txtCod_Moneda.Text & "' WHEN 'PEN' THEN  IiF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  isnull(vc.TipoCambio,t.tcventa))" & _
                                            "WHEN 'USD' THEN  IiF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  isnull(vc.TipoCambio,t.tcventa)) " & _
                    "END as VVUnit,(Vc.IdAlmacen + ' ' + A.GlsAlmacen) As Almacen,Z.GlsConcepto,pr.IdTallaPeso,Pr.CodigoRapido "
                       
        strSQL = strSQL & "FROM valescab vc " & _
                    "inner join valesdet vd " & _
                        "on vc.idValesCab = vd.idValesCab AND vc.idEmpresa = vd.idEmpresa AND vc.idSucursal = vd.idSucursal " & _
                    "AND vc.tipoVale = vd.tipoVale " & _
                    "Left Join Conceptos Z On Vc.IdConcepto = Z.IdConcepto " & _
                    "left join personas pe " & _
                        "on vc.idProvCliente = pe.IdPersona " & _
                    "Inner Join Almacenes A " & _
                        "On Vc.IdEmpresa = A.IdEmpresa And Vc.IdAlmacen = A.IdAlmacen And Vc.idSucursal =  A.idSucursal " & _
                    "inner join productos pr " & _
                        "on vd.idProducto = pr.idProducto AND vd.idEmpresa = pr.idEmpresa " & _
                    "inner join unidadmedida um " & _
                        "on pr.idUMCompra = um.idUM " & _
                    "inner join niveles n " & _
                        "on pr.idnivel = n.idnivel and pr.idempresa = n.idempresa  "
        strSQL = strSQL & _
                    "Left Join tiposdecambio t " & _
                        "On vc.fechaEmision = t.fecha "
        strSQL = strSQL & "WHERE vc.idEmpresa = '" & glsEmpresa & "' AND pr.idEmpresa = '" & glsEmpresa & "' " & _
                    IIf(Len(Trim(TxtCod_Almacen.Text)) = 0, "", "AND vc.idAlmacen = '" & TxtCod_Almacen.Text & "' ") & _
                    "AND (vc.idPeriodoInv) IN " & _
                    "(" & _
                        "SELECT pi.idPeriodoInv " & _
                        "FROM periodosinv pi " & _
                        "WHERE pi.idEmpresa = vc.idEmpresa AND pi.idSucursal = vc.idSucursal and CAST(pi.FecInicio AS DATE) <= CAST('" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' AS DATE) " & _
                        "and (CAST(pi.FecFin AS DATE) >= CAST('" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' AS DATE) or pi.FecFin is null)" & _
                    ") " & _
                    "AND vc.estValeCab <> 'ANU'  And pr.estProducto = 'A' AND CAST(vc.FechaEmision AS DATE) BETWEEN CAST('" & strFecIni & "' AS DATE) AND CAST('" & strFecFin & "' AS DATE)  " & _
                    "And (pr.CodigoRapido = '" & TxtCodigoRapido.Text & "' Or '' = '" & TxtCodigoRapido.Text & "') " & _
                    cadenaniveles
                    
    Else
    
        For X = 1 To glsNumNiveles
            cNiveles = cNiveles & "vn.idNivel" & Format(X, "00") & ", vn.GlsNivel" & Format(X, "00") & ","
        Next X
    
        strSQL = "SELECT " & cNiveles & " vc.idValesCab," & _
                    "vc.tipoVale," & _
                    "vc.fechaEmision,iif('" & visCodRapido & "' = 'S',pr.Codigorapido,vd.IdProducto) xIdProducto, " & _
                    "pr.IdProducto," & _
                    "pr.GlsProducto," & _
                    "vc.IdAlmacen," & _
                    "vc.idProvCliente,pe.ruc,pe.GlsPersona," & _
                    "vc.GlsDocReferencia," & _
                    "um.abreUM," & _
                    "vd.Cantidad, " & _
                    "CASE '" & txtCod_Moneda.Text & "' WHEN 'PEN' THEN  IiF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  isnull(vc.TipoCambio,t.tcventa) )" & _
                                            "WHEN 'USD' THEN  IiF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  isnull(vc.TipoCambio,t.tcventa) ) " & _
                    "END as VVUnit,(Vc.IdAlmacen + ' ' + A.GlsAlmacen) As Almacen,Z.GlsConcepto,pr.IdTallaPeso,Pr.CodigoRapido "
                       
        strSQL = strSQL & "FROM valescab vc " & _
                    "inner join valesdet vd " & _
                        "on vc.idValesCab = vd.idValesCab AND vc.idEmpresa = vd.idEmpresa AND vc.idSucursal = vd.idSucursal " & _
                    "AND vc.tipoVale = vd.tipoVale " & _
                    "Left Join Conceptos Z On Vc.IdConcepto = Z.IdConcepto " & _
                    "left join personas pe " & _
                        "on vc.idProvCliente = pe.IdPersona " & _
                    "Inner Join Almacenes A " & _
                        "On Vc.IdEmpresa = A.IdEmpresa And Vc.IdAlmacen = A.IdAlmacen And Vc.idSucursal =  A.idSucursal " & _
                    "inner join productos pr " & _
                        "on vd.idProducto = pr.idProducto AND vd.idEmpresa = pr.idEmpresa " & _
                    "inner join unidadmedida um " & _
                        "on pr.idUMCompra = um.idUM " & _
                    "inner join niveles n " & _
                        "on pr.idnivel = n.idnivel and pr.idempresa = n.idempresa  " & _
                    "inner join vw_niveles vn on pr.idnivel = vn.idNivel01 and pr.idempresa = vn.idempresa "
        strSQL = strSQL & _
                    "Left Join tiposdecambio t " & _
                        "On vc.fechaEmision = t.fecha "
        strSQL = strSQL & "WHERE vc.idEmpresa = '" & glsEmpresa & "' AND pr.idEmpresa = '" & glsEmpresa & "' " & _
                    IIf(Len(Trim(TxtCod_Almacen.Text)) = 0, "", "AND vc.idAlmacen = '" & TxtCod_Almacen.Text & "' ") & _
                    "AND (vc.idPeriodoInv) IN " & _
                    "(" & _
                        "SELECT pi.idPeriodoInv " & _
                        "FROM periodosinv pi " & _
                        "WHERE pi.idEmpresa = vc.idEmpresa AND pi.idSucursal = vc.idSucursal and CAST(pi.FecInicio AS DATE) <= CAST('" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' AS DATE) " & _
                        "and (CAST(pi.FecFin AS DATE) >= CAST('" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' AS DATE) or pi.FecFin is null)" & _
                    ") " & _
                    "AND vc.estValeCab <> 'ANU'  AND pr.estProducto = 'A' AND CAST(vc.FechaEmision AS DATE) BETWEEN CAST('" & strFecIni & "' AS DATE) AND CAST('" & strFecFin & "' AS DATE)  " & _
                    "And (pr.CodigoRapido = '" & TxtCodigoRapido.Text & "' Or '' = '" & TxtCodigoRapido.Text & "') " & _
                    cWhereNiveles & cadenaniveles
                    
    End If
    
    If codproducto <> "" Then
        strSQL = strSQL & "  AND vd.idProducto = '" & codproducto & "'"
    End If
    
    If leeParametro("ORDENA_KARDEX_DESCRIPCION") = "S" Then
        strSQL = strSQL & " ORDER BY vc.IdAlmacen, vd.GlsProducto,vc.FechaEmision,vc.tipovale,vc.idvalescab"
    Else
        If visCodRapido = "S" Or ChkAgrupadoCR.Value = 1 Then
            strSQL = strSQL & " ORDER BY vc.IdAlmacen, pr.Codigorapido,Pr.IdProducto,vc.FechaEmision,vc.tipovale,vc.idvalescab"
        Else
            strSQL = strSQL & " ORDER BY vc.IdAlmacen, vd.IdProducto,vc.FechaEmision,vc.tipovale,vc.idvalescab"
        End If
    End If
            
    If rstemp.State = 1 Then rstemp.Close
    rstemp.Open strSQL, Cn, adOpenKeyset, adLockOptimistic
    
    Set rstemp.ActiveConnection = Nothing
    
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idValesCab", adVarChar, 15, adFldIsNullable
    RsD.Fields.Append "fechaEmision", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "IdProducto", adVarChar, 50, adFldIsNullable
    RsD.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    RsD.Fields.Append "IdAlmacen", adVarChar, 8, adFldIsNullable
    RsD.Fields.Append "idProvCliente", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "ruc", adVarChar, 20, adFldIsNullable
    RsD.Fields.Append "GlsPersona", adVarChar, 120, adFldIsNullable
    RsD.Fields.Append "GlsDocReferencia", adVarChar, 300, adFldIsNullable
    RsD.Fields.Append "abreUM", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "Ingreso", adDouble, , adFldIsNullable
    RsD.Fields.Append "Salida", adDouble, , adFldIsNullable
    RsD.Fields.Append "Saldo", adDouble, , adFldIsNullable
    RsD.Fields.Append "ValorUnit", adDouble, , adFldIsNullable
    RsD.Fields.Append "ValorTotal", adDouble, , adFldIsNullable
    RsD.Fields.Append "ValorSaldo", adDouble, , adFldIsNullable
    RsD.Fields.Append "FechaInicio", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "FechaFin", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "SimboloMoneda", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "Empresa", adVarChar, 255, adFldIsNullable
    RsD.Fields.Append "ruc_empresa", adVarChar, 11, adFldIsNullable
    RsD.Fields.Append "Sistema", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "Almacen", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsConcepto", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "idNivel01", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsNivel01", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "idNivel02", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsNivel02", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "idNivel03", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsNivel03", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "idNivel04", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsNivel04", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "CostoPromedio", adDouble, , adFldIsNullable
    RsD.Fields.Append "PesoIngreso", adDouble, , adFldIsNullable
    RsD.Fields.Append "PesoSalida", adDouble, , adFldIsNullable
    RsD.Fields.Append "PesoSaldo", adDouble, , adFldIsNullable
    RsD.Fields.Append "CodigoRapido", adVarChar, 50, adFldIsNullable
    RsD.Open
    
    CodProd = ""
    Codalmax = ""
    i = 0
    dblSaldo = 0
    dblValSaldo = 0
    
    'Set rstempAnt = DataProcedimiento("Spu_ListaSaldoInicialProductos", strMsgError, glsEmpresa, txtCod_Almacen.Text, codproducto, TxtCod_Moneda.Text, Format(dtpfInicio.Value, "yyyy-mm-dd"), Format(dtpFFinal.Value, "yyyy-mm-dd"))
    'If strMsgError <> "" Then GoTo ERR
    
    rstempAnt.Open "Exec Spu_ListaSaldoInicialProductos '" & glsEmpresa & "','" & TxtCod_Almacen.Text & "','" & codproducto & "','" & txtCod_Moneda.Text & "','" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "','" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "','" & cWhereNivelesSp & "','" & CTipoNivel & "',0,'" & TxtCodigoRapido.Text & "'", Cn, adOpenStatic, adLockReadOnly
    Set RsTempAntClone = rstempAnt.Clone(adLockReadOnly)
    
    Do While Not rstempAnt.EOF
        
        RsD.AddNew
                
        i = 0
        RsD.Fields("Item") = i
        RsD.Fields("idValesCab") = "I"
        RsD.Fields("fechaEmision") = ""
        RsD.Fields("IdProducto") = "" & rstempAnt.Fields("xIdProducto")
        RsD.Fields("GlsProducto") = rstempAnt.Fields("GlsProducto")
        RsD.Fields("IdAlmacen") = "" & rstempAnt.Fields("IdAlmacen")
        RsD.Fields("idProvCliente") = ""
        RsD.Fields("ruc") = ""
        RsD.Fields("GlsPersona") = ""
        RsD.Fields("GlsDocReferencia") = "SALDO INICIAL"
        RsD.Fields("abreUM") = "" & rstempAnt.Fields("abreUM")
        RsD.Fields("Ingreso") = "" & rstempAnt.Fields("Stock")
        RsD.Fields("Salida") = 0
        RsD.Fields("Saldo") = "" & rstempAnt.Fields("Stock")
        
        RsD.Fields("PesoIngreso") = Val("" & rstempAnt.Fields("Stock")) * Val("" & traerCampo("Productos", "IdTallaPeso", "IdProducto", rstempAnt.Fields("xIdProducto"), True))
        RsD.Fields("PesoSalida") = 0
        RsD.Fields("PesoSaldo") = Val("" & rstempAnt.Fields("Stock")) * Val("" & traerCampo("Productos", "IdTallaPeso", "IdProducto", rstempAnt.Fields("xIdProducto"), True))
        
        RsD.Fields("ValorUnit") = rstempAnt.Fields("VVUnit")
        RsD.Fields("ValorTotal") = rstempAnt.Fields("SaldoInicial") 'rstempAnt.Fields("Stock") * rstempAnt.Fields("SaldoInicial")
        RsD.Fields("ValorSaldo") = rstempAnt.Fields("SaldoInicial")
        RsD.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
        RsD.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
        RsD.Fields("SimboloMoneda") = strSimboloMoneda
        RsD.Fields("Empresa") = GlsNom_Empresa
        RsD.Fields("ruc_empresa") = Glsruc
        RsD.Fields("Sistema") = ""
        RsD.Fields("Almacen") = "" & rstempAnt.Fields("Almacen")
        RsD.Fields("GlsConcepto") = ""
        RsD.Fields("CodigoRapido") = "" & rstempAnt.Fields("CodigoRapido")
        
        dblSaldo = Val(dblSaldo) + Val("" & rstempAnt.Fields("Stock"))
        dblValSaldo = Val(dblValSaldo) + Val("" & rstempAnt.Fields("SaldoInicial"))
    
        If Opt2.Value = True Then
            If glsNumNiveles = 1 Then
                
                traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01", "A.IdProducto", Trim("" & rstempAnt.Fields("xIdProducto")), 2, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                
                If Len(Trim(CArray(0))) > 0 Then
                
                    RsD.Fields("idNivel01") = CArray(0)
                    RsD.Fields("GlsNivel01") = CArray(1)
                
                End If
                
            ElseIf glsNumNiveles = 2 Then
            
                traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02", "A.IdProducto", Trim("" & rstempAnt.Fields("xIdProducto")), 4, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                
                If Len(Trim(CArray(0))) > 0 Then
                
                    RsD.Fields("idNivel01") = CArray(2)
                    RsD.Fields("GlsNivel01") = CArray(3)
                    RsD.Fields("idNivel02") = CArray(0)
                    RsD.Fields("GlsNivel02") = CArray(1)
                
                End If
                
            ElseIf glsNumNiveles = 3 Then
                
                traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02,B.IdNivel03,B.GlsNivel03", "A.IdProducto", Trim("" & rstempAnt.Fields("xIdProducto")), 6, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                
                If Len(Trim(CArray(0))) > 0 Then
                
                    RsD.Fields("idNivel01") = CArray(4)
                    RsD.Fields("GlsNivel01") = CArray(5)
                    RsD.Fields("idNivel02") = CArray(2)
                    RsD.Fields("GlsNivel02") = CArray(3)
                    RsD.Fields("idNivel03") = CArray(0)
                    RsD.Fields("GlsNivel03") = CArray(1)
                
                End If
                
            ElseIf glsNumNiveles = 4 Then
                
                traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02,B.IdNivel03,B.GlsNivel03,B.IdNivel04,B.GlsNivel04", "A.IdProducto", Trim("" & rstempAnt.Fields("xIdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                
                If Len(Trim(CArray(0))) > 0 Then
                
                    RsD.Fields("idNivel01") = CArray(6)
                    RsD.Fields("GlsNivel01") = CArray(7)
                    RsD.Fields("idNivel02") = CArray(4)
                    RsD.Fields("GlsNivel02") = CArray(5)
                    RsD.Fields("idNivel03") = CArray(2)
                    RsD.Fields("GlsNivel03") = CArray(3)
                    RsD.Fields("idNivel04") = CArray(0)
                    RsD.Fields("GlsNivel04") = CArray(1)
                
                End If
                
            End If
        End If
        
        RsD.Fields("ValorUnit") = rstempAnt.Fields("VVUnit")
        RsD.Fields("ValorTotal") = rstempAnt.Fields("SaldoInicial") 'rstempAnt.Fields("Stock") * rstempAnt.Fields("SaldoInicial")
        RsD.Fields("ValorSaldo") = rstempAnt.Fields("SaldoInicial")
                
        If Val("" & rstempAnt.Fields("Stock")) > 0 Then
        
            RsD.Fields("CostoPromedio") = Val("" & rstempAnt.Fields("SaldoInicial")) / Val("" & rstempAnt.Fields("Stock"))
            'DblCostoPAnt = Val("" & RsTempAntClone.Fields("SaldoInicial")) / Val("" & RsTempAntClone.Fields("Stock"))
            
        Else
        
            RsD.Fields("CostoPromedio") = DblCostoPAnt
        
        End If

        rstempAnt.MoveNext

    Loop
   
    Do While Not rstemp.EOF
        If StrCodalmacenx = "" Then
            StrCodalmacenx = Trim("" & rstemp.Fields("IdAlmacen"))
        End If
        
        If StrCodalmacenx <> Trim("" & rstemp.Fields("IdAlmacen")) Then
            StrCodalmacenx = Trim("" & rstemp.Fields("IdAlmacen"))
            dblSaldo = 0
            dblValSaldo = 0
        End If
        
        If visCodRapido = "S" Then
            
            If CodProd <> Trim("" & rstemp.Fields("XIdProducto")) Then
                dblSaldo = 0
                dblValSaldo = 0
            End If
            
        Else
        
            If CodProd <> Trim("" & rstemp.Fields("IdProducto")) Then
                dblSaldo = 0
                dblValSaldo = 0
            End If
        
        End If
        
        IndPasa = False
        
        If visCodRapido = "S" Then
            If (CodProd <> rstemp.Fields("XIdProducto") And Codalmax <> Trim("" & rstemp.Fields("IdAlmacen"))) Or (CodProd = rstemp.Fields("XIdProducto") And Codalmax <> Trim("" & rstemp.Fields("IdAlmacen"))) Or (CodProd <> rstemp.Fields("XIdProducto") And Codalmax = Trim("" & rstemp.Fields("IdAlmacen"))) Then
                IndPasa = True
            End If
        Else
            If (CodProd <> rstemp.Fields("IdProducto") And Codalmax <> Trim("" & rstemp.Fields("IdAlmacen"))) Or (CodProd = rstemp.Fields("IdProducto") And Codalmax <> Trim("" & rstemp.Fields("IdAlmacen"))) Or (CodProd <> rstemp.Fields("IdProducto") And Codalmax = Trim("" & rstemp.Fields("IdAlmacen"))) Then
                IndPasa = True
            End If
        End If
        
        If IndPasa Then
                        
            If visCodRapido = "S" Then
                RsTempAntClone.Filter = "IdAlmacen = '" & Trim("" & rstemp.Fields("IdAlmacen")) & "' And XIdProducto = '" & Trim("" & rstemp.Fields("XIdProducto")) & "'"
            Else
                RsTempAntClone.Filter = "IdAlmacen = '" & Trim("" & rstemp.Fields("IdAlmacen")) & "' And IdProducto = '" & Trim("" & rstemp.Fields("IdProducto")) & "'"
            End If
            
            If Not RsTempAntClone.EOF Then
            
                dblSaldo = Val(dblSaldo) + Val("" & RsTempAntClone.Fields("Stock"))
                dblValSaldo = Val(dblValSaldo) + Val("" & RsTempAntClone.Fields("SaldoInicial"))
            
                If Val("" & RsTempAntClone.Fields("Stock")) > 0 Then
                
                    DblCostoPAnt = Val("" & RsTempAntClone.Fields("SaldoInicial")) / Val("" & RsTempAntClone.Fields("Stock"))
                    
                Else
                
                    DblCostoPAnt = 0
                
                End If
        
            Else
            
                RsD.AddNew
                i = 0
                RsD.Fields("Item") = i
                RsD.Fields("idValesCab") = "I"
                RsD.Fields("fechaEmision") = ""
                RsD.Fields("IdProducto") = "" & rstemp.Fields("XIdProducto")
                RsD.Fields("GlsProducto") = rstemp.Fields("GlsProducto")
                RsD.Fields("IdAlmacen") = "" & rstemp.Fields("IdAlmacen")
                RsD.Fields("idProvCliente") = ""
                RsD.Fields("ruc") = ""
                RsD.Fields("GlsPersona") = ""
                RsD.Fields("GlsDocReferencia") = "SALDO INICIAL"
                RsD.Fields("abreUM") = "" & rstemp.Fields("abreUM")
                RsD.Fields("Ingreso") = 0
                RsD.Fields("Salida") = 0
                RsD.Fields("Saldo") = 0
                RsD.Fields("PesoIngreso") = 0
                RsD.Fields("PesoSalida") = 0
                RsD.Fields("PesoSaldo") = 0
                RsD.Fields("ValorUnit") = 0
                RsD.Fields("ValorTotal") = 0
                RsD.Fields("ValorSaldo") = 0
                RsD.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
                RsD.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
                RsD.Fields("SimboloMoneda") = strSimboloMoneda
                RsD.Fields("Empresa") = GlsNom_Empresa
                RsD.Fields("ruc_empresa") = Glsruc
                RsD.Fields("Sistema") = ""
                RsD.Fields("Almacen") = "" & rstemp.Fields("Almacen")
                RsD.Fields("GlsConcepto") = ""
                RsD.Fields("CodigoRapido") = "" & rstemp.Fields("CodigoRapido")
                
                If Opt2.Value = True Then
                    If glsNumNiveles = 1 Then
                        
                        traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01", "A.IdProducto", Trim("" & rstemp.Fields("IdProducto")), 2, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                        
                        If Len(Trim(CArray(0))) > 0 Then
                        
                            RsD.Fields("idNivel01") = CArray(0)
                            RsD.Fields("GlsNivel01") = CArray(1)
                        
                        End If
                        
                    ElseIf glsNumNiveles = 2 Then
                    
                        traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02", "A.IdProducto", Trim("" & rstemp.Fields("IdProducto")), 4, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                        
                        If Len(Trim(CArray(0))) > 0 Then
                        
                            RsD.Fields("idNivel01") = CArray(2)
                            RsD.Fields("GlsNivel01") = CArray(3)
                            RsD.Fields("idNivel02") = CArray(0)
                            RsD.Fields("GlsNivel02") = CArray(1)
                        
                        End If
                        
                    ElseIf glsNumNiveles = 3 Then
                        
                        traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02,B.IdNivel03,B.GlsNivel03", "A.IdProducto", Trim("" & rstemp.Fields("IdProducto")), 6, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                        
                        If Len(Trim(CArray(0))) > 0 Then
                        
                            RsD.Fields("idNivel01") = CArray(4)
                            RsD.Fields("GlsNivel01") = CArray(5)
                            RsD.Fields("idNivel02") = CArray(2)
                            RsD.Fields("GlsNivel02") = CArray(3)
                            RsD.Fields("idNivel03") = CArray(0)
                            RsD.Fields("GlsNivel03") = CArray(1)
                        
                        End If
                        
                    ElseIf glsNumNiveles = 4 Then
                        
                        traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02,B.IdNivel03,B.GlsNivel03,B.IdNivel04,B.GlsNivel04", "A.IdProducto", Trim("" & rstemp.Fields("IdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                        
                        If Len(Trim(CArray(0))) > 0 Then
                        
                            RsD.Fields("idNivel01") = CArray(6)
                            RsD.Fields("GlsNivel01") = CArray(7)
                            RsD.Fields("idNivel02") = CArray(4)
                            RsD.Fields("GlsNivel02") = CArray(5)
                            RsD.Fields("idNivel03") = CArray(2)
                            RsD.Fields("GlsNivel03") = CArray(3)
                            RsD.Fields("idNivel04") = CArray(0)
                            RsD.Fields("GlsNivel04") = CArray(1)
                        
                        End If
                        
                    End If
                
                Else
                    
                    RsD.Fields("idNivel01") = ""
                    RsD.Fields("GlsNivel01") = ""
                    RsD.Fields("idNivel02") = ""
                    RsD.Fields("GlsNivel02") = ""
                    RsD.Fields("idNivel03") = ""
                    RsD.Fields("GlsNivel03") = ""
                    RsD.Fields("idNivel04") = ""
                    RsD.Fields("GlsNivel04") = ""
                
                End If
                
                RsD.Fields("CostoPromedio") = 0
                    
                RsTempAntClone.Filter = ""
            
            End If
            
'            dblSaldo = traerCantSaldo(rstemp.Fields("IdProducto"), "" & rstemp.Fields("IdAlmacen"), strFecIni, strMsgError)
'            If strMsgError <> "" Then GoTo ERR
'
'            dblValSaldo = traerCostoUnit(rstemp.Fields("IdProducto"), "" & rstemp.Fields("IdAlmacen"), strFecIni, TxtCod_Moneda.Text, strMsgError)
'            If strMsgError <> "" Then GoTo ERR
'
'            i = 0
'            If (Val(dblSaldo) > 0 Or Val(dblValSaldo) > 0) Then
'                rsd.AddNew
'                dblValSaldoIni = Val(dblValSaldo)
'                dblValSaldo = Val(dblValSaldo) * Val(dblSaldo)
'                rsd.Fields("Item") = i
'                rsd.Fields("idValesCab") = "I"
'                rsd.Fields("fechaEmision") = ""
'                rsd.Fields("IdProducto") = "" & rstemp.Fields("xIdProducto")
'                rsd.Fields("GlsProducto") = rstemp.Fields("GlsProducto")
'                rsd.Fields("IdAlmacen") = "" & rstemp.Fields("IdAlmacen")
'                rsd.Fields("idProvCliente") = ""
'                rsd.Fields("ruc") = ""
'                rsd.Fields("GlsPersona") = ""
'                rsd.Fields("GlsDocReferencia") = "SALDO INICIAL"
'                rsd.Fields("abreUM") = "" & rstemp.Fields("abreUM")
'                rsd.Fields("Ingreso") = Val(dblSaldo)
'                rsd.Fields("Salida") = 0
'                rsd.Fields("Saldo") = Val(dblSaldo)
'                rsd.Fields("ValorUnit") = Val(dblValSaldoIni)
'                rsd.Fields("ValorTotal") = Val(dblSaldo) * Val(dblValSaldoIni)
'                rsd.Fields("ValorSaldo") = Val(dblValSaldo)
'                rsd.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
'                rsd.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
'                rsd.Fields("SimboloMoneda") = strSimboloMoneda
'                rsd.Fields("Empresa") = "" & GlsNom_Empresa
'                rsd.Fields("ruc_empresa") = "" & Glsruc
'                rsd.Fields("Sistema") = "" & glssistema
'                rsd.Fields("Almacen") = "" & rstemp.Fields("Almacen")
'                rsd.Fields("GlsConcepto") = ""
'                If Val(dblSaldo) > 0 Then
'
'                    rsd.Fields("CostoPromedio") = Val(dblValSaldo) / Val(dblSaldo)
'                    DblCostoPAnt = Val(dblValSaldo) / Val(dblSaldo)
'
'                Else
'
'                    rsd.Fields("CostoPromedio") = DblCostoPAnt
'
'                End If
'
'                If Opt2.Value = True Then
'                    If glsNumNiveles = 1 Then
'                        rsd.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel01"))
'                        rsd.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel01"))
'                    ElseIf glsNumNiveles = 2 Then
'                        rsd.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel02"))
'                        rsd.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel02"))
'                        rsd.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel01"))
'                        rsd.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel01"))
'                    ElseIf glsNumNiveles = 3 Then
'                        rsd.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel03"))
'                        rsd.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel03"))
'                        rsd.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel02"))
'                        rsd.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel02"))
'                        rsd.Fields("idNivel03") = Trim("" & rstemp.Fields("idNivel01"))
'                        rsd.Fields("GlsNivel03") = Trim("" & rstemp.Fields("GlsNivel01"))
'                    ElseIf glsNumNiveles = 4 Then
'
'                        rsd.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel04"))
'                        rsd.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel04"))
'                        rsd.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel03"))
'                        rsd.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel03"))
'                        rsd.Fields("idNivel03") = Trim("" & rstemp.Fields("idNivel02"))
'                        rsd.Fields("GlsNivel03") = Trim("" & rstemp.Fields("GlsNivel02"))
'                        rsd.Fields("idNivel04") = Trim("" & rstemp.Fields("idNivel01"))
'                        rsd.Fields("GlsNivel04") = Trim("" & rstemp.Fields("GlsNivel01"))
'
'                    End If
'                End If
'            End If
        End If
        
        RsD.AddNew
        i = i + 1
        RsD.Fields("Item") = i
        If rstemp.Fields("tipoVale") = "S" Then
            RsD.Fields("idValesCab") = "S " & rstemp.Fields("idValesCab")
        Else
            RsD.Fields("idValesCab") = "I  " & rstemp.Fields("idValesCab")
        End If
        RsD.Fields("fechaEmision") = "" & rstemp.Fields("fechaEmision")
        RsD.Fields("IdProducto") = "" & rstemp.Fields("xIdProducto")
        RsD.Fields("GlsProducto") = "" & rstemp.Fields("GlsProducto")
        RsD.Fields("IdAlmacen") = "" & rstemp.Fields("IdAlmacen")
        RsD.Fields("idProvCliente") = "" & rstemp.Fields("idProvCliente")
        RsD.Fields("ruc") = "" & rstemp.Fields("ruc")
        RsD.Fields("GlsPersona") = "" & rstemp.Fields("GlsPersona")
        RsD.Fields("GlsDocReferencia") = "" & rstemp.Fields("GlsDocReferencia")
        RsD.Fields("abreUM") = "" & rstemp.Fields("abreUM")
        RsD.Fields("ValorUnit") = IsNull("" & rstemp.Fields("VVUnit"))
        
'        If dblSaldo < 0 Then
'            dblSaldo = 0
'        End If
'
'        If dblValSaldo < 0 Then
'            dblValSaldo = 0
'        End If
        'If rstemp("idValesCab") = "15030003" Then MsgBox ""
        If rstemp.Fields("tipoVale") = "S" Then
            RsD.Fields("Ingreso") = 0
            RsD.Fields("PesoIngreso") = 0
            RsD.Fields("Salida") = Val(rstemp.Fields("Cantidad"))
            RsD.Fields("PesoSalida") = Val("" & rstemp.Fields("Cantidad")) * Val("" & rstemp.Fields("IdTallaPeso"))
            dblSaldo = Val(dblSaldo) - Val(rstemp.Fields("Cantidad"))
            RsD.Fields("Saldo") = Val(dblSaldo)
            RsD.Fields("PesoSaldo") = Val(dblSaldo) * Val("" & rstemp.Fields("IdTallaPeso"))
            
            RsD.Fields("ValorUnit") = Val(rstemp.Fields("VVUnit"))
            RsD.Fields("ValorTotal") = Val(rstemp.Fields("Cantidad")) * Val(rstemp.Fields("VVUnit"))
            dblValSaldo = Val(dblValSaldo) - Val(RsD.Fields("ValorTotal"))
        Else
            RsD.Fields("Ingreso") = Val(rstemp.Fields("Cantidad"))
            RsD.Fields("PesoIngreso") = Val("" & rstemp.Fields("Cantidad")) * Val("" & rstemp.Fields("IdTallaPeso"))
            RsD.Fields("Salida") = 0
            RsD.Fields("PesoSalida") = 0
            dblSaldo = Val(dblSaldo) + Val(rstemp.Fields("Cantidad"))
            RsD.Fields("Saldo") = Val(dblSaldo)
            RsD.Fields("PesoSaldo") = Val(dblSaldo) * Val("" & rstemp.Fields("IdTallaPeso"))
            RsD.Fields("ValorUnit") = Val(rstemp.Fields("VVUnit"))
            RsD.Fields("ValorTotal") = Val(rstemp.Fields("Cantidad")) * Val(rstemp.Fields("VVUnit"))
            dblValSaldo = Val(dblValSaldo) + Val(RsD.Fields("ValorTotal"))
        End If
            
        RsD.Fields("ValorSaldo") = Val(dblValSaldo)
        RsD.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
        RsD.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
        RsD.Fields("SimboloMoneda") = strSimboloMoneda
        
        If visCodRapido = "S" Then
            CodProd = rstemp.Fields("XIdProducto")
        Else
            CodProd = rstemp.Fields("IdProducto")
        End If
        
        Codalmax = rstemp.Fields("idAlmacen")
        RsD.Fields("Empresa") = "" & GlsNom_Empresa
        RsD.Fields("ruc_empresa") = "" & Glsruc
        RsD.Fields("Sistema") = "" & glssistema
        RsD.Fields("Almacen") = "" & rstemp.Fields("Almacen")
        RsD.Fields("GlsConcepto") = "" & rstemp.Fields("GlsConcepto")
        RsD.Fields("CodigoRapido") = "" & rstemp.Fields("CodigoRapido")
        
        If Opt2.Value = True Then
            If glsNumNiveles = 1 Then
                RsD.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel01"))
                RsD.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel01"))
            ElseIf glsNumNiveles = 2 Then
                RsD.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel02"))
                RsD.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel02"))
                RsD.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel01"))
                RsD.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel01"))
            ElseIf glsNumNiveles = 3 Then
                RsD.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel03"))
                RsD.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel03"))
                RsD.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel02"))
                RsD.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel02"))
                RsD.Fields("idNivel03") = Trim("" & rstemp.Fields("idNivel01"))
                RsD.Fields("GlsNivel03") = Trim("" & rstemp.Fields("GlsNivel01"))
            ElseIf glsNumNiveles = 4 Then
                
                RsD.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel04"))
                RsD.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel04"))
                RsD.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel03"))
                RsD.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel03"))
                RsD.Fields("idNivel03") = Trim("" & rstemp.Fields("idNivel02"))
                RsD.Fields("GlsNivel03") = Trim("" & rstemp.Fields("GlsNivel02"))
                RsD.Fields("idNivel04") = Trim("" & rstemp.Fields("idNivel01"))
                RsD.Fields("GlsNivel04") = Trim("" & rstemp.Fields("GlsNivel01"))
                
            End If
        End If
        
        If Val(dblSaldo) > 0 Then
                    
            RsD.Fields("CostoPromedio") = Val(dblValSaldo) / Val(dblSaldo)
            DblCostoPAnt = Val(dblValSaldo) / Val(dblSaldo)
            
        Else
        
            RsD.Fields("CostoPromedio") = DblCostoPAnt
            ''JACH 28/01/2016
            '' SI LA CANTIDAD ES CERO NO QUEDA SALDO EN MONTOS
            'RsD.Fields("ValorTotal") = Val(RsD.Fields("ValorTotal")) + Val(dblValSaldo)
            'RsD.Fields("ValorSaldo") = 0
            'dblValSaldo = 0
        End If
        
        If visCodRapido = "S" Then
            CodProd = rstemp.Fields("XIdProducto")
        Else
            CodProd = rstemp.Fields("IdProducto")
        End If
        
        Codalmax = rstemp.Fields("idAlmacen")
        rstemp.MoveNext
    Loop
    Set kardex = RsD
    If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
    
    Exit Function
    
Err:
'Resume
    If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
    
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Function
    Resume
End Function

Private Function kardex_Ant(ByRef StrMsgError As String, Optional codproducto As String) As Recordset 'Copia de Procedimiento PQS 110414
On Error GoTo Err
Dim rstemp                      As New ADODB.Recordset
Dim RsD                         As New ADODB.Recordset
Dim rsx                         As New ADODB.Recordset
Dim strSQL                      As String
Dim i                           As Integer
Dim dblSaldo                    As Double
Dim dblValSaldo                 As Double
Dim CodProd                     As String
Dim CodProdAnt                  As String
Dim strFecIni                   As String
Dim strFecFin                   As String
Dim cadenaniveles               As String
Dim strSimboloMoneda            As String
Dim cNiveles                    As String
Dim X                           As Integer
Dim visCodRapido                As String
Dim dblValSaldoIni              As Double
Dim strProAnt                   As String
Dim rstempAnt                   As New ADODB.Recordset
Dim indVerifica                 As Boolean
Dim StrCodalmacenx              As String
Dim Codalmax                    As String
Dim DblCostoPAnt                As Double
Dim CArray(8)                   As String
Dim RsTempAntClone              As ADODB.Recordset

    StrCodalmacenx = ""
    indVerifica = False
    DblCostoPAnt = 0
    
    visCodRapido = Trim("" & traerCampo("parametros", "ValParametro", "GlsParametro", "VIZUALIZA_CODIGO_RAPIDO", True))
    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    cadenaniveles = right(cbodatos.Text, 5)
    If cadenaniveles = "20004" Then
        cadenaniveles = " "
    Else
        cadenaniveles = " and n.tipo =  '" & cadenaniveles & "' "
    End If
    cNiveles = ""
    strSimboloMoneda = txtGls_Moneda.Text
    
    If Opt1.Value = True Then
    
        strSQL = "SELECT vc.idValesCab,vc.idConcepto, " & _
                    "vc.tipoVale," & _
                    "vc.fechaEmision,if('" & visCodRapido & "' = 'S',pr.Codigorapido,vd.IdProducto) xIdProducto, " & _
                    "pr.IdProducto," & _
                    "pr.GlsProducto," & _
                    "vc.IdAlmacen," & _
                    "vc.idProvCliente,pe.ruc,pe.GlsPersona," & _
                    "vc.GlsDocReferencia," & _
                    "um.abreUM," & _
                    "vd.Cantidad, " & _
                    "CASE '" & txtCod_Moneda.Text & "' WHEN 'PEN' THEN  IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  ifnull(vc.TipoCambio,t.tcventa))" & _
                                            "WHEN 'USD' THEN  IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  ifnull(vc.TipoCambio,t.tcventa)) " & _
                    "END as VVUnit,ConCat(Vc.IdAlmacen,' ',A.GlsAlmacen) As Almacen,Z.GlsConcepto,pr.IdTallaPeso "
                       
        strSQL = strSQL & "FROM valescab vc " & _
                    "inner join valesdet vd " & _
                        "on vc.idValesCab = vd.idValesCab AND vc.idEmpresa = vd.idEmpresa AND vc.idSucursal = vd.idSucursal " & _
                    "AND vc.tipoVale = vd.tipoVale " & _
                    "Left Join Conceptos Z On Vc.IdConcepto = Z.IdConcepto " & _
                    "left join personas pe " & _
                        "on vc.idProvCliente = pe.IdPersona " & _
                    "Inner Join Almacenes A " & _
                        "On Vc.IdEmpresa = A.IdEmpresa And Vc.IdAlmacen = A.IdAlmacen And Vc.idSucursal =  A.idSucursal " & _
                    "inner join productos pr " & _
                        "on vd.idProducto = pr.idProducto AND vd.idEmpresa = pr.idEmpresa " & _
                    "inner join unidadmedida um " & _
                        "on pr.idUMCompra = um.idUM " & _
                    "inner join niveles n " & _
                        "on pr.idnivel = n.idnivel and pr.idempresa = n.idempresa  "
        strSQL = strSQL & "Left Join tiposdecambio t " & _
                        "On vc.fechaEmision = t.fecha "
        strSQL = strSQL & "WHERE vc.idEmpresa = '" & glsEmpresa & "' AND pr.idEmpresa = '" & glsEmpresa & "' " & _
                    IIf(Len(Trim(TxtCod_Almacen.Text)) = 0, "", "AND vc.idAlmacen = '" & TxtCod_Almacen.Text & "' ") & _
                    "AND (vc.idPeriodoInv) IN " & _
                    "(" & _
                        "SELECT pi.idPeriodoInv " & _
                        "FROM periodosinv pi " & _
                        "WHERE pi.idEmpresa = vc.idEmpresa AND pi.idSucursal = vc.idSucursal and pi.FecInicio <= '" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' " & _
                        "and (pi.FecFin >= '" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' or pi.FecFin is null)" & _
                    ") " & _
                    "AND vc.estValeCab <> 'ANU'  And pr.estProducto = 'A' AND vc.FechaEmision BETWEEN '" & strFecIni & "' AND '" & strFecFin & "'  " & _
                    cadenaniveles
    Else
    
        For X = 1 To glsNumNiveles
            cNiveles = cNiveles & "vn.idNivel" & Format(X, "00") & ", vn.GlsNivel" & Format(X, "00") & ","
        Next X
    
        strSQL = "SELECT " & cNiveles & " vc.idValesCab," & _
                    "vc.tipoVale," & _
                    "vc.fechaEmision,if('" & visCodRapido & "' = 'S',pr.Codigorapido,vd.IdProducto) xIdProducto, " & _
                    "pr.IdProducto," & _
                    "pr.GlsProducto," & _
                    "vc.IdAlmacen," & _
                    "vc.idProvCliente,pe.ruc,pe.GlsPersona," & _
                    "vc.GlsDocReferencia," & _
                    "um.abreUM," & _
                    "vd.Cantidad, " & _
                    "CASE '" & txtCod_Moneda.Text & "' WHEN 'PEN' THEN  IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  ifnull(vc.TipoCambio,t.tcventa) )" & _
                                            "WHEN 'USD' THEN  IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  ifnull(vc.TipoCambio,t.tcventa) ) " & _
                    "END as VVUnit,ConCat(Vc.IdAlmacen,' ',A.GlsAlmacen) As Almacen,Z.GlsConcepto,pr.IdTallaPeso "
                       
        strSQL = strSQL & "FROM valescab vc " & _
                    "inner join valesdet vd " & _
                        "on vc.idValesCab = vd.idValesCab AND vc.idEmpresa = vd.idEmpresa AND vc.idSucursal = vd.idSucursal " & _
                    "AND vc.tipoVale = vd.tipoVale " & _
                    "Left Join Conceptos Z On Vc.IdConcepto = Z.IdConcepto " & _
                    "left join personas pe " & _
                        "on vc.idProvCliente = pe.IdPersona " & _
                    "Inner Join Almacenes A " & _
                        "On Vc.IdEmpresa = A.IdEmpresa And Vc.IdAlmacen = A.IdAlmacen And Vc.idSucursal =  A.idSucursal " & _
                    "inner join productos pr " & _
                        "on vd.idProducto = pr.idProducto AND vd.idEmpresa = pr.idEmpresa " & _
                    "inner join unidadmedida um " & _
                        "on pr.idUMCompra = um.idUM " & _
                    "inner join niveles n " & _
                        "on pr.idnivel = n.idnivel and pr.idempresa = n.idempresa  " & _
                    "inner join vw_niveles vn on pr.idnivel = vn.idNivel01 and pr.idempresa = vn.idempresa "
        strSQL = strSQL & "Left Join tiposdecambio t " & _
                        "On vc.fechaEmision = t.fecha "
        strSQL = strSQL & "WHERE vc.idEmpresa = '" & glsEmpresa & "' AND pr.idEmpresa = '" & glsEmpresa & "' " & _
                    IIf(Len(Trim(TxtCod_Almacen.Text)) = 0, "", "AND vc.idAlmacen = '" & TxtCod_Almacen.Text & "' ") & _
                    "AND (vc.idPeriodoInv) IN " & _
                    "(" & _
                        "SELECT pi.idPeriodoInv " & _
                        "FROM periodosinv pi " & _
                        "WHERE pi.idEmpresa = vc.idEmpresa AND pi.idSucursal = vc.idSucursal and pi.FecInicio <= '" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' " & _
                        "and (pi.FecFin >= '" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' or pi.FecFin is null)" & _
                    ") " & _
                    "AND vc.estValeCab <> 'ANU'  AND pr.estProducto = 'A' AND vc.FechaEmision BETWEEN '" & strFecIni & "' AND '" & strFecFin & "'  " & _
                    cadenaniveles
    End If
    
    If codproducto <> "" Then
        strSQL = strSQL & "  AND vd.idProducto = '" & codproducto & "'"
    End If
    strSQL = strSQL & " ORDER BY vc.IdAlmacen, vd.IdProducto,vc.FechaEmision,vc.tipovale,vc.idvalescab"
            
    If rstemp.State = 1 Then rstemp.Close
    rstemp.Open strSQL, Cn, adOpenKeyset, adLockOptimistic
    
    Set rstemp.ActiveConnection = Nothing
    
    RsD.Fields.Append "Item", adInteger, , adFldRowID
    RsD.Fields.Append "idValesCab", adVarChar, 15, adFldIsNullable
    RsD.Fields.Append "fechaEmision", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "IdProducto", adVarChar, 50, adFldIsNullable
    RsD.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
    RsD.Fields.Append "IdAlmacen", adVarChar, 8, adFldIsNullable
    RsD.Fields.Append "idProvCliente", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "ruc", adVarChar, 20, adFldIsNullable
    RsD.Fields.Append "GlsPersona", adVarChar, 120, adFldIsNullable
    RsD.Fields.Append "GlsDocReferencia", adVarChar, 180, adFldIsNullable
    RsD.Fields.Append "abreUM", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "Ingreso", adDouble, , adFldIsNullable
    RsD.Fields.Append "Salida", adDouble, , adFldIsNullable
    RsD.Fields.Append "Saldo", adDouble, , adFldIsNullable
    RsD.Fields.Append "ValorUnit", adDouble, , adFldIsNullable
    RsD.Fields.Append "ValorTotal", adDouble, , adFldIsNullable
    RsD.Fields.Append "ValorSaldo", adDouble, , adFldIsNullable
    RsD.Fields.Append "FechaInicio", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "FechaFin", adVarChar, 10, adFldIsNullable
    RsD.Fields.Append "SimboloMoneda", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "Empresa", adVarChar, 255, adFldIsNullable
    RsD.Fields.Append "ruc_empresa", adVarChar, 11, adFldIsNullable
    RsD.Fields.Append "Sistema", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "Almacen", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsConcepto", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "idNivel01", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsNivel01", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "idNivel02", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsNivel02", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "idNivel03", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsNivel03", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "idNivel04", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "GlsNivel04", adVarChar, 200, adFldIsNullable
    RsD.Fields.Append "CostoPromedio", adDouble, , adFldIsNullable
    RsD.Fields.Append "PesoIngreso", adDouble, , adFldIsNullable
    RsD.Fields.Append "PesoSalida", adDouble, , adFldIsNullable
    RsD.Fields.Append "PesoSaldo", adDouble, , adFldIsNullable
    RsD.Open
    
    CodProd = ""
    Codalmax = ""
    i = 0
    dblSaldo = 0
    dblValSaldo = 0
    
    'Set rstempAnt = DataProcedimiento("Spu_ListaSaldoInicialProductos", strMsgError, glsEmpresa, txtCod_Almacen.Text, codproducto, TxtCod_Moneda.Text, Format(dtpfInicio.Value, "yyyy-mm-dd"), Format(dtpFFinal.Value, "yyyy-mm-dd"))
    'If strMsgError <> "" Then GoTo ERR
    
    rstempAnt.Open "Call Spu_ListaSaldoInicialProductos('" & glsEmpresa & "','" & TxtCod_Almacen.Text & "','" & codproducto & "','" & txtCod_Moneda.Text & "','" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "','" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "')", Cn, adOpenStatic, adLockReadOnly
    Set RsTempAntClone = rstempAnt.Clone(adLockReadOnly)
    
'    Do While Not rstempAnt.EOF
'        rsd.AddNew
'        i = 0
'        rsd.Fields("Item") = i
'        rsd.Fields("idValesCab") = "I"
'        rsd.Fields("fechaEmision") = ""
'        rsd.Fields("IdProducto") = "" & rstempAnt.Fields("xIdProducto")
'        rsd.Fields("GlsProducto") = rstempAnt.Fields("GlsProducto")
'        rsd.Fields("IdAlmacen") = "" & rstempAnt.Fields("IdAlmacen")
'        rsd.Fields("idProvCliente") = ""
'        rsd.Fields("ruc") = ""
'        rsd.Fields("GlsPersona") = ""
'        rsd.Fields("GlsDocReferencia") = "SALDO INICIAL"
'        rsd.Fields("abreUM") = "" & rstempAnt.Fields("abreUM")
'        rsd.Fields("Ingreso") = "" & rstempAnt.Fields("Stock")
'        rsd.Fields("Salida") = 0
'        rsd.Fields("Saldo") = "" & rstempAnt.Fields("Stock")
'        rsd.Fields("ValorUnit") = rstempAnt.Fields("VVUnit")
'        rsd.Fields("ValorTotal") = rstempAnt.Fields("SaldoInicial") 'rstempAnt.Fields("Stock") * rstempAnt.Fields("SaldoInicial")
'        rsd.Fields("ValorSaldo") = rstempAnt.Fields("SaldoInicial")
'        rsd.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
'        rsd.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
'        rsd.Fields("SimboloMoneda") = ""
'        rsd.Fields("Empresa") = ""
'        rsd.Fields("ruc_empresa") = ""
'        rsd.Fields("Sistema") = ""
'        rsd.Fields("Almacen") = "" & rstempAnt.Fields("Almacen")
'        rsd.Fields("GlsConcepto") = ""
'
'        If Opt2.Value = True Then
'            If glsNumNiveles = 1 Then
'
'                traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01", "A.IdProducto", Trim("" & rstempAnt.Fields("xIdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
'
'                If Len(Trim(CArray(0))) > 0 Then
'
'                    rsd.Fields("idNivel01") = CArray(0)
'                    rsd.Fields("GlsNivel01") = CArray(1)
'
'                End If
'
'            ElseIf glsNumNiveles = 2 Then
'
'                traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02", "A.IdProducto", Trim("" & rstempAnt.Fields("xIdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
'
'                If Len(Trim(CArray(0))) > 0 Then
'
'                    rsd.Fields("idNivel01") = CArray(2)
'                    rsd.Fields("GlsNivel01") = CArray(3)
'                    rsd.Fields("idNivel02") = CArray(0)
'                    rsd.Fields("GlsNivel02") = CArray(1)
'
'                End If
'
'            ElseIf glsNumNiveles = 3 Then
'
'                traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02,B.IdNivel03,B.GlsNivel03", "A.IdProducto", Trim("" & rstempAnt.Fields("xIdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
'
'                If Len(Trim(CArray(0))) > 0 Then
'
'                    rsd.Fields("idNivel01") = CArray(4)
'                    rsd.Fields("GlsNivel01") = CArray(5)
'                    rsd.Fields("idNivel02") = CArray(2)
'                    rsd.Fields("GlsNivel02") = CArray(3)
'                    rsd.Fields("idNivel03") = CArray(0)
'                    rsd.Fields("GlsNivel03") = CArray(1)
'
'                End If
'
'            ElseIf glsNumNiveles = 4 Then
'
'                traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02,B.IdNivel03,B.GlsNivel03,B.IdNivel04,B.GlsNivel04", "A.IdProducto", Trim("" & rstempAnt.Fields("xIdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
'
'                If Len(Trim(CArray(0))) > 0 Then
'
'                    rsd.Fields("idNivel01") = CArray(6)
'                    rsd.Fields("GlsNivel01") = CArray(7)
'                    rsd.Fields("idNivel02") = CArray(4)
'                    rsd.Fields("GlsNivel02") = CArray(5)
'                    rsd.Fields("idNivel03") = CArray(2)
'                    rsd.Fields("GlsNivel03") = CArray(3)
'                    rsd.Fields("idNivel04") = CArray(0)
'                    rsd.Fields("GlsNivel04") = CArray(1)
'
'                End If
'
'            End If
'        End If
'
''        If Opt2.Value = True Then
''            If glsNumNiveles = 1 Then
''                rsd.Fields("idNivel01") = Trim("" & rstempAnt.Fields("idNivel01"))
''                rsd.Fields("GlsNivel01") = Trim("" & rstempAnt.Fields("GlsNivel01"))
''            ElseIf glsNumNiveles = 2 Then
''                rsd.Fields("idNivel01") = Trim("" & rstempAnt.Fields("idNivel01"))
''                rsd.Fields("GlsNivel01") = Trim("" & rstempAnt.Fields("GlsNivel01"))
''                rsd.Fields("idNivel02") = Trim("" & rstempAnt.Fields("idNivel02"))
''                rsd.Fields("GlsNivel02") = Trim("" & rstempAnt.Fields("GlsNivel02"))
''            ElseIf glsNumNiveles = 3 Then
''                rsd.Fields("idNivel01") = Trim("" & rstempAnt.Fields("idNivel01"))
''                rsd.Fields("GlsNivel01") = Trim("" & rstempAnt.Fields("GlsNivel01"))
''                rsd.Fields("idNivel02") = Trim("" & rstempAnt.Fields("idNivel02"))
''                rsd.Fields("GlsNivel02") = Trim("" & rstempAnt.Fields("GlsNivel02"))
''                rsd.Fields("idNivel03") = Trim("" & rstempAnt.Fields("idNivel03"))
''                rsd.Fields("GlsNivel03") = Trim("" & rstempAnt.Fields("GlsNivel03"))
''             ElseIf glsNumNiveles = 4 Then
''                rsd.Fields("idNivel01") = Trim("" & rstempAnt.Fields("idNivel01"))
''                rsd.Fields("GlsNivel01") = Trim("" & rstempAnt.Fields("GlsNivel01"))
''                rsd.Fields("idNivel02") = Trim("" & rstempAnt.Fields("idNivel02"))
''                rsd.Fields("GlsNivel02") = Trim("" & rstempAnt.Fields("GlsNivel02"))
''                rsd.Fields("idNivel03") = Trim("" & rstempAnt.Fields("idNivel03"))
''                rsd.Fields("GlsNivel03") = Trim("" & rstempAnt.Fields("GlsNivel03"))
''                rsd.Fields("idNivel04") = Trim("" & rstempAnt.Fields("idNivel04"))
''                rsd.Fields("GlsNivel04") = Trim("" & rstempAnt.Fields("GlsNivel04"))
''            End If
''        End If
'
'
'
''        rsd.Fields("ValorUnit") = rstempAnt.Fields("VVUnit")
''        rsd.Fields("ValorTotal") = rstempAnt.Fields("SaldoInicial") 'rstempAnt.Fields("Stock") * rstempAnt.Fields("SaldoInicial")
''        rsd.Fields("ValorSaldo") = rstempAnt.Fields("SaldoInicial")
''
''        If Val("" & rstempAnt.Fields("Stock")) > 0 Then
''
''            rsd.Fields("CostoPromedio") = Val("" & rstempAnt.Fields("SaldoInicial")) / Val("" & rstempAnt.Fields("Stock"))
''            DblCostoPAnt = Val("" & rstempAnt.Fields("SaldoInicial")) / Val("" & rstempAnt.Fields("Stock"))
''
''        Else
''
''            rsd.Fields("CostoPromedio") = DblCostoPAnt
''
''        End If
'
'        rstempAnt.MoveNext
'
'    Loop
   
    Do While Not rstemp.EOF
        If StrCodalmacenx = "" Then
            StrCodalmacenx = Trim("" & rstemp.Fields("IdAlmacen"))
        End If
        
        If StrCodalmacenx <> Trim("" & rstemp.Fields("IdAlmacen")) Then
            StrCodalmacenx = Trim("" & rstemp.Fields("IdAlmacen"))
            dblSaldo = 0
            dblValSaldo = 0
        End If
        
        If CodProd <> Trim("" & rstemp.Fields("IdProducto")) Then
            dblSaldo = 0
            dblValSaldo = 0
        End If
        
        If (CodProd <> rstemp.Fields("IdProducto") And Codalmax <> Trim("" & rstemp.Fields("IdAlmacen"))) Or (CodProd = rstemp.Fields("IdProducto") And Codalmax <> Trim("" & rstemp.Fields("IdAlmacen"))) Or (CodProd <> rstemp.Fields("IdProducto") And Codalmax = Trim("" & rstemp.Fields("IdAlmacen"))) Then
                        
            RsTempAntClone.Filter = "IdAlmacen = '" & Trim("" & rstemp.Fields("IdAlmacen")) & "' And IdProducto = '" & Trim("" & rstemp.Fields("IdProducto")) & "'"
            
            If Not RsTempAntClone.EOF Then
            
                RsD.AddNew
                i = 0
                RsD.Fields("Item") = i
                RsD.Fields("idValesCab") = "I"
                RsD.Fields("fechaEmision") = ""
                RsD.Fields("IdProducto") = "" & RsTempAntClone.Fields("xIdProducto")
                RsD.Fields("GlsProducto") = RsTempAntClone.Fields("GlsProducto")
                RsD.Fields("IdAlmacen") = "" & RsTempAntClone.Fields("IdAlmacen")
                RsD.Fields("idProvCliente") = ""
                RsD.Fields("ruc") = ""
                RsD.Fields("GlsPersona") = ""
                RsD.Fields("GlsDocReferencia") = "SALDO INICIAL"
                RsD.Fields("abreUM") = "" & RsTempAntClone.Fields("abreUM")
                RsD.Fields("Ingreso") = "" & RsTempAntClone.Fields("Stock")
                RsD.Fields("Salida") = 0
                RsD.Fields("Saldo") = "" & RsTempAntClone.Fields("Stock")
                
                RsD.Fields("PesoIngreso") = Val("" & RsTempAntClone.Fields("Stock")) * Val("" & rstemp.Fields("IdTallaPeso"))
                RsD.Fields("PesoSalida") = 0
                RsD.Fields("PesoSaldo") = Val("" & RsTempAntClone.Fields("Stock")) * Val("" & rstemp.Fields("IdTallaPeso"))
                
                RsD.Fields("ValorUnit") = RsTempAntClone.Fields("VVUnit")
                RsD.Fields("ValorTotal") = RsTempAntClone.Fields("SaldoInicial") 'rstempAnt.Fields("Stock") * rstempAnt.Fields("SaldoInicial")
                RsD.Fields("ValorSaldo") = RsTempAntClone.Fields("SaldoInicial")
                RsD.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
                RsD.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
                RsD.Fields("SimboloMoneda") = strSimboloMoneda
                RsD.Fields("Empresa") = ""
                RsD.Fields("ruc_empresa") = ""
                RsD.Fields("Sistema") = ""
                RsD.Fields("Almacen") = "" & RsTempAntClone.Fields("Almacen")
                RsD.Fields("GlsConcepto") = ""
                
                dblSaldo = Val(dblSaldo) + Val("" & RsTempAntClone.Fields("Stock"))
                dblValSaldo = Val(dblValSaldo) + Val("" & RsTempAntClone.Fields("SaldoInicial"))
            
                If Opt2.Value = True Then
                    If glsNumNiveles = 1 Then
                        
                        traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01", "A.IdProducto", Trim("" & RsTempAntClone.Fields("xIdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                        
                        If Len(Trim(CArray(0))) > 0 Then
                        
                            RsD.Fields("idNivel01") = CArray(0)
                            RsD.Fields("GlsNivel01") = CArray(1)
                        
                        End If
                        
                    ElseIf glsNumNiveles = 2 Then
                    
                        traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02", "A.IdProducto", Trim("" & RsTempAntClone.Fields("xIdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                        
                        If Len(Trim(CArray(0))) > 0 Then
                        
                            RsD.Fields("idNivel01") = CArray(2)
                            RsD.Fields("GlsNivel01") = CArray(3)
                            RsD.Fields("idNivel02") = CArray(0)
                            RsD.Fields("GlsNivel02") = CArray(1)
                        
                        End If
                        
                    ElseIf glsNumNiveles = 3 Then
                        
                        traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02,B.IdNivel03,B.GlsNivel03", "A.IdProducto", Trim("" & RsTempAntClone.Fields("xIdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                        
                        If Len(Trim(CArray(0))) > 0 Then
                        
                            RsD.Fields("idNivel01") = CArray(4)
                            RsD.Fields("GlsNivel01") = CArray(5)
                            RsD.Fields("idNivel02") = CArray(2)
                            RsD.Fields("GlsNivel02") = CArray(3)
                            RsD.Fields("idNivel03") = CArray(0)
                            RsD.Fields("GlsNivel03") = CArray(1)
                        
                        End If
                        
                    ElseIf glsNumNiveles = 4 Then
                        
                        traerCampos "Productos A Inner Join VW_Niveles B On A.IdEmpresa = B.IdEmpresa And A.IdNivel = B.IdNivel01", "B.IdNivel01,B.GlsNivel01,B.IdNivel02,B.GlsNivel02,B.IdNivel03,B.GlsNivel03,B.IdNivel04,B.GlsNivel04", "A.IdProducto", Trim("" & RsTempAntClone.Fields("xIdProducto")), 8, CArray, False, "A.IdEmpresa = '" & glsEmpresa & "'"
                        
                        If Len(Trim(CArray(0))) > 0 Then
                        
                            RsD.Fields("idNivel01") = CArray(6)
                            RsD.Fields("GlsNivel01") = CArray(7)
                            RsD.Fields("idNivel02") = CArray(4)
                            RsD.Fields("GlsNivel02") = CArray(5)
                            RsD.Fields("idNivel03") = CArray(2)
                            RsD.Fields("GlsNivel03") = CArray(3)
                            RsD.Fields("idNivel04") = CArray(0)
                            RsD.Fields("GlsNivel04") = CArray(1)
                        
                        End If
                        
                    End If
                End If
                
                RsD.Fields("ValorUnit") = RsTempAntClone.Fields("VVUnit")
                RsD.Fields("ValorTotal") = RsTempAntClone.Fields("SaldoInicial") 'rstempAnt.Fields("Stock") * rstempAnt.Fields("SaldoInicial")
                RsD.Fields("ValorSaldo") = RsTempAntClone.Fields("SaldoInicial")
                        
                If Val("" & RsTempAntClone.Fields("Stock")) > 0 Then
                
                    RsD.Fields("CostoPromedio") = Val("" & RsTempAntClone.Fields("SaldoInicial")) / Val("" & RsTempAntClone.Fields("Stock"))
                    DblCostoPAnt = Val("" & RsTempAntClone.Fields("SaldoInicial")) / Val("" & RsTempAntClone.Fields("Stock"))
                    
                Else
                
                    RsD.Fields("CostoPromedio") = DblCostoPAnt
                
                End If
        
            Else
            
                RsD.AddNew
                i = 0
                RsD.Fields("Item") = i
                RsD.Fields("idValesCab") = "I"
                RsD.Fields("fechaEmision") = ""
                RsD.Fields("IdProducto") = "" & rstemp.Fields("IdProducto")
                RsD.Fields("GlsProducto") = rstemp.Fields("GlsProducto")
                RsD.Fields("IdAlmacen") = "" & rstemp.Fields("IdAlmacen")
                RsD.Fields("idProvCliente") = ""
                RsD.Fields("ruc") = ""
                RsD.Fields("GlsPersona") = ""
                RsD.Fields("GlsDocReferencia") = "SALDO INICIAL"
                RsD.Fields("abreUM") = "" & rstemp.Fields("abreUM")
                RsD.Fields("Ingreso") = 0
                RsD.Fields("Salida") = 0
                RsD.Fields("Saldo") = 0
                RsD.Fields("PesoIngreso") = 0
                RsD.Fields("PesoSalida") = 0
                RsD.Fields("PesoSaldo") = 0
                RsD.Fields("ValorUnit") = 0
                RsD.Fields("ValorTotal") = 0
                RsD.Fields("ValorSaldo") = 0
                RsD.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
                RsD.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
                RsD.Fields("SimboloMoneda") = strSimboloMoneda
                RsD.Fields("Empresa") = ""
                RsD.Fields("ruc_empresa") = ""
                RsD.Fields("Sistema") = ""
                RsD.Fields("Almacen") = "" & rstemp.Fields("Almacen")
                RsD.Fields("GlsConcepto") = ""
                
                RsD.Fields("idNivel01") = ""
                RsD.Fields("GlsNivel01") = ""
                RsD.Fields("idNivel02") = ""
                RsD.Fields("GlsNivel02") = ""
                RsD.Fields("idNivel03") = ""
                RsD.Fields("GlsNivel03") = ""
                RsD.Fields("idNivel04") = ""
                RsD.Fields("GlsNivel04") = ""
                
                RsD.Fields("CostoPromedio") = 0
                    
                RsTempAntClone.Filter = ""
            
            End If
            
'            dblSaldo = traerCantSaldo(rstemp.Fields("IdProducto"), "" & rstemp.Fields("IdAlmacen"), strFecIni, strMsgError)
'            If strMsgError <> "" Then GoTo ERR
'
'            dblValSaldo = traerCostoUnit(rstemp.Fields("IdProducto"), "" & rstemp.Fields("IdAlmacen"), strFecIni, TxtCod_Moneda.Text, strMsgError)
'            If strMsgError <> "" Then GoTo ERR
'
'            i = 0
'            If (Val(dblSaldo) > 0 Or Val(dblValSaldo) > 0) Then
'                rsd.AddNew
'                dblValSaldoIni = Val(dblValSaldo)
'                dblValSaldo = Val(dblValSaldo) * Val(dblSaldo)
'                rsd.Fields("Item") = i
'                rsd.Fields("idValesCab") = "I"
'                rsd.Fields("fechaEmision") = ""
'                rsd.Fields("IdProducto") = "" & rstemp.Fields("xIdProducto")
'                rsd.Fields("GlsProducto") = rstemp.Fields("GlsProducto")
'                rsd.Fields("IdAlmacen") = "" & rstemp.Fields("IdAlmacen")
'                rsd.Fields("idProvCliente") = ""
'                rsd.Fields("ruc") = ""
'                rsd.Fields("GlsPersona") = ""
'                rsd.Fields("GlsDocReferencia") = "SALDO INICIAL"
'                rsd.Fields("abreUM") = "" & rstemp.Fields("abreUM")
'                rsd.Fields("Ingreso") = Val(dblSaldo)
'                rsd.Fields("Salida") = 0
'                rsd.Fields("Saldo") = Val(dblSaldo)
'                rsd.Fields("ValorUnit") = Val(dblValSaldoIni)
'                rsd.Fields("ValorTotal") = Val(dblSaldo) * Val(dblValSaldoIni)
'                rsd.Fields("ValorSaldo") = Val(dblValSaldo)
'                rsd.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
'                rsd.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
'                rsd.Fields("SimboloMoneda") = strSimboloMoneda
'                rsd.Fields("Empresa") = "" & GlsNom_Empresa
'                rsd.Fields("ruc_empresa") = "" & Glsruc
'                rsd.Fields("Sistema") = "" & glssistema
'                rsd.Fields("Almacen") = "" & rstemp.Fields("Almacen")
'                rsd.Fields("GlsConcepto") = ""
'                If Val(dblSaldo) > 0 Then
'
'                    rsd.Fields("CostoPromedio") = Val(dblValSaldo) / Val(dblSaldo)
'                    DblCostoPAnt = Val(dblValSaldo) / Val(dblSaldo)
'
'                Else
'
'                    rsd.Fields("CostoPromedio") = DblCostoPAnt
'
'                End If
'
'                If Opt2.Value = True Then
'                    If glsNumNiveles = 1 Then
'                        rsd.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel01"))
'                        rsd.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel01"))
'                    ElseIf glsNumNiveles = 2 Then
'                        rsd.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel02"))
'                        rsd.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel02"))
'                        rsd.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel01"))
'                        rsd.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel01"))
'                    ElseIf glsNumNiveles = 3 Then
'                        rsd.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel03"))
'                        rsd.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel03"))
'                        rsd.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel02"))
'                        rsd.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel02"))
'                        rsd.Fields("idNivel03") = Trim("" & rstemp.Fields("idNivel01"))
'                        rsd.Fields("GlsNivel03") = Trim("" & rstemp.Fields("GlsNivel01"))
'                    ElseIf glsNumNiveles = 4 Then
'
'                        rsd.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel04"))
'                        rsd.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel04"))
'                        rsd.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel03"))
'                        rsd.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel03"))
'                        rsd.Fields("idNivel03") = Trim("" & rstemp.Fields("idNivel02"))
'                        rsd.Fields("GlsNivel03") = Trim("" & rstemp.Fields("GlsNivel02"))
'                        rsd.Fields("idNivel04") = Trim("" & rstemp.Fields("idNivel01"))
'                        rsd.Fields("GlsNivel04") = Trim("" & rstemp.Fields("GlsNivel01"))
'
'                    End If
'                End If
'            End If
        End If
        
        RsD.AddNew
        i = i + 1
        RsD.Fields("Item") = i
        If rstemp.Fields("tipoVale") = "S" Then
            RsD.Fields("idValesCab") = "S " & rstemp.Fields("idValesCab")
        Else
            RsD.Fields("idValesCab") = "I  " & rstemp.Fields("idValesCab")
        End If
        RsD.Fields("fechaEmision") = "" & rstemp.Fields("fechaEmision")
        RsD.Fields("IdProducto") = "" & rstemp.Fields("xIdProducto")
        RsD.Fields("GlsProducto") = "" & rstemp.Fields("GlsProducto")
        RsD.Fields("IdAlmacen") = "" & rstemp.Fields("IdAlmacen")
        RsD.Fields("idProvCliente") = "" & rstemp.Fields("idProvCliente")
        RsD.Fields("ruc") = "" & rstemp.Fields("ruc")
        RsD.Fields("GlsPersona") = "" & rstemp.Fields("GlsPersona")
        RsD.Fields("GlsDocReferencia") = "" & rstemp.Fields("GlsDocReferencia")
        RsD.Fields("abreUM") = "" & rstemp.Fields("abreUM")
        RsD.Fields("ValorUnit") = IsNull("" & rstemp.Fields("VVUnit"))
        
'        If dblSaldo < 0 Then
'            dblSaldo = 0
'        End If
'
'        If dblValSaldo < 0 Then
'            dblValSaldo = 0
'        End If
                
        If rstemp.Fields("tipoVale") = "S" Then
            RsD.Fields("Ingreso") = 0
            RsD.Fields("PesoIngreso") = 0
            RsD.Fields("Salida") = Val(rstemp.Fields("Cantidad"))
            RsD.Fields("PesoSalida") = Val("" & rstemp.Fields("Cantidad")) * Val("" & rstemp.Fields("IdTallaPeso"))
            dblSaldo = Val(dblSaldo) - Val(rstemp.Fields("Cantidad"))
            RsD.Fields("Saldo") = Val(dblSaldo)
            RsD.Fields("PesoSaldo") = Val(dblSaldo) * Val("" & rstemp.Fields("IdTallaPeso"))
            
            RsD.Fields("ValorUnit") = Val(rstemp.Fields("VVUnit"))
            RsD.Fields("ValorTotal") = Val(rstemp.Fields("Cantidad")) * Val(rstemp.Fields("VVUnit"))
            dblValSaldo = Val(dblValSaldo) - Val(RsD.Fields("ValorTotal"))
        Else
            RsD.Fields("Ingreso") = Val(rstemp.Fields("Cantidad"))
            RsD.Fields("PesoIngreso") = Val("" & rstemp.Fields("Cantidad")) * Val("" & rstemp.Fields("IdTallaPeso"))
            RsD.Fields("Salida") = 0
            RsD.Fields("PesoSalida") = 0
            dblSaldo = Val(dblSaldo) + Val(rstemp.Fields("Cantidad"))
            RsD.Fields("Saldo") = Val(dblSaldo)
            RsD.Fields("PesoSaldo") = Val(dblSaldo) * Val("" & rstemp.Fields("IdTallaPeso"))
            RsD.Fields("ValorUnit") = Val(rstemp.Fields("VVUnit"))
            RsD.Fields("ValorTotal") = Val(rstemp.Fields("Cantidad")) * Val(rstemp.Fields("VVUnit"))
            dblValSaldo = Val(dblValSaldo) + Val(RsD.Fields("ValorTotal"))
        End If
            
        RsD.Fields("ValorSaldo") = Val(dblValSaldo)
        RsD.Fields("FechaInicio") = Format(strFecIni, "dd/mm/yyyy")
        RsD.Fields("FechaFin") = Format(strFecFin, "dd/mm/yyyy")
        RsD.Fields("SimboloMoneda") = strSimboloMoneda
        
        CodProd = rstemp.Fields("IdProducto")
        Codalmax = rstemp.Fields("idAlmacen")
        RsD.Fields("Empresa") = "" & GlsNom_Empresa
        RsD.Fields("ruc_empresa") = "" & Glsruc
        RsD.Fields("Sistema") = "" & glssistema
        RsD.Fields("Almacen") = "" & rstemp.Fields("Almacen")
        RsD.Fields("GlsConcepto") = "" & rstemp.Fields("GlsConcepto")
        
        If Opt2.Value = True Then
            If glsNumNiveles = 1 Then
                RsD.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel01"))
                RsD.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel01"))
            ElseIf glsNumNiveles = 2 Then
                RsD.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel02"))
                RsD.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel02"))
                RsD.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel01"))
                RsD.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel01"))
            ElseIf glsNumNiveles = 3 Then
                RsD.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel03"))
                RsD.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel03"))
                RsD.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel02"))
                RsD.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel02"))
                RsD.Fields("idNivel03") = Trim("" & rstemp.Fields("idNivel01"))
                RsD.Fields("GlsNivel03") = Trim("" & rstemp.Fields("GlsNivel01"))
            ElseIf glsNumNiveles = 4 Then
                
                RsD.Fields("idNivel01") = Trim("" & rstemp.Fields("idNivel04"))
                RsD.Fields("GlsNivel01") = Trim("" & rstemp.Fields("GlsNivel04"))
                RsD.Fields("idNivel02") = Trim("" & rstemp.Fields("idNivel03"))
                RsD.Fields("GlsNivel02") = Trim("" & rstemp.Fields("GlsNivel03"))
                RsD.Fields("idNivel03") = Trim("" & rstemp.Fields("idNivel02"))
                RsD.Fields("GlsNivel03") = Trim("" & rstemp.Fields("GlsNivel02"))
                RsD.Fields("idNivel04") = Trim("" & rstemp.Fields("idNivel01"))
                RsD.Fields("GlsNivel04") = Trim("" & rstemp.Fields("GlsNivel01"))
                
            End If
        End If
        
        If Val(dblSaldo) > 0 Then
                    
            RsD.Fields("CostoPromedio") = Val(dblValSaldo) / Val(dblSaldo)
            DblCostoPAnt = Val(dblValSaldo) / Val(dblSaldo)
            
        Else
        
            RsD.Fields("CostoPromedio") = DblCostoPAnt
            
        End If
                
        CodProd = rstemp.Fields("IdProducto")
        Codalmax = rstemp.Fields("idAlmacen")
        rstemp.MoveNext
    Loop
    Set kardex_Ant = RsD
    If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
    
    Exit Function
    
Err:
'Resume
    If rstemp.State = 1 Then rstemp.Close: Set rstemp = Nothing
    
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Function
    Resume
End Function

Private Function traerCantSaldo(ByVal codproducto As String, ByVal codalmacen As String, ByVal strFecha As String, ByRef StrMsgError As String) As Double
On Error GoTo Err
Dim rsSaldo As New ADODB.Recordset
Dim strSQL As String

    traerCantSaldo = 0
    strSQL = "SELECT  Round(SUM(IF(vc.tipoVale = 'I',vd.Cantidad," & _
                                            "(vd.Cantidad * -1)" & _
                           ")" & _
                        "),3)   AS STOCK " & _
            "FROM valescab vc,valesdet vd " & _
            "WHERE vc.idValesCab = vd.idValesCab " & _
            "AND vc.idEmpresa = vd.idEmpresa " & _
            "AND vc.idSucursal = vd.idSucursal " & _
            "AND vc.tipoVale = vd.tipoVale " & _
            "AND vc.idEmpresa = '" & glsEmpresa & "'" & _
            "AND vc.IdAlmacen = '" & codalmacen & "' " & _
            "AND vd.idProducto = '" & codproducto & "' " & _
            "AND (vc.idPeriodoInv) IN " & _
                    "(" & _
                        "SELECT pi.idPeriodoInv " & _
                        "FROM periodosinv pi " & _
                        "WHERE pi.idEmpresa = vc.idEmpresa AND pi.idSucursal = vc.idSucursal and pi.FecInicio <= '" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' " & _
                        "and (pi.FecFin >= '" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' or pi.FecFin is null)" & _
                    ") " & _
            "AND vc.fechaEmision < '" & strFecha & "' " & _
            "AND vc.estValeCab <> 'ANU' Order By FechaEmision "
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

    csql = "Select (SUM(IF((valescab.tipoVale = 'I'),(valesdet.Cantidad),((valesdet.Cantidad) * -(1))) * " & _
            "CASE 'PEN' " & _
            "WHEN 'PEN' THEN IF(valescab.idMoneda = 'PEN', valesdet.VVUnit,valesdet.VVUnit * ValesCab.TipoCambio) " & _
            "WHEN 'USD' THEN IF(valescab.idMoneda = 'USD', valesdet.VVUnit,valesdet.VVUnit / ValesCab.TipoCambio) " & _
            "End) / " & _
            "Round(SUM(IF((valescab.tipoVale = 'I'),(valesdet.Cantidad),((valesdet.Cantidad) * -(1)))),3)) AS COSTO_UNITARIO "
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
                        "WHERE pi.idEmpresa = valescab.idEmpresa AND pi.idSucursal = valescab.idSucursal and pi.FecInicio <= '" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' " & _
                        "and (pi.FecFin >= '" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' or pi.FecFin is null)" & _
                    ") "
    csql = csql & " AND valescab.fechaEmision < '" & PFecha & "' And valesdet.idProducto = '" & codproducto & "' "
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
    If rsj.State = 1 Then rsj.Close
    Set rsj = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
