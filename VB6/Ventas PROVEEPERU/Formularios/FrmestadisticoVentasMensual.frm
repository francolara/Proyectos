VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmestadisticoVentasMensual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadístico de Ventas Mensual"
   ClientHeight    =   7695
   ClientLeft      =   3975
   ClientTop       =   1380
   ClientWidth     =   9270
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
   ScaleHeight     =   7695
   ScaleWidth      =   9270
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   450
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7065
      Width           =   1290
   End
   Begin VB.CommandButton BtnSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7065
      Width           =   1290
   End
   Begin VB.Frame FrmestadisticoVentasMensual 
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
      Height          =   855
      Left            =   270
      TabIndex        =   24
      Top             =   5760
      Width           =   8715
      Begin VB.CheckBox ChkCantidad 
         Caption         =   "Cantidad"
         Height          =   285
         Left            =   5355
         TabIndex        =   16
         Top             =   360
         Width           =   1230
      End
      Begin VB.CheckBox ChkIGV 
         Caption         =   "IGV"
         Height          =   285
         Left            =   7155
         TabIndex        =   17
         Top             =   360
         Width           =   690
      End
      Begin VB.OptionButton OptFamilia 
         Caption         =   "Sin Familia"
         Height          =   285
         Index           =   1
         Left            =   3375
         TabIndex        =   15
         Top             =   360
         Width           =   1320
      End
      Begin VB.OptionButton OptFamilia 
         Caption         =   "Con Familia"
         Height          =   285
         Index           =   0
         Left            =   1395
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame frmReporteStockVentasRotacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6900
      Left            =   45
      TabIndex        =   20
      Top             =   0
      Width           =   9150
      Begin VB.CommandButton cmbAyudaZona 
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
         Left            =   8500
         Picture         =   "FrmestadisticoVentasMensual.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   3420
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaCliente 
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
         Left            =   8500
         Picture         =   "FrmestadisticoVentasMensual.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3015
         Width           =   390
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
         Left            =   8500
         Picture         =   "FrmestadisticoVentasMensual.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2610
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaVendedor 
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
         Left            =   8500
         Picture         =   "FrmestadisticoVentasMensual.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2205
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
         Left            =   8500
         Picture         =   "FrmestadisticoVentasMensual.frx":0E28
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1845
         Width           =   390
      End
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
         Left            =   8500
         Picture         =   "FrmestadisticoVentasMensual.frx":11B2
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1260
         Width           =   390
      End
      Begin VB.CheckBox ChkOficial 
         Caption         =   "Documentos"
         Height          =   240
         Left            =   6930
         TabIndex        =   3
         Top             =   1575
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Frame fraContenido 
         Appearance      =   0  'Flat
         Caption         =   " Filtros "
         ForeColor       =   &H00000000&
         Height          =   1635
         Left            =   225
         TabIndex        =   25
         Top             =   3960
         Width           =   8715
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
            Height          =   360
            Left            =   45
            TabIndex        =   26
            Top             =   270
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
               Index           =   4
               Left            =   8100
               Picture         =   "FrmestadisticoVentasMensual.frx":153C
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   1440
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
               Picture         =   "FrmestadisticoVentasMensual.frx":18C6
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   1080
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
               Picture         =   "FrmestadisticoVentasMensual.frx":1C50
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   720
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
               Picture         =   "FrmestadisticoVentasMensual.frx":1FDA
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   360
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
               Left            =   8100
               Picture         =   "FrmestadisticoVentasMensual.frx":2364
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   0
               Width           =   390
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   0
               Left            =   1305
               TabIndex        =   9
               Tag             =   "TidNivelPred"
               Top             =   30
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
               Container       =   "FrmestadisticoVentasMensual.frx":26EE
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   0
               Left            =   2280
               TabIndex        =   32
               Top             =   30
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
               Container       =   "FrmestadisticoVentasMensual.frx":270A
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   1
               Left            =   1305
               TabIndex        =   10
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
               Container       =   "FrmestadisticoVentasMensual.frx":2726
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   1
               Left            =   2280
               TabIndex        =   33
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
               Container       =   "FrmestadisticoVentasMensual.frx":2742
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   2
               Left            =   1305
               TabIndex        =   11
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
               Container       =   "FrmestadisticoVentasMensual.frx":275E
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   2
               Left            =   2280
               TabIndex        =   34
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
               Container       =   "FrmestadisticoVentasMensual.frx":277A
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   3
               Left            =   1305
               TabIndex        =   12
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
               Container       =   "FrmestadisticoVentasMensual.frx":2796
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   3
               Left            =   2280
               TabIndex        =   35
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
               Container       =   "FrmestadisticoVentasMensual.frx":27B2
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   4
               Left            =   1305
               TabIndex        =   13
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
               Container       =   "FrmestadisticoVentasMensual.frx":27CE
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   4
               Left            =   2280
               TabIndex        =   36
               Top             =   1470
               Width           =   5790
               _ExtentX        =   10213
               _ExtentY        =   556
               BackColor       =   12648447
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
               Container       =   "FrmestadisticoVentasMensual.frx":27EA
               Vacio           =   -1  'True
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   135
               TabIndex        =   41
               Top             =   1485
               Width           =   345
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   135
               TabIndex        =   40
               Top             =   1125
               Width           =   390
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   135
               TabIndex        =   39
               Top             =   765
               Width           =   345
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   135
               TabIndex        =   38
               Top             =   405
               Width           =   390
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   135
               TabIndex        =   37
               Top             =   45
               Width           =   345
            End
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   1
         Left            =   225
         TabIndex        =   21
         Top             =   225
         Width           =   8715
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   2430
            TabIndex        =   0
            Top             =   315
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   143130625
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   5355
            TabIndex        =   1
            Top             =   315
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   143130625
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4770
            TabIndex        =   23
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1755
            TabIndex        =   22
            Top             =   360
            Width           =   465
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1260
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
         Container       =   "FrmestadisticoVentasMensual.frx":2806
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2150
         TabIndex        =   43
         Top             =   1260
         Width           =   6350
         _ExtentX        =   11192
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
         Container       =   "FrmestadisticoVentasMensual.frx":2822
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   1215
         TabIndex        =   4
         Tag             =   "TidMoneda"
         Top             =   1845
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
         Container       =   "FrmestadisticoVentasMensual.frx":283E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   2150
         TabIndex        =   46
         Top             =   1845
         Width           =   6350
         _ExtentX        =   11192
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
         Container       =   "FrmestadisticoVentasMensual.frx":285A
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   315
         Left            =   1215
         TabIndex        =   5
         Tag             =   "TidPerCliente"
         Top             =   2250
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
         Container       =   "FrmestadisticoVentasMensual.frx":2876
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   315
         Left            =   2150
         TabIndex        =   49
         Tag             =   "TGlsCliente"
         Top             =   2250
         Width           =   6350
         _ExtentX        =   11192
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
         Locked          =   -1  'True
         Container       =   "FrmestadisticoVentasMensual.frx":2892
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1215
         TabIndex        =   6
         Tag             =   "TidMoneda"
         Top             =   2640
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
         Container       =   "FrmestadisticoVentasMensual.frx":28AE
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2150
         TabIndex        =   52
         Top             =   2640
         Width           =   6350
         _ExtentX        =   11192
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
         Container       =   "FrmestadisticoVentasMensual.frx":28CA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1215
         TabIndex        =   7
         Tag             =   "TidMoneda"
         Top             =   3030
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
         Container       =   "FrmestadisticoVentasMensual.frx":28E6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2150
         TabIndex        =   55
         Top             =   3030
         Width           =   6350
         _ExtentX        =   11192
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
         Container       =   "FrmestadisticoVentasMensual.frx":2902
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Zona 
         Height          =   315
         Left            =   1215
         TabIndex        =   8
         Tag             =   "TidMoneda"
         Top             =   3435
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
         Container       =   "FrmestadisticoVentasMensual.frx":291E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Zona 
         Height          =   315
         Left            =   2150
         TabIndex        =   58
         Top             =   3435
         Width           =   6350
         _ExtentX        =   11192
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
         Container       =   "FrmestadisticoVentasMensual.frx":293A
         Vacio           =   -1  'True
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Zona"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   59
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   56
         Top             =   3075
         Width           =   480
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   53
         Top             =   2715
         Width           =   570
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   50
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   47
         Top             =   1875
         Width           =   645
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   44
         Top             =   1305
         Width           =   645
      End
   End
End
Attribute VB_Name = "FrmestadisticoVentasMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i               As Integer
Dim rsj             As New ADODB.Recordset

Private Sub btnSalir_Click()
    
    Unload Me

End Sub

Private Sub ChkCantidad_Click()

    ChkIGV.Value = 0
 
End Sub

Private Sub ChkIGV_Click()
    
    ChkCantidad.Value = 0

End Sub

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda

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

Private Sub cmbAyudaVendedor_Click()
    
    mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor

End Sub

Private Sub cmbAyudaZona_Click()
    
    mostrarAyuda "ZONA", txtCod_Zona, txtGls_Zona

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError             As String
Dim cNiveles                As String
Dim X                       As Integer
Dim CGlsReporte             As String
Dim cWhereNiveles           As String
Dim zonas                   As String
      
    If Len(Trim(txtCod_Nivel(0).Text)) > 0 Then
        cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles, "00") & " = ''" & txtCod_Nivel(0).Text & "'' "
        If Len(Trim(txtCod_Nivel(1).Text)) > 0 Then
            cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles - 1, "00") & "   = ''" & txtCod_Nivel(1).Text & "'' "
            If Len(Trim(txtCod_Nivel(2).Text)) > 0 Then
                cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles - 2, "00") & "  = ''" & txtCod_Nivel(2).Text & "'' "
            End If
        End If
    End If
    zonas = ""
    If Len(Trim(txtCod_Zona.Text)) > 0 Then
        zonas = "And ub.IdZona=''" & txtCod_Zona.Text & "'' "
    End If
    
    For X = 1 To glsNumNiveles
        cNiveles = cNiveles & "vn.idNivel" & Format(X, "00") & ", vn.GlsNivel" & Format(X, "00") & ","
    Next X
    
    If OptFamilia(0).Value Then
        If ChkIGV.Value = 0 And ChkCantidad.Value = 0 Then
            CGlsReporte = "rptEstadisticoMensual" & Format(glsNumNiveles, "00") & ".rpt"
            mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parProducto|parMoneda|parFecDesde|parFecHasta|parCliente|parOficial|parNiveles|parNiveles1|parZona", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Vendedor.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & cNiveles & "|" & cWhereNiveles & "|" & zonas, "Detalle por Producto Mensual", StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
         
         Else
            If ChkIGV.Value = 1 And ChkCantidad.Value = 0 Then
                If ChkIGV.Value = 1 Then
                    CGlsReporte = "rptEstadisticoMensual" & Format(glsNumNiveles, "00") & "IGV" & ".rpt"
                Else
                    CGlsReporte = "rptEstadisticoMensual" & Format(glsNumNiveles, "00") & ".rpt"
                End If
                mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parProducto|parMoneda|parFecDesde|parFecHasta|parCliente|parOficial|parNiveles|parNiveles1|parZona", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Vendedor.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & cNiveles & "|" & cWhereNiveles & "|" & zonas, "Detalle por Producto Mensual", StrMsgError
                If StrMsgError <> "" Then GoTo Err
                Exit Sub
            
            Else
                If ChkCantidad.Value = 1 Then
                   CGlsReporte = "rptEstadisticoMensual" & Format(glsNumNiveles, "00") & "Cantidad" & ".rpt"
                End If
                mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parProducto|parMoneda|parFecDesde|parFecHasta|parCliente|parOficial|parNiveles|parNiveles1|parZona", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Vendedor.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & cNiveles & "|" & cWhereNiveles & "|" & zonas, "Detalle por Producto Mensual", StrMsgError
                If StrMsgError <> "" Then GoTo Err
                Exit Sub
            End If
        End If
    
    Else
        If ChkIGV.Value = 0 And ChkCantidad.Value = 0 Then
            CGlsReporte = "rptEstadisticoMensual.rpt"
            mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parProducto|parMoneda|parFecDesde|parFecHasta|parCliente|parOficial|parNiveles|parNiveles1|parZona", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Vendedor.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & cNiveles & "|" & cWhereNiveles & "|" & zonas, "Detalle por Producto Mensual", StrMsgError
            If StrMsgError <> "" Then GoTo Err
            Exit Sub
            
        Else
            If ChkIGV.Value = 1 And ChkCantidad.Value = 0 Then
                If ChkIGV.Value = 1 Then
                    CGlsReporte = "rptEstadisticoMensualIGV.rpt"
                Else
                    CGlsReporte = "rptEstadisticoMensual.rpt"
                End If
                mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parProducto|parMoneda|parFecDesde|parFecHasta|parCliente|parOficial|parNiveles|parNiveles1|parZona", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Vendedor.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & cNiveles & "|" & cWhereNiveles & "|" & zonas, "Detalle por Producto Mensual", StrMsgError
                If StrMsgError <> "" Then GoTo Err
                Exit Sub
            
            Else
                If ChkCantidad.Value = 1 Then
                   CGlsReporte = "rptEstadisticoMensualCantidad.rpt"
                End If
                mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parProducto|parMoneda|parFecDesde|parFecHasta|parCliente|parOficial|parNiveles|parNiveles1|parZona", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Vendedor.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "|" & Trim(txtCod_Cliente.Text) & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & cNiveles & "|" & cWhereNiveles & "|" & zonas, "Detalle por Producto Mensual", StrMsgError
                If StrMsgError <> "" Then GoTo Err
                Exit Sub
            End If
        End If
    End If

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    txtGls_Zona.Text = "TODAS LAS ZONAS"
    txtCod_Moneda.Text = "PEN"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    ChkOficial.Visible = IIf(GlsVisualiza_Filtro_Documento = "S", True, False)
    
    txtGls_Nivel(0).Text = "TODOS LOS GRUPOS"
    txtGls_Nivel(1).Text = "TODAS LAS CATEGORIAS"
    txtGls_Nivel(2).Text = "TODAS LAS SUB CATEGORIAS"
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_Serie_LostFocus()
    
    txt_serie.Text = Format(txt_serie.Text, "0000")

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
        txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    End If

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Private Sub txtCod_Producto_Change()
    
    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", "idProducto", txtCod_Producto.Text, True)
    Else
        txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    End If

End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    End If

End Sub

Private Sub txtCod_Vendedor_Change()
    
    If txtCod_Vendedor.Text = "" Then
        txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    Else
        txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)
    End If

End Sub

Private Sub txtCod_Zona_Change()
    
    If txtCod_Zona.Text = "" Then
        txtGls_Zona.Text = "TODAS LAS ZONAS"
    Else
        txtGls_Zona.Text = traerCampo("Zonas", "GlsZona", "idZona", txtCod_Zona.Text, False)
    End If

End Sub

Private Sub mostrarNiveles(ByRef StrMsgError As String)
On Error GoTo Err

    rsj.Open "SELECT GlsTipoNivel FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' Order BY Peso ASC", Cn, adOpenForwardOnly, adLockReadOnly
    fraNivel.Height = 355 * glsNumNiveles
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

Private Sub txtCod_Nivel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        If txtCod_Nivel(0).Text <> "" Then
            txtCod_Nivel(0).Text = ""
            txtGls_Nivel(0).Text = "TODOS LOS GRUPOS"
            
            txtCod_Nivel(1).Text = ""
            txtGls_Nivel(1).Text = "TODAS LAS CATEGORIAS"
            
            txtCod_Nivel(2).Text = ""
            txtGls_Nivel(2).Text = "TODAS LAS SUB CATEGORIAS"
            Exit Sub
        End If
    End If

End Sub
