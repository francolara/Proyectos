VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmRepResumenStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Resumen de Stock"
   ClientHeight    =   4830
   ClientLeft      =   6300
   ClientTop       =   3375
   ClientWidth     =   7350
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
   ScaleHeight     =   4830
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4230
      Width           =   1230
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4230
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   7260
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
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1035
         Width           =   6960
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
            TabIndex        =   17
            Top             =   120
            Width           =   6690
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
               Picture         =   "FrmRepResumenStock.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   22
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
               Left            =   8100
               Picture         =   "FrmRepResumenStock.frx":038A
               Style           =   1  'Graphical
               TabIndex        =   21
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
               Left            =   8100
               Picture         =   "FrmRepResumenStock.frx":0714
               Style           =   1  'Graphical
               TabIndex        =   20
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
               Left            =   8100
               Picture         =   "FrmRepResumenStock.frx":0A9E
               Style           =   1  'Graphical
               TabIndex        =   19
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
               Picture         =   "FrmRepResumenStock.frx":0E28
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   45
               Width           =   390
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   0
               Left            =   1305
               TabIndex        =   23
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
               Container       =   "FrmRepResumenStock.frx":11B2
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   0
               Left            =   2280
               TabIndex        =   24
               Top             =   45
               Width           =   3945
               _ExtentX        =   6959
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
               Container       =   "FrmRepResumenStock.frx":11CE
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   1
               Left            =   1305
               TabIndex        =   25
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
               Container       =   "FrmRepResumenStock.frx":11EA
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   1
               Left            =   2280
               TabIndex        =   26
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
               Container       =   "FrmRepResumenStock.frx":1206
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   2
               Left            =   1305
               TabIndex        =   27
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
               Container       =   "FrmRepResumenStock.frx":1222
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   2
               Left            =   2280
               TabIndex        =   28
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
               Container       =   "FrmRepResumenStock.frx":123E
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   3
               Left            =   1305
               TabIndex        =   29
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
               Container       =   "FrmRepResumenStock.frx":125A
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   3
               Left            =   2280
               TabIndex        =   30
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
               Container       =   "FrmRepResumenStock.frx":1276
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   4
               Left            =   1305
               TabIndex        =   31
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
               Container       =   "FrmRepResumenStock.frx":1292
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   4
               Left            =   2280
               TabIndex        =   32
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
               Container       =   "FrmRepResumenStock.frx":12AE
               Vacio           =   -1  'True
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel:"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   135
               TabIndex        =   37
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
               Left            =   135
               TabIndex        =   36
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
               Left            =   135
               TabIndex        =   35
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
               Left            =   135
               TabIndex        =   34
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
               Left            =   135
               TabIndex        =   33
               Top             =   45
               Width           =   345
            End
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   165
         Width           =   6960
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   0
            Top             =   300
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
            Format          =   145883137
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
            TabIndex        =   1
            Top             =   300
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
            Format          =   145883137
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3915
            TabIndex        =   15
            Top             =   375
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   945
            TabIndex        =   14
            Top             =   375
            Width           =   465
         End
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
         Left            =   6735
         Picture         =   "FrmRepResumenStock.frx":12CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3495
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
         Left            =   6735
         Picture         =   "FrmRepResumenStock.frx":1654
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3105
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Almacen 
         Height          =   315
         Left            =   975
         TabIndex        =   2
         Tag             =   "TidAlmacen"
         Top             =   3120
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
         Container       =   "FrmRepResumenStock.frx":19DE
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Almacen 
         Height          =   315
         Left            =   1935
         TabIndex        =   8
         Top             =   3120
         Width           =   4770
         _ExtentX        =   8414
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
         Container       =   "FrmRepResumenStock.frx":19FA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   975
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   3495
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
         Container       =   "FrmRepResumenStock.frx":1A16
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   1935
         TabIndex        =   11
         Top             =   3495
         Width           =   4770
         _ExtentX        =   8414
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
         Container       =   "FrmRepResumenStock.frx":1A32
         Vacio           =   -1  'True
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Producto"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   165
         TabIndex        =   12
         Top             =   3525
         Width           =   765
      End
      Begin VB.Label lbl_Almacen 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   165
         TabIndex        =   9
         Top             =   3150
         Width           =   630
      End
   End
End
Attribute VB_Name = "FrmRepResumenStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaAlmacen_Click()
Dim strCondicion As String
    
    mostrarAyuda "ALMACENVTA", txtCod_Almacen, txtGls_Almacen, strCondicion

End Sub

Private Sub cmbAyudaProducto_Click()
    
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError             As String
Dim Fini                    As String
Dim Ffin                    As String
Dim cWhereNiveles           As String
    
    If Len(Trim(txtCod_Nivel(0).Text)) > 0 Then
        cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles, "00") & " = ''" & txtCod_Nivel(0).Text & "'' "
        If Len(Trim(txtCod_Nivel(1).Text)) > 0 Then
            cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles - 1, "00") & "   = ''" & txtCod_Nivel(1).Text & "'' "
            If Len(Trim(txtCod_Nivel(2).Text)) > 0 Then
                cWhereNiveles = cWhereNiveles & "And vn.idNivel" & Format(glsNumNiveles - 2, "00") & "  = ''" & txtCod_Nivel(2).Text & "'' "
            End If
        End If
    End If
    
    Fini = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    mostrarReporte "RepResumenStock.rpt", "parEmpresa|parAlmacen|parProducto|parFecDesde|parFecHasta|parNiveles", glsEmpresa & "|" & Trim(txtCod_Almacen.Text) & "|" & Trim(txtCod_Producto.Text) & "|" & Fini & "|" & Ffin & "|" & cWhereNiveles, "Reporte de Resumen de Stock", StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
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
Dim StrMsgError                                 As String

    Me.top = 0
    Me.left = 0
    
    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    txtGls_Almacen.Text = "TODOS LOS ALMACENES"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Almacen_Change()
    
    If txtCod_Almacen.Text <> "" Then
        txtGls_Almacen.Text = traerCampo("almacenes", "GlsAlmacen", "idAlmacen", txtCod_Almacen.Text, True)
    Else
        txtGls_Almacen.Text = "TODOS LOS ALMACENES"
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
        mostrarAyudaKeyascii KeyAscii, "ALMACENVTA", txtCod_Almacen, txtGls_Almacen, strCondicion
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtCod_Producto_Change()

    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", "idProducto", txtCod_Producto.Text, True)
    Else
        txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    End If
    
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
