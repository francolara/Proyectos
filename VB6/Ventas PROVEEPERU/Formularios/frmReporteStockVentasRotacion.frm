VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmReporteStockVentasRotacion 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Stock - Ventas - Rotacion"
   ClientHeight    =   4245
   ClientLeft      =   2130
   ClientTop       =   3135
   ClientWidth     =   9165
   DrawMode        =   14  'Copy Pen
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
   ScaleHeight     =   4245
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
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
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   180
      TabIndex        =   30
      Top             =   2700
      Width           =   8790
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
         Left            =   8190
         Picture         =   "frmReporteStockVentasRotacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   225
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   1350
         TabIndex        =   32
         Tag             =   "TidMoneda"
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
         Container       =   "frmReporteStockVentasRotacion.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   2340
         TabIndex        =   33
         Top             =   225
         Width           =   5805
         _ExtentX        =   10239
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
         Container       =   "frmReporteStockVentasRotacion.frx":03A6
         Vacio           =   -1  'True
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   135
         TabIndex        =   34
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3700
      Width           =   1230
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3700
      Width           =   1230
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
      Height          =   3570
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9105
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   0
         Left            =   135
         TabIndex        =   23
         Top             =   225
         Width           =   8805
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   3015
            TabIndex        =   24
            Top             =   270
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
            Format          =   107610113
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   5895
            TabIndex        =   25
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
            Format          =   107610113
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5355
            TabIndex        =   27
            Top             =   345
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2295
            TabIndex        =   26
            Top             =   345
            Width           =   465
         End
      End
      Begin VB.Frame fraContenido 
         Appearance      =   0  'Flat
         Caption         =   " Filtros "
         ForeColor       =   &H00000000&
         Height          =   1635
         Left            =   135
         TabIndex        =   1
         Top             =   990
         Width           =   8790
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
            Height          =   345
            Left            =   45
            TabIndex        =   2
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
               Index           =   0
               Left            =   8100
               Picture         =   "frmReporteStockVentasRotacion.frx":03C2
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   0
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
               Picture         =   "frmReporteStockVentasRotacion.frx":074C
               Style           =   1  'Graphical
               TabIndex        =   6
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
               Index           =   2
               Left            =   8100
               Picture         =   "frmReporteStockVentasRotacion.frx":0AD6
               Style           =   1  'Graphical
               TabIndex        =   5
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
               Index           =   3
               Left            =   8100
               Picture         =   "frmReporteStockVentasRotacion.frx":0E60
               Style           =   1  'Graphical
               TabIndex        =   4
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
               Index           =   4
               Left            =   8100
               Picture         =   "frmReporteStockVentasRotacion.frx":11EA
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   1440
               Width           =   390
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   0
               Left            =   1305
               TabIndex        =   8
               Tag             =   "TidNivelPred"
               Top             =   0
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
               Container       =   "frmReporteStockVentasRotacion.frx":1574
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   0
               Left            =   2280
               TabIndex        =   9
               Top             =   0
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
               Container       =   "frmReporteStockVentasRotacion.frx":1590
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
               Container       =   "frmReporteStockVentasRotacion.frx":15AC
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   1
               Left            =   2280
               TabIndex        =   11
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
               Container       =   "frmReporteStockVentasRotacion.frx":15C8
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   2
               Left            =   1305
               TabIndex        =   12
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
               Container       =   "frmReporteStockVentasRotacion.frx":15E4
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   2
               Left            =   2280
               TabIndex        =   13
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
               Container       =   "frmReporteStockVentasRotacion.frx":1600
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   3
               Left            =   1305
               TabIndex        =   14
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
               Container       =   "frmReporteStockVentasRotacion.frx":161C
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   3
               Left            =   2280
               TabIndex        =   15
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
               Container       =   "frmReporteStockVentasRotacion.frx":1638
               Vacio           =   -1  'True
            End
            Begin CATControls.CATTextBox txtCod_Nivel 
               Height          =   315
               Index           =   4
               Left            =   1305
               TabIndex        =   16
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
               Container       =   "frmReporteStockVentasRotacion.frx":1654
               Estilo          =   1
               Vacio           =   -1  'True
               EnterTab        =   -1  'True
            End
            Begin CATControls.CATTextBox txtGls_Nivel 
               Height          =   315
               Index           =   4
               Left            =   2280
               TabIndex        =   17
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
               Container       =   "frmReporteStockVentasRotacion.frx":1670
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
               TabIndex        =   22
               Top             =   45
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
               TabIndex        =   21
               Top             =   405
               Width           =   345
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   135
               TabIndex        =   20
               Top             =   765
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
               TabIndex        =   19
               Top             =   1125
               Width           =   345
            End
            Begin VB.Label lblNivel 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   135
               TabIndex        =   18
               Top             =   1485
               Width           =   345
            End
         End
      End
   End
End
Attribute VB_Name = "frmReporteStockVentasRotacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i               As Integer
Dim rsj             As New ADODB.Recordset

Private Sub Btnaceptar_Click()
On Error GoTo Err
Dim fIni                    As String
Dim Ffin                    As String
Dim StrMsgError             As String
Dim strMoneda               As String
Dim cNiveles                As String
Dim X                       As Integer
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
    
    For X = 1 To glsNumNiveles
        cNiveles = cNiveles & "vn.idNivel" & Format(X, "00") & ", vn.GlsNivel" & Format(X, "00") & ","
    Next X

    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strMoneda = "PEN"
     
    mostrarReporte "rptStockVentasRatios" & Format(glsNumNiveles, "00") & ".rpt", "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta|parCondNiveles|parNiveles|parProducto", glsEmpresa & "|" & glsSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|" & cWhereNiveles & "|" & cNiveles & "|" & txtCod_Producto.Text, GlsForm, StrMsgError
    If StrMsgError <> "" Then GoTo Err
       
      
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
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
    
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto, " and idnivel = '" & txtCod_Nivel(Val(glsNumNiveles) - 1).Text & "' "

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    mostrarNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
 
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")

    txtGls_Nivel(0).Text = "TODO LOS GRUPOS"
    txtGls_Nivel(1).Text = "TODA LAS CATEGORIAS"
    txtGls_Nivel(2).Text = "TODA LAS SUB CATEGORIAS"
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
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
            txtGls_Nivel(0).Text = "TODO LOS GRUPOS"
            
            txtCod_Nivel(1).Text = ""
            txtGls_Nivel(1).Text = "TODA LAS CATEGORIAS"
            
            txtCod_Nivel(2).Text = ""
            txtGls_Nivel(2).Text = "TODA LAS SUB CATEGORIAS"
            
            txtCod_Producto.Text = ""
            txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
            Exit Sub
        End If
    End If
    
End Sub
 
Private Sub btnCancelar_Click()
    
    Unload Me

End Sub

Private Sub txtCod_Producto_Change()

    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", "idProducto", txtCod_Producto.Text, True)
    Else
        txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    End If

End Sub
