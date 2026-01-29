VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmRankingVentasPorLinea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ranking de Ventas Por Línea"
   ClientHeight    =   4875
   ClientLeft      =   5775
   ClientTop       =   3315
   ClientWidth     =   7140
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
   ScaleHeight     =   4875
   ScaleWidth      =   7140
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   3525
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4275
      Width           =   1300
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   2115
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4275
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Height          =   4065
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   6990
      Begin VB.CheckBox ChkEspeciales 
         Caption         =   "Sólo Clientes Especiales"
         Height          =   240
         Left            =   270
         TabIndex        =   21
         Top             =   3600
         Width           =   2895
      End
      Begin VB.CheckBox ChkMuestrasG 
         Caption         =   "Sólo Muestras Gratuitas"
         Height          =   240
         Left            =   270
         TabIndex        =   20
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Frame FraOrdenRes 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         ForeColor       =   &H00000000&
         Height          =   810
         Left            =   225
         TabIndex        =   17
         Top             =   2250
         Width           =   6510
         Begin VB.OptionButton OptOrden 
            Caption         =   "Cliente"
            Height          =   240
            Index           =   0
            Left            =   1170
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton OptOrden 
            Caption         =   "Precio Venta"
            Height          =   240
            Index           =   1
            Left            =   4275
            TabIndex        =   18
            Top             =   360
            Width           =   2025
         End
      End
      Begin VB.CheckBox ChkOficial 
         Caption         =   "Documentos"
         Height          =   240
         Left            =   5400
         TabIndex        =   16
         Top             =   3285
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1230
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
         Left            =   6385
         Picture         =   "FrmRankingVentasPorLinea.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1780
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
         Left            =   6385
         Picture         =   "FrmRankingVentasPorLinea.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1410
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   1
         Left            =   225
         TabIndex        =   7
         Top             =   405
         Width           =   6510
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   0
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   107610113
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
            Format          =   107610113
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   9
            Top             =   375
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3960
            TabIndex        =   8
            Top             =   375
            Width           =   420
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1245
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1410
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
         Container       =   "FrmRankingVentasPorLinea.frx":0714
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2190
         TabIndex        =   11
         Top             =   1440
         Width           =   4185
         _ExtentX        =   7382
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
         Container       =   "FrmRankingVentasPorLinea.frx":0730
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1245
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   1785
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
         Container       =   "FrmRankingVentasPorLinea.frx":074C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2190
         TabIndex        =   14
         Top             =   1785
         Width           =   4185
         _ExtentX        =   7382
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
         Container       =   "FrmRankingVentasPorLinea.frx":0768
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   250
         TabIndex        =   15
         Top             =   1860
         Width           =   570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   250
         TabIndex        =   12
         Top             =   1455
         Width           =   645
      End
   End
End
Attribute VB_Name = "FrmRankingVentasPorLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError         As String
Dim COrden              As String
Dim CIndMuestras        As String
Dim CIndEspecial        As String

    If OptOrden(0).Value Then
        
        COrden = "D"
    
    Else
        
        COrden = "V"
    
    End If
    
    CIndMuestras = ""
    
    If ChkMuestrasG.Value = "1" Then
        
        CIndMuestras = "1"
        
    End If
    
    CIndEspecial = ""
    
    If ChkEspeciales.Value = "1" Then
        
        CIndEspecial = "1"
        
    End If
    
    FrmRankingVentasPorLineaConsulta.MostrarDatos Format(dtpfInicio.Value, "yyyy-mm-dd"), Format(dtpFFinal.Value, "yyyy-mm-dd"), txtCod_Moneda.Text, txtCod_Sucursal.Text, ChkOficial.Value, StrMsgError, COrden, CIndMuestras, CIndEspecial
    If StrMsgError <> "" Then GoTo Err
                 
    Exit Sub
    
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
    
    Me.top = 0
    Me.left = 0
    
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtCod_Moneda.Text = "PEN"
    txtGls_Moneda.Text = "NUEVOS SOLES"
    ChkOficial.Visible = IIf(GlsVisualiza_Filtro_Documento = "S", True, False)
    
End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    End If
    
    Me.Caption = Me.Caption & " - " & txtGls_Sucursal.Text
    
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

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub txtCod_Moneda_Change()
    
    If Len(Trim(txtCod_Moneda.Text)) > 0 Then
        txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
    Else
        txtGls_Moneda.Text = "MONEDA ORIGINAL"
    End If
    
End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"

End Sub

