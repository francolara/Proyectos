VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmRptSituacionCotizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situación de la Cotización"
   ClientHeight    =   4800
   ClientLeft      =   3390
   ClientTop       =   2085
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7065
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4020
      Left            =   90
      TabIndex        =   7
      Top             =   0
      Width           =   6855
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Formato "
         ForeColor       =   &H00000000&
         Height          =   800
         Index           =   2
         Left            =   180
         TabIndex        =   21
         Top             =   3060
         Width           =   6420
         Begin VB.OptionButton OptFormato 
            Caption         =   "Detallado"
            Height          =   210
            Index           =   1
            Left            =   3915
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton OptFormato 
            Caption         =   "Resumen"
            Height          =   210
            Index           =   0
            Left            =   1530
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton CmbAyudaMoneda 
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
         Left            =   6180
         Picture         =   "FrmRptSituacionCotizacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   810
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
         Left            =   6180
         Picture         =   "FrmRptSituacionCotizacion.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   330
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   800
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   1305
         Width           =   6420
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1605
            TabIndex        =   1
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   133955585
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4605
            TabIndex        =   2
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   133955585
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4005
            TabIndex        =   13
            Top             =   375
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   990
            TabIndex        =   12
            Top             =   375
            Width           =   465
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Documento - Cotización "
         ForeColor       =   &H00000000&
         Height          =   800
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   2205
         Width           =   6420
         Begin VB.TextBox txt_numdoc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4545
            TabIndex        =   4
            Top             =   270
            Width           =   1365
         End
         Begin VB.TextBox txt_serie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1530
            TabIndex        =   3
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Serie"
            Height          =   210
            Left            =   855
            TabIndex        =   10
            Top             =   315
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   210
            Left            =   3645
            TabIndex        =   9
            Top             =   315
            Width           =   555
         End
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1095
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   330
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
         Container       =   "FrmRptSituacionCotizacion.frx":0714
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2025
         TabIndex        =   15
         Top             =   330
         Width           =   4130
         _ExtentX        =   7276
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
         Container       =   "FrmRptSituacionCotizacion.frx":0730
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox TxtCod_Moneda 
         Height          =   315
         Left            =   1095
         TabIndex        =   18
         Tag             =   "TidMoneda"
         Top             =   810
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
         Container       =   "FrmRptSituacionCotizacion.frx":074C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGls_Moneda 
         Height          =   315
         Left            =   2025
         TabIndex        =   19
         Top             =   810
         Width           =   4125
         _ExtentX        =   7276
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
         Container       =   "FrmRptSituacionCotizacion.frx":0768
         Vacio           =   -1  'True
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   20
         Top             =   855
         Width           =   570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   16
         Top             =   375
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4185
      Width           =   1230
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4185
      Width           =   1230
   End
End
Attribute VB_Name = "FrmRptSituacionCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    
End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError                 As String
Dim FecDesde                    As String
Dim FecHasta                    As String
Dim CGlsReporte                 As String

    FecDesde = Format(dtpfInicio.Value, "yyyy-mm-dd")
    FecHasta = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    If OptFormato(0).Value Then
        
        CGlsReporte = "rptSituacionCotizacionResumen.rpt"
    
    Else
        
        CGlsReporte = "rptSituacionCotizacion.rpt"
        
    End If
    
    mostrarReporte CGlsReporte, "parEmpresa|parCliente|parSerie|parDocVentas|parFechaIni|parFechaFin|parMoneda", glsEmpresa & "|" & txtCod_Cliente.Text & "|" & txt_serie.Text & "|" & txt_numdoc.Text & "|" & FecDesde & "|" & FecHasta & "|" & txtCod_Moneda.Text, Me.Caption, StrMsgError
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
    
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    txtCod_Moneda.Text = "PEN"
 
End Sub

Private Sub txt_numdoc_LostFocus()

    txt_numdoc.Text = Format("" & txt_numdoc.Text, "00000000")
    
End Sub

Private Sub txtCod_Cliente_Change()

    If txtCod_Cliente.Text <> "" Then
        txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
    Else
        txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    End If
    
End Sub

Private Sub txt_Serie_LostFocus()

    txt_serie.Text = Format("" & txt_serie.Text, "000")
    
End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("Monedas", "GlsMoneda", "IdMoneda", txtCod_Moneda.Text, False)
    
End Sub
