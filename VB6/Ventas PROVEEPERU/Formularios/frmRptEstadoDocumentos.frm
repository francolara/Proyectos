VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmRptEstadoDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Estado de Documentos"
   ClientHeight    =   3435
   ClientLeft      =   4680
   ClientTop       =   2700
   ClientWidth     =   7440
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7440
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1185
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
      Height          =   2700
      Index           =   4
      Left            =   90
      TabIndex        =   7
      Top             =   45
      Width           =   7275
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   1
         Left            =   225
         TabIndex        =   17
         Top             =   1575
         Width           =   6825
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1650
            TabIndex        =   3
            Top             =   345
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
            Format          =   107085825
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4455
            TabIndex        =   4
            Top             =   345
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
            Format          =   107085825
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3915
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   375
            Width           =   465
         End
      End
      Begin VB.CommandButton cmbAyudaEstado 
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
         Picture         =   "frmRptEstadoDocumentos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1125
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaDocumento 
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
         Picture         =   "frmRptEstadoDocumentos.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
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
         Left            =   6600
         Picture         =   "frmRptEstadoDocumentos.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1300
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   315
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
         Container       =   "frmRptEstadoDocumentos.frx":0A9E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2250
         TabIndex        =   9
         Top             =   315
         Width           =   4320
         _ExtentX        =   7620
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
         Container       =   "frmRptEstadoDocumentos.frx":0ABA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Documento 
         Height          =   315
         Left            =   1300
         TabIndex        =   1
         Tag             =   "TidMoneda"
         Top             =   735
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
         Container       =   "frmRptEstadoDocumentos.frx":0AD6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Documento 
         Height          =   315
         Left            =   2250
         TabIndex        =   12
         Top             =   735
         Width           =   4320
         _ExtentX        =   7620
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
         Container       =   "frmRptEstadoDocumentos.frx":0AF2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Estado 
         Height          =   315
         Left            =   1300
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1140
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
         Container       =   "frmRptEstadoDocumentos.frx":0B0E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Estado 
         Height          =   315
         Left            =   2250
         TabIndex        =   15
         Top             =   1140
         Width           =   4320
         _ExtentX        =   7620
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
         Container       =   "frmRptEstadoDocumentos.frx":0B2A
         Vacio           =   -1  'True
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   1185
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   810
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmRptEstadoDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaDocumento_Click()

    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento

End Sub

Private Sub cmbAyudaEstado_Click()
    
    mostrarAyuda "ESTDOCUMENTOS", txtCod_Estado, txtGls_Estado

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub Command1_Click()
On Error GoTo Err
Dim StrMsgError As String

    mostrarReporte "rptEstadoDocumentos.rpt", "parEmpresa|parSucursal|parDocumento|parEstado|parFecIni|parFecFin", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & txtCod_Documento.Text & "|" & txtCod_Estado.Text & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd"), "Estado de Documento", StrMsgError
    If StrMsgError <> "" Then GoTo Err
        
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Command2_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    txtGls_Documento.Text = "TODOS LOS DOCUMENTOS"
    txtGls_Estado.Text = "TODOS LOS ESTADOS"
    dtpfInicio.Value = Format(Now, "DD/MM/YYYY")
    dtpFFinal.Value = Format(Now, "DD/MM/YYYY")

End Sub

Private Sub txtCod_Documento_Change()

    If txtCod_Documento.Text = "" Then
        txtGls_Documento.Text = "TODOS LOS DOCUMENTOS"
    Else
        txtGls_Documento.Text = traerCampo("documentos", "GlsDocumento", "idDocumento", txtCod_Documento.Text, False)
    End If
    
End Sub

Private Sub txtCod_Estado_Change()

    Select Case txtCod_Estado.Text
        Case "": txtGls_Estado.Text = "TODOS LOS ESTADOS"
        Case "ANU": txtGls_Estado.Text = "ANULADO"
        Case "GEN": txtGls_Estado.Text = "GENERADO"
        Case "CAN": txtGls_Estado.Text = "CANCELADO"
        Case "IMP": txtGls_Estado.Text = "IMPRESO"
    End Select
    
End Sub

Private Sub txtCod_Sucursal_Change()

    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    
End Sub
