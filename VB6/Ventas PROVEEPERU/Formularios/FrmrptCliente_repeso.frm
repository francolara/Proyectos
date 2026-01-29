VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmrptCliente_repeso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes - Repeso por Cliente"
   ClientHeight    =   3495
   ClientLeft      =   4935
   ClientTop       =   2865
   ClientWidth     =   6765
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
   ScaleHeight     =   3495
   ScaleWidth      =   6765
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2955
      Width           =   1200
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2955
      Width           =   1200
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
      Height          =   2805
      Left            =   45
      TabIndex        =   7
      Top             =   45
      Width           =   6675
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   765
         Index           =   11
         Left            =   180
         TabIndex        =   12
         Top             =   990
         Width           =   6375
         Begin VB.CommandButton cmbAyudaCliente 
            Height          =   315
            Left            =   5805
            Picture         =   "FrmrptCliente_repeso.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   315
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Cliente 
            Height          =   315
            Left            =   900
            TabIndex        =   2
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
            Container       =   "FrmrptCliente_repeso.frx":038A
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Cliente 
            Height          =   315
            Left            =   1860
            TabIndex        =   14
            Top             =   315
            Width           =   3915
            _ExtentX        =   6906
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
            Container       =   "FrmrptCliente_repeso.frx":03A6
            Vacio           =   -1  'True
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   1845
         Width           =   6375
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado"
            Height          =   240
            Left            =   3690
            TabIndex        =   4
            Top             =   360
            Width           =   1725
         End
         Begin VB.OptionButton OptResumido 
            Caption         =   "Resumido"
            Height          =   240
            Left            =   1305
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1860
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   135
         Width           =   6375
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
            Format          =   107085825
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
            Format          =   107085825
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   375
            Width           =   420
         End
      End
   End
End
Attribute VB_Name = "FrmrptCliente_repeso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim fIni As String
Dim Ffin As String
 
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    If OptResumido = True Then
        mostrarReporte "RptRepesoporClienteResumido.Rpt", "ParEmpresa|ParFecInicio|ParFecFinal|ParCliente", glsEmpresa & "|" & fIni & "|" & Ffin & "|" & txtCod_Cliente.Text, "Repeso por Cliente - Resumido", StrMsgError
    Else
        mostrarReporte "RptRepesoporClienteDetallado.Rpt", "ParEmpresa|ParFecInicio|ParFecFinal|ParCliente", glsEmpresa & "|" & fIni & "|" & Ffin & "|" & txtCod_Cliente.Text, "Repeso por Cliente - Detallado", StrMsgError
    End If
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
    
    OptResumido.Value = True
    dtpfInicio.Value = Format(getFechaSistema, "DD/MM/YYYY")
    dtpFFinal.Value = Format(getFechaSistema, "DD/MM/YYYY")
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    
End Sub

Private Sub txtCod_Cliente_Change()
    
    If txtCod_Cliente.Text <> "" Then
        txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
    Else
        txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    End If
     
End Sub
