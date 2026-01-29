VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRptVtaGratuitas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas Gratuitas"
   ClientHeight    =   2355
   ClientLeft      =   4815
   ClientTop       =   3195
   ClientWidth     =   6465
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
   ScaleHeight     =   2355
   ScaleWidth      =   6465
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   6270
      Begin VB.Frame fraReportes 
         Caption         =   " Rango de Fechas "
         Height          =   900
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   5925
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
            Format          =   84213761
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
            Format          =   84213761
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   375
            Width           =   420
         End
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1200
   End
End
Attribute VB_Name = "FrmRptVtaGratuitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
On Error GoTo ERR
Dim StrMsgError As String
Dim fIni As String
Dim fFin As String

    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    fFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
        
    mostrarReporte "RptVentasGratuitas.Rpt", "ParEmpresa|ParFecInicio|ParFecFinal", glsEmpresa & "|" & fIni & "|" & fFin, "Reporte de Ventas Gratuitas", StrMsgError
    If StrMsgError <> "" Then GoTo ERR
    
    Exit Sub
    
ERR:
    If StrMsgError = "" Then StrMsgError = ERR.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    Me.top = 0
    Me.left = 0
    dtpfInicio.Value = Format(getFechaSistema, "DD/MM/YYYY")
    dtpFFinal.Value = Format(getFechaSistema, "DD/MM/YYYY")

End Sub
