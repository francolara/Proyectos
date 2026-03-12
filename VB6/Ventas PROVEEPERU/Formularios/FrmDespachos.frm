VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmDespachos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Despachos "
   ClientHeight    =   3375
   ClientLeft      =   3795
   ClientTop       =   2805
   ClientWidth     =   7410
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
   ScaleHeight     =   3375
   ScaleWidth      =   7410
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2850
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2850
      Width           =   1230
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
      Height          =   2700
      Index           =   12
      Left            =   90
      TabIndex        =   7
      Top             =   90
      Width           =   7230
      Begin VB.CommandButton cmbAyudaEmpTrans 
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
         Left            =   6645
         Picture         =   "FrmDespachos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   330
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   900
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   1665
         Width           =   6870
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1575
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
            Format          =   107479041
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
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
            Format          =   107479041
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4005
            TabIndex        =   16
            Top             =   420
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   990
            TabIndex        =   15
            Top             =   420
            Width           =   465
         End
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
         Left            =   6645
         Picture         =   "FrmDespachos.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1155
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaChofer 
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
         Left            =   6645
         Picture         =   "FrmDespachos.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   750
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Chofer 
         Height          =   315
         Left            =   1065
         TabIndex        =   1
         Tag             =   "TidPerChofer"
         Top             =   765
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
         Container       =   "FrmDespachos.frx":0A9E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Chofer 
         Height          =   315
         Left            =   2010
         TabIndex        =   9
         Tag             =   "TglsChofer"
         Top             =   765
         Width           =   4590
         _ExtentX        =   8096
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
         Container       =   "FrmDespachos.frx":0ABA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1065
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1170
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
         Container       =   "FrmDespachos.frx":0AD6
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2025
         TabIndex        =   12
         Top             =   1170
         Width           =   4590
         _ExtentX        =   8096
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
         Container       =   "FrmDespachos.frx":0AF2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_EmpTrans 
         Height          =   285
         Left            =   1065
         TabIndex        =   0
         Tag             =   "TidPerEmpTrans"
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
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
         Container       =   "FrmDespachos.frx":0B0E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_EmpTrans 
         Height          =   285
         Left            =   2025
         TabIndex        =   18
         Tag             =   "TGlsEmpTrans"
         Top             =   360
         Width           =   4590
         _ExtentX        =   8096
         _ExtentY        =   503
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
         Container       =   "FrmDespachos.frx":0B2A
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_EmpTrans 
         Appearance      =   0  'Flat
         Caption         =   "Emp.Transp."
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   180
         TabIndex        =   19
         Top             =   375
         Width           =   765
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   195
         TabIndex        =   13
         Top             =   1215
         Width           =   645
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Chofer"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   195
         TabIndex        =   10
         Top             =   855
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmDespachos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaChofer_Click()

    mostrarAyuda "CHOFER", txtCod_Chofer, txtGls_Chofer

End Sub

Private Sub cmbAyudaEmpTrans_Click()
    mostrarAyuda "EMPTRANS", txtCod_EmpTrans, txtGls_EmpTrans
End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub Command1_Click()
On Error GoTo Err
Dim StrMsgError     As String
Dim fIni            As String
Dim Ffin            As String
    
    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    mostrarReporte "rptDespachoXChofer.rpt", "parEmpresa|parSucursal|parChofer|parDesde|parHasta|parEmpTrans", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Chofer.Text) & "|" & fIni & "|" & Ffin & "|" & Trim(txtCod_EmpTrans.Text), "Resumen de Despachos ", StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Command2_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"

End Sub

Private Sub txtCod_Chofer_Change()
    If txtCod_Chofer.Text <> "" Then
        txtGls_Chofer.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Chofer.Text, False)
    Else
        txtGls_Chofer.Text = ""
    End If
End Sub

Private Sub txtCod_Chofer_Click()
    
    txtGls_Chofer.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Chofer.Text, False)

End Sub

Private Sub txtCod_EmpTrans_Change()
    If txtCod_EmpTrans.Text <> "" Then
        txtGls_EmpTrans.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_EmpTrans.Text, False)
    Else
        txtGls_EmpTrans.Text = ""
    End If
End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If

End Sub
