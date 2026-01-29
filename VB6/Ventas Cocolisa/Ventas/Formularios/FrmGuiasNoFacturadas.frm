VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmGuiasNoFacturadas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Guías No Facturadas"
   ClientHeight    =   2625
   ClientLeft      =   4755
   ClientTop       =   3525
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Btnaceptar 
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   2340
      TabIndex        =   3
      Top             =   2070
      Width           =   1410
   End
   Begin VB.CommandButton BtnSalir 
      Caption         =   "Salir"
      Height          =   420
      Left            =   3825
      TabIndex        =   4
      Top             =   2070
      Width           =   1410
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
      Height          =   1860
      Left            =   45
      TabIndex        =   5
      Top             =   90
      Width           =   7170
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
         Left            =   6570
         Picture         =   "FrmGuiasNoFacturadas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1305
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Width           =   6780
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
            Format          =   104792065
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
            Format          =   104792065
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   375
            Width           =   420
         End
      End
      Begin CATControls.CATTextBox txtCod_Almacen 
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Tag             =   "TidAlmacen"
         Top             =   1305
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
         Container       =   "FrmGuiasNoFacturadas.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Almacen 
         Height          =   315
         Left            =   1980
         TabIndex        =   10
         Top             =   1305
         Width           =   4575
         _ExtentX        =   8070
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
         Container       =   "FrmGuiasNoFacturadas.frx":03A6
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Almacen 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   11
         Top             =   1335
         Width           =   630
      End
   End
End
Attribute VB_Name = "FrmGuiasNoFacturadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btnaceptar_Click()
On Error GoTo Err
Dim StrMsgError             As String
Dim fIni                    As String
Dim Ffin                    As String
Dim GlsForm                 As String
Dim Almacen                 As String

    GlsForm = "Guias No Facturadas"
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    Almacen = Trim(txtCod_Almacen.Text)
    
    mostrarReporte "RptListaGuiasNoFacturadas.rpt", "parEmpresa|parAlmacen|parFecIni|parFecFin", glsEmpresa & "|" & Almacen & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
              
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub btnSalir_Click()

    Unload Me
    
End Sub

Private Sub cmbAyudaAlmacen_Click()
    
    mostrarAyuda "ALMACEN", txtCod_Almacen, txtGls_Almacen
    If txtCod_Almacen.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub Form_Load()
    
    Me.top = 0
    Me.left = 0
    txtGls_Almacen.Text = "TODOS LOS ALMACENES"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    
End Sub

Private Sub txtCod_Almacen_Change()

    If txtCod_Almacen.Text <> "" Then
        txtGls_Almacen.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Almacen.Text, False)
    Else
        txtGls_Almacen.Text = "TODOS LOS ALMACENES"
    End If

End Sub

Private Sub txtCod_Almacen_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Almacen.Text = ""
    End If

End Sub
