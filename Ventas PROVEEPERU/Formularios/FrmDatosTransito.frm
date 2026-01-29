VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmDatosTransito 
   Caption         =   "Datos en Tránsito"
   ClientHeight    =   3675
   ClientLeft      =   8310
   ClientTop       =   3960
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   8385
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8295
      Begin VB.CommandButton BtnSalir 
         Caption         =   "Cancelar"
         Height          =   465
         Left            =   4170
         TabIndex        =   10
         Top             =   3000
         Width           =   1245
      End
      Begin VB.CommandButton BtnAceptar 
         Caption         =   "Aceptar"
         Height          =   465
         Left            =   2790
         TabIndex        =   9
         Top             =   3000
         Width           =   1245
      End
      Begin VB.TextBox txtGls_ObsDatTran 
         Height          =   1785
         Left            =   1380
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "TGlsObservacion"
         Top             =   1020
         Width           =   6105
      End
      Begin VB.CommandButton cmbPerTran 
         Height          =   315
         Left            =   7485
         Picture         =   "FrmDatosTransito.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_PerTran 
         Height          =   315
         Left            =   1425
         TabIndex        =   2
         Tag             =   "TidProvCliente"
         Top             =   240
         Width           =   930
         _ExtentX        =   1640
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
         MaxLength       =   20
         Container       =   "FrmDatosTransito.frx":038A
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_PerTran 
         Height          =   315
         Left            =   2370
         TabIndex        =   3
         Top             =   240
         Width           =   5100
         _ExtentX        =   8996
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
         Container       =   "FrmDatosTransito.frx":03A6
         Vacio           =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtp_FecDatTran 
         Height          =   315
         Left            =   1410
         TabIndex        =   7
         Tag             =   "FfechaEmision"
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   108593153
         CurrentDate     =   38955
      End
      Begin VB.Label lbl_FechaEmision 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   90
         TabIndex        =   8
         Top             =   630
         Width           =   450
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   90
         TabIndex        =   6
         Top             =   1050
         Width           =   930
      End
      Begin VB.Label lblProvClie 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Personal Tránsito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   1260
      End
   End
End
Attribute VB_Name = "FrmDatosTransito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strNumValeAux As String

Private Sub BtnAceptar_Click()
Dim StrMsgError As String
On Error GoTo Err

 If Trim(txtCod_PerTran.Text) = "" Then
    StrMsgError = "Ingrese unPersonal en Tránsito": GoTo Err
 End If

 csql = "Update ValesCab Set idPerTran = '" & txtCod_PerTran.Text & "', FecDatTran = '" & Format(dtp_FecDatTran.Value, "yyyy-mm-dd") & "', ObsDatTran = '" & txtGls_ObsDatTran.Text & "'" & _
        "Where idValesCab = '" & strNumValeAux & "' And TipoVale = 'S' And idEmpresa  = '" & glsEmpresa & "' And idSucursal ='" & glsSucursal & "'"
 Cn.Execute (csql)

 MsgBox "Datos ingresados correctamente", vbInformation, App.Title


  Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub MostrarForm(strNumVale As String, StrMsgError As String)
Dim rst As New ADODB.Recordset
On Error GoTo Err

  strNumValeAux = strNumVale
  
  csql = "Select  IfNull(idPerTran,'') idPerTran, IfNull(FecDatTran,'') FecDatTran, IfNull(ObsDatTran,'') ObsDatTran  " & _
         "From ValesCab " & _
         "Where idValesCab = '" & strNumValeAux & "' And TipoVale = 'S' And idEmpresa  = '" & glsEmpresa & "' And idSucursal ='" & glsSucursal & "'"
  rst.Open csql, Cn, adOpenStatic, adLockReadOnly
  
  If Not rst.EOF Then
    txtCod_PerTran.Text = Trim("" & rst.Fields("idPerTran"))
    dtp_FecDatTran.Value = IIf(Trim("" & rst.Fields("FecDatTran")) = "", getFechaSistema, Trim("" & rst.Fields("FecDatTran")))
    txtGls_ObsDatTran.Text = Trim("" & rst.Fields("ObsDatTran"))
  End If
  
  If rst.State = 1 Then rst.Close: Set rst = Nothing
  
  
  Me.Show

  Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub BtnSalir_Click()

    Unload Me
    
End Sub

Private Sub cmbPerTran_Click()

 mostrarAyuda "PROVEEDOR", txtCod_PerTran, txtGls_PerTran
 
End Sub

Private Sub txtCod_PerTran_Change()
    If txtCod_PerTran.Text = "" Then
        txtGls_PerTran.Text = ""
    Else
        txtGls_PerTran.Text = traerCampo("Personas", "GlsPersona", "idPersona", txtCod_PerTran.Text, False)
    End If
End Sub

