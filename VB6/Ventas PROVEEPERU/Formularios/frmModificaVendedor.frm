VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmModificaVendedor 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar Vendedor"
   ClientHeight    =   2235
   ClientLeft      =   765
   ClientTop       =   3345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPrincipal 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   2220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8865
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   390
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1620
         Width           =   1365
      End
      Begin VB.CommandButton BtnSalir 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1620
         Width           =   1320
      End
      Begin VB.CommandButton cmbAyudaCliente 
         Height          =   315
         Left            =   8055
         Picture         =   "frmModificaVendedor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1170
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaVendedorNuevo 
         Height          =   315
         Left            =   8055
         Picture         =   "frmModificaVendedor.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaVendedorActual 
         Height          =   315
         Left            =   8055
         Picture         =   "frmModificaVendedor.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Vendedor_Actual 
         Height          =   285
         Left            =   1290
         TabIndex        =   3
         Tag             =   "TidPerCliente"
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "frmModificaVendedor.frx":0A9E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor_Actual 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Tag             =   "TGlsCliente"
         Top             =   225
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   503
         BackColor       =   12648447
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Locked          =   -1  'True
         Container       =   "frmModificaVendedor.frx":0ABA
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Vendedor_Nuevo 
         Height          =   285
         Left            =   1305
         TabIndex        =   8
         Tag             =   "TidPerCliente"
         Top             =   765
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "frmModificaVendedor.frx":0AD6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor_Nuevo 
         Height          =   285
         Left            =   2295
         TabIndex        =   9
         Tag             =   "TGlsCliente"
         Top             =   765
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   503
         BackColor       =   12648447
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Locked          =   -1  'True
         Container       =   "frmModificaVendedor.frx":0AF2
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   285
         Left            =   1305
         TabIndex        =   11
         Tag             =   "TidPerCliente"
         Top             =   1215
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   8
         Container       =   "frmModificaVendedor.frx":0B0E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   285
         Left            =   2295
         TabIndex        =   12
         Tag             =   "TGlsCliente"
         Top             =   1215
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   503
         BackColor       =   12648447
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Locked          =   -1  'True
         Container       =   "frmModificaVendedor.frx":0B2A
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Vendedor Nuevo:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   45
         TabIndex        =   6
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Vendedor Actual:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   45
         TabIndex        =   5
         Top             =   285
         Width           =   1350
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         Caption         =   "Cliente:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   1260
         Width           =   555
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   480
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":0B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":0EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":1332
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":16CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":1A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":1E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":219A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":2534
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":28CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":2C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":3002
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificaVendedor.frx":3CC4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmModificaVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strVendedor As String

Private Sub cmbAyudaCliente_Click()
    strVendedor = Trim(txtCod_Vendedor_Actual.Text)
    mostrarAyuda "CLIENTES2", txtCod_Cliente, txtGls_Cliente, " And c.idVendedorCampo ='" & strVendedor & "'"
    
End Sub

Private Sub cmbAyudaVendedorActual_Click()
    mostrarAyuda "VENDEDOR", txtCod_Vendedor_Actual, txtGls_Vendedor_Actual
End Sub

Private Sub cmbAyudaVendedorNuevo_Click()
    mostrarAyuda "VENDEDOR", txtCod_Vendedor_Nuevo, txtGls_Vendedor_Nuevo
End Sub

Private Sub CmdAceptar_Click()
Dim strMsgError As String
On Error GoTo ERR

    grabar strMsgError
    If strMsgError <> "" Then GoTo ERR
    
    Exit Sub
ERR:
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim strMsgError As String
On Error GoTo ERR
 
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    
    Exit Sub
ERR:
    If strMsgError = "" Then strMsgError = ERR.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Public Sub mostrarForm(ByVal strVarCodVendedor As String)
  Load Me
   strVarCodVendedor = txtCod_Vendedor_Actual.Text
 Unload Me
End Sub

Private Sub grabar(ByRef strMsgError As String)
Dim strMsg                      As String
Dim rst                         As New ADODB.Recordset
Dim strCliente                  As String
Dim strVendedor                 As String
Dim StrVendedorCTD              As String
Dim item                        As Integer
Dim indTrans                    As Boolean

On Error GoTo ERR
       
item = traerCampo("clientes_vendedor", "max(item)", " Idempresa ", glsEmpresa, True)
 
validaFormSQL Me, strMsgError
If strMsgError <> "" Then GoTo ERR

Cn.BeginTrans
indTrans = True

If Trim(txtCod_Cliente.Text = "") Then
  
            csql = "SELECT p.idPersona   ,p.GlsPersona,idvendedorcampo   FROM personas p INNER JOIN clientes c ON c.idEmpresa = '" & glsEmpresa & "' AND p.idPersona = c.idCliente WHERE c.idVendedorCampo ='" & Trim(txtCod_Vendedor_Actual.Text) & "'"
            If rst.State = adStateOpen Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
               If Not rst.EOF Then
                    rst.MoveFirst
                        Do While Not rst.EOF
                             item = item + 1
                             strCliente = "" & rst.Fields("idPersona")
                             strVendedor = "" & rst.Fields("idVendedorCampo")
                                csql = "Insert Into clientes_vendedor(idEmpresa, idCliente, idvendedor_anterior, idvendedor_actual, fechaMod, idusuarioMod,item) " & _
                                "Values('" & glsEmpresa & "','" & strCliente & "','" & strVendedor & "','" & Trim(txtCod_Vendedor_Nuevo.Text) & "','" & Format(getFechaHoraSistema, "yyyy-mm-dd") & "', '" & glsUser & "','" & item & "')"
                                Cn.Execute (csql)
                                
                                csql = "Update Clientes Set idVendedorCampo='" & Trim(txtCod_Vendedor_Nuevo.Text) & "' Where idVendedorCampo='" & strVendedor & "'And Idempresa='" & glsEmpresa & "' "
                                Cn.Execute (csql)
                             rst.MoveNext
                        Loop
                End If
   
              
              csql = "SELECT valSaldo,idCliente,idVendedor FROM Cta_dcto WHERE valSaldo > 0 AND idVendedor='" & Trim(txtCod_Vendedor_Actual.Text) & "' AND idEmpresa='" & glsEmpresa & "' "
              If rst.State = adStateOpen Then rst.Close
              rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
              If Not rst.EOF Then
                rst.MoveFirst
                    Do While Not rst.EOF
                         StrVendedorCTD = "" & rst.Fields("idVendedor")
                            csql = "Update Cta_dcto Set idVendedor='" & Trim(txtCod_Vendedor_Nuevo.Text) & "' Where idVendedor='" & Trim(txtCod_Vendedor_Actual.Text) & "' And idEmpresa='" & glsEmpresa & "' And valSaldo > 0"
                            Cn.Execute (csql)
                         rst.MoveNext
                    Loop
              End If
        Else
            csql = "SELECT p.idPersona  ,p.GlsPersona,idvendedorcampo   FROM personas p INNER JOIN clientes c ON c.idEmpresa = '" & glsEmpresa & "' AND p.idPersona = c.idCliente WHERE c.idVendedorCampo ='" & Trim(txtCod_Vendedor_Actual.Text) & "' And idCliente='" & Trim(txtCod_Cliente.Text) & "'"
            If rst.State = adStateOpen Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
                If Not rst.EOF Then
                    rst.MoveFirst
                             item = item + 1
                             csql = "Insert Into clientes_vendedor(idEmpresa, idCliente, idvendedor_anterior, idvendedor_actual, fechaMod, idusuarioMod,item) " & _
                             "Values('" & glsEmpresa & "','" & rst.Fields("idPersona") & "','" & rst.Fields("idVendedorCampo") & "','" & Trim(txtCod_Vendedor_Nuevo.Text) & "','" & Format(getFechaHoraSistema, "yyyy-mm-dd") & "','" & glsUser & "','" & item & "')"
                             Cn.Execute (csql)
                             
                             csql = "Update Clientes Set idVendedorCampo='" & Trim(txtCod_Vendedor_Nuevo.Text) & "' Where idVendedorCampo='" & Trim(txtCod_Vendedor_Actual.Text) & "'And Idempresa='" & glsEmpresa & "' And idCliente='" & Trim(txtCod_Cliente.Text) & "' "
                             Cn.Execute (csql)
                End If
            csql = "SELECT valSaldo,idCliente,idVendedor FROM Cta_dcto WHERE valSaldo > 0 AND idVendedor='" & Trim(txtCod_Vendedor_Actual.Text) & "' AND idEmpresa='" & glsEmpresa & "' And idCliente='" & Trim(txtCod_Cliente.Text) & "' "
            If rst.State = adStateOpen Then rst.Close
            rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
                If Not rst.EOF Then
                  rst.MoveFirst
                        Do While Not rst.EOF
                              StrVendedorCTD = "" & rst.Fields("idVendedor")
                              csql = "Update Cta_dcto Set idVendedor='" & Trim(txtCod_Vendedor_Nuevo.Text) & "' Where idVendedor='" & Trim(txtCod_Vendedor_Actual.Text) & "' And idEmpresa='" & glsEmpresa & "'  And idCliente='" & Trim(txtCod_Cliente.Text) & "' And valSaldo > 0 "
                              Cn.Execute (csql)
                            rst.MoveNext
                        Loop
                End If
    End If
Cn.CommitTrans

    rst.Close: Set rst = Nothing
    strMsg = "Grabo"
    limpia
Exit Sub
ERR:
    If strMsgError = "" Then strMsgError = ERR.Description
    If indTrans Then Cn.RollbackTrans
    
End Sub
Private Sub limpia()
    txtCod_Vendedor_Actual.Text = ""
    txtGls_Vendedor_Actual.Text = ""
    txtCod_Vendedor_Nuevo.Text = ""
    txtGls_Vendedor_Nuevo.Text = ""
    txtCod_Cliente.Text = ""
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
End Sub


 
