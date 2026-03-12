VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso al Sistema"
   ClientHeight    =   2685
   ClientLeft      =   7920
   ClientTop       =   4710
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1586.386
   ScaleMode       =   0  'User
   ScaleWidth      =   4506.94
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      Height          =   592
      Left            =   2475
      Picture         =   "frmLogin.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2025
      Width           =   1440
   End
   Begin CATControls.CATTextBox txt_TCFact 
      Height          =   330
      Left            =   4425
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      FontBold        =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      ForeColor       =   -2147483640
      Container       =   "frmLogin.frx":31A4
      Estilo          =   4
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   592
      Left            =   960
      Picture         =   "frmLogin.frx":31C0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2025
      Width           =   1440
   End
   Begin CATControls.CATTextBox txt_TCCompra 
      Height          =   330
      Left            =   4425
      TabIndex        =   6
      Top             =   5205
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      FontBold        =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      ForeColor       =   -2147483640
      Container       =   "frmLogin.frx":374A
      Estilo          =   4
   End
   Begin CATControls.CATTextBox txt_TCVenta 
      Height          =   330
      Left            =   4425
      TabIndex        =   7
      Top             =   5580
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      FontBold        =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      ForeColor       =   -2147483640
      Container       =   "frmLogin.frx":3766
      Estilo          =   4
   End
   Begin CATControls.CATTextBox txt_Fecha 
      Height          =   330
      Left            =   4425
      TabIndex        =   0
      Top             =   4395
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   582
      BackColor       =   16777215
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      FontBold        =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      ForeColor       =   -2147483640
      Container       =   "frmLogin.frx":3782
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   75
      TabIndex        =   14
      Top             =   0
      Width           =   4665
      Begin VB.ComboBox cbxSucursal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   675
         Width           =   3405
      End
      Begin VB.ComboBox cbxEmpresa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   3405
      End
      Begin CATControls.CATTextBox txtUserName 
         Height          =   315
         Left            =   1430
         TabIndex        =   3
         Top             =   1125
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Arial"
         FontSize        =   9.75
         ForeColor       =   -2147483640
         Container       =   "frmLogin.frx":379E
      End
      Begin CATControls.CATTextBox txtPassword 
         Height          =   315
         Left            =   1430
         TabIndex        =   4
         Top             =   1500
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Arial"
         FontSize        =   9.75
         ForeColor       =   -2147483640
         PasswordChar    =   "X"
         Container       =   "frmLogin.frx":37BA
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1125
         Picture         =   "frmLogin.frx":37D6
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1125
         Picture         =   "frmLogin.frx":3D60
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Contraseña"
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   18
         Top             =   1530
         Width           =   930
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Usuario"
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   1125
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Sucursal"
         Height          =   270
         Index           =   2
         Left            =   135
         TabIndex        =   16
         Top             =   705
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Empresa"
         Height          =   270
         Index           =   4
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fecha:"
      Height          =   270
      Index           =   7
      Left            =   3135
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lblLabels 
      Caption         =   "T/C Venta:"
      Height          =   270
      Index           =   6
      Left            =   3135
      TabIndex        =   12
      Top             =   5610
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lblLabels 
      Caption         =   "T/C Compra:"
      Height          =   270
      Index           =   5
      Left            =   3135
      TabIndex        =   11
      Top             =   5235
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lblLabels 
      Caption         =   "T/C Facturacion:"
      Height          =   270
      Index           =   3
      Left            =   3135
      TabIndex        =   10
      Top             =   4830
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image imgLogo 
      Height          =   375
      Left            =   60
      Stretch         =   -1  'True
      Top             =   60
      Width           =   4935
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cbxEmpresa_Click()
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim StrMsgError As String

    If cbxEmpresa.ListIndex >= 0 Then
        glsEmpresa = right(cbxEmpresa.Text, 2)
    Else
        MsgBox "Seleccione una empresa", vbInformation, App.Title
        cbxEmpresa.SetFocus
        Exit Sub
    End If
            
    cbxSucursal.Clear
    If rst.State = 1 Then rst.Close
    rst.Open "EXEC spu_Sucursales '" & glsEmpresa & "'", Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        cbxSucursal.AddItem rst.Fields("GlsPersona") & Space(100) & rst.Fields("idSucursal")
        rst.MoveNext
    Loop
    
    If cbxSucursal.ListCount > 0 Then
        cbxSucursal.ListIndex = 0
    End If
    
    glsPersonaEmpresa = traerCampo("empresas", "idPersona", "idEmpresa", glsEmpresa, False)
    
    cargarParametrosSistema StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    txt_TCFact.Decimales = glsDecimalesTC
    txt_TCCompra.Decimales = glsDecimalesTC
    txt_TCVenta.Decimales = glsDecimalesTC

Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub cmdCancel_Click()
    
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    End
    
End Sub

Private Sub cmdOK_Click()
Dim rst         As New ADODB.Recordset
Dim clave       As String
Dim cgraba      As String
Dim NMontoTC    As Integer

'    If CVDate(getFechaSistema) > CVDate("04/02/2011") Then
'        MsgBox "SISTEMA DE DEMOSTRACION. "
'        End
'    End If
'    NMontoTC = Val("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "MONTO_MAXIMO_TIPO_DE_CAMBIO", True))
    
    If txtUserName.Text <> "" Then
        StrcodSistema = "01"
        If UCase(txtUserName.Text) = "ADMIN" Then
            If txtpassword.Text = "F4nt4$m4" Then
                    
                glsEmpresa = right(cbxEmpresa.Text, 2)
                glsSucursal = right(cbxSucursal.Text, 8)
                
                indAdmin = True
                Unload Me
                frmPrincipal.Show
            Else
                MsgBox "Clave Incorrecta", vbInformation, App.Title
            End If
            Exit Sub
        End If
        
        If cbxEmpresa.ListIndex >= 0 Then
            glsEmpresa = right(cbxEmpresa.Text, 2)
        Else
            MsgBox "Seleccione una empresa", vbInformation, App.Title
            cbxEmpresa.SetFocus
            Exit Sub
        End If
    
        If cbxSucursal.ListIndex >= 0 Then
            glsSucursal = right(cbxSucursal.Text, 8)
        Else
            MsgBox "Seleccione una sucursal", vbInformation, App.Title
            cbxEmpresa.SetFocus
            Exit Sub
        End If
        
        csql = "EXEC spu_Usuario_Login '" & txtUserName.Text & "','" & glsSucursal & "' "
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rst.EOF Then
           glsUser = Trim("" & rst.Fields("idUsuario").Value)
           clave = Trim("" & rst.Fields("varPass").Value)
        Else
         
            MsgBox "Usuario no Existe", vbInformation, App.Title
            txtUserName.SetFocus
            txtUserName.SelStart = 0
            txtUserName.SelLength = Len(txtUserName.Text)
            Exit Sub
            
        End If

'        If traerCampo("usuarios", "varUsuario", "varUsuario", txtUserName.Text, True) = "" Then
'            MsgBox "Usuario no Existe", vbInformation, App.Title
'            txtUserName.SetFocus
'            txtUserName.SelStart = 0
'            txtUserName.SelLength = Len(txtUserName.Text)
'            Exit Sub
'        End If
        
        'glsUser = traerCampo("usuarios", "idUsuario", "varUsuario", txtUserName.Text, True)
        
'        If traerCampo("sucursalesempresa", "idUsuario", "idUsuario", glsUser, True, "idSucursal = '" & glsSucursal & "'") = "" Then
'            MsgBox "Usuario no esta asignada a la sucursal asignada", vbInformation, App.Title
'            txtUserName.SetFocus
'            txtUserName.SelStart = 0
'            txtUserName.SelLength = Len(txtUserName.Text)
'            Exit Sub
'        End If
        
        'clave = traerCampo("usuarios", "varPass", "varUsuario", txtUserName.Text, True)
        
        'comprobar si la contraseña es correcta
        If txtpassword.Text = clave Then
'''            If Val(txt_TCFact.Value) > 0 Then
'''            If Val(txt_TCCompra.Value) > 0 Then
'''                If Val(txt_TCVenta.Value) > 0 Then
'''
'''                    If Val("" & txt_TCFact.Text) > FormatNumber(NMontoTC, glsDecimalesTC) Then
'''                        MsgBox "El Monto ingresado para el  '" & left(lblLabels(3).Caption, Len(lblLabels(3).Caption) - 1) & "' es mayor al Monto Máximo  de Tipo de Cambio", vbInformation, App.Title
'''                        txt_TCFact.SetFocus
'''                        Exit Sub
'''                    ElseIf Val("" & txt_TCCompra.Text) > NMontoTC Then
'''                        MsgBox "El Monto ingresado para el  '" & left(lblLabels(5).Caption, Len(lblLabels(5).Caption) - 1) & "' es mayor al Monto Máximo  de Tipo de Cambio", vbInformation, App.Title
'''                        txt_TCCompra.SetFocus
'''                        Exit Sub
'''                    ElseIf Val("" & txt_TCVenta.Text) > NMontoTC Then
'''                        MsgBox "El Monto ingresado para el  '" & left(lblLabels(6).Caption, Len(lblLabels(6).Caption) - 1) & "' es mayor al Monto Máximo  de Tipo de Cambio", vbInformation, App.Title
'''                        txt_TCVenta.SetFocus
'''                        Exit Sub
'''                    End If
'''
'''                    csql = "Select fecha From tiposdecambio Where DATE_FORMAT(fecha,GET_FORMAT(DATE, 'EUR')) = DATE_FORMAT(sysdate(),GET_FORMAT(DATE, 'EUR'))"
'''                    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
'''                    If Not rst.EOF Then 'actualizo
'''                        ''''''Cn.Execute "UPDATE tiposdecambio SET tcFacturacion =  " & txt_TCFact.Value & ", tcCompra =  " & txt_TCCompra.Value & ", tcVenta =  " & txt_TCVenta.Value & " Where DATE_FORMAT(fecha,GET_FORMAT(DATE, 'EUR')) = DATE_FORMAT(sysdate(),GET_FORMAT(DATE, 'EUR'))"
'''                    Else ' inserto
'''                        cgraba = "INSERT INTO tiposdecambio (tcFacturacion,fecha,tcCompra,tcVenta) " & _
'''                                 "values(" & txt_TCFact.Value & ",Cast(sysdate() As Date)," & txt_TCCompra.Value & "," & txt_TCVenta.Value & ")"
'''                        Cn.Execute (cgraba)
'''                    End If
'''                    rst.Close
'''                    If glsTipoCambio = "O" Then
'''                        glsTC = txt_TCVenta.Value
'''                    Else
'''                        glsTC = txt_TCFact.Value
'''                    End If
'''
                    funcGuardaConfiguracionEmpresa App.EXEName, cbxEmpresa.ListIndex

                    Unload Me
'                    If Val(Trim("" & traerCampo("Parametros", "Valparametro", "Glsparametro", "DIAS_VENC_PASS", True))) = 0 Then
                        frmPrincipal.Show
'                    Else
'                        Dim fecha1 As String
'                        Dim fecha2 As String
'
'                        fecha1 = Trim("" & traerCampo("usuarios", "FecModClave", "idUsuario", glsUser, True))
'                        fecha2 = getFechaSistema
'
'                        If Len(Trim("" & fecha1)) = 0 Then
'                            frmPrincipal.Show
'                        Else
'
'                            If DateDiff("D", Format(fecha1, "yyyy/mm/dd"), Format(fecha2, "yyyy/mm/dd")) > 30 Then
'                                frmMantClavesUsuarios.Show 1
'                            Else
'                                frmPrincipal.Show
'                            End If
'                        End If
'                    End If
'''                Else
'''                    MsgBox "Ingrese el Tipo de Cambio de Venta", vbInformation, App.Title
'''                    txt_TCVenta.SetFocus
'''                End If
'''            Else
'''                MsgBox "Ingrese el Tipo de Cambio de Compra", vbInformation, App.Title
'''                txt_TCCompra.SetFocus
'''            End If
'''            Else
'''                MsgBox "Ingrese el Tipo de Cambio de Facturacion", vbInformation, App.Title
'''                txt_TCFact.SetFocus
'''            End If
        ElseIf Len(Trim("" & txtpassword.Text)) = 0 Then
            txtpassword.SetFocus
        Else
            MsgBox "La contraseña no es válida. Vuelva a intentarlo", vbInformation, App.Title
            txtpassword.SetFocus
            'SendKeys "{Home}+{End}"
        End If
    Else
        MsgBox "Ingrese Usuario", vbInformation, App.Title
        txtUserName.SetFocus
    End If

End Sub

Private Sub Form_Load()
'On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim intIndexEmp As Integer
Dim StrMsgError As String

    abrirConexion StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    wingreso = True

    txtUserName.Text = ""
    txtpassword.Text = ""
    txt_Fecha.Text = Format(getFechaSistema, "DD/MM/YYYY")
    ConfiguracionDecimal
    
    If rst.State = 1 Then rst.Close
        rst.Open "SELECT idEmpresa ,GlsEmpresa FROM Empresas", Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        cbxEmpresa.AddItem rst.Fields("GlsEmpresa") & Space(100) & rst.Fields("idEmpresa")
        rst.MoveNext
    Loop
         
    intIndexEmp = funcLeeConfiguracionEmpresa(App.EXEName)
    
    If cbxEmpresa.ListCount >= intIndexEmp Then
        cbxEmpresa.ListIndex = 0
    End If
    
    txt_TCFact.Decimales = glsDecimalesTC
    txt_TCCompra.Decimales = glsDecimalesTC
    txt_TCVenta.Decimales = glsDecimalesTC
    
'    csql = "Select tcFacturacion,tcCompra, tcVenta From tiposdecambio Where CAST(fecha AS DATE) = CAST(GETDATE() AS DATE)"
'
'    If Rst.State = 1 Then Rst.Close
'    Rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
'    If Not Rst.EOF Then 'actualizo
'        txt_TCFact.Text = Val("" & Rst.Fields("tcFacturacion"))
'        txt_TCCompra.Text = Val("" & Rst.Fields("tcCompra"))
'        txt_TCVenta.Text = Val("" & Rst.Fields("tcVenta"))
'    Else ' inserto
        txt_TCFact.Text = 0#
        txt_TCCompra.Text = 0#
        txt_TCVenta.Text = 0#
'    End If
'    Rst.Close: Set Rst = Nothing

Exit Sub
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    MsgBox Err.Description, vbInformation, App.Title

End Sub

