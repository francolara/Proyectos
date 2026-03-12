VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmAprobacionDoc 
   Caption         =   "Acceso de Aprobacion de Descuento"
   ClientHeight    =   2625
   ClientLeft      =   6090
   ClientTop       =   2535
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   4425
   Begin VB.CommandButton cmdopera 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   990
      TabIndex        =   8
      Top             =   2115
      Width           =   1095
   End
   Begin VB.CommandButton cmdopera 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   2340
      TabIndex        =   7
      Top             =   2115
      Width           =   1095
   End
   Begin VB.PictureBox SSFrame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   225
      ScaleHeight     =   1830
      ScaleWidth      =   3990
      TabIndex        =   0
      Top             =   90
      Width           =   4020
      Begin CATControls.CATTextBox txtusuario 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
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
         Container       =   "frmAprobacionDoc.frx":0000
      End
      Begin CATControls.CATTextBox txtpassword 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1260
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
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
         PasswordChar    =   "*"
         Container       =   "frmAprobacionDoc.frx":001C
      End
      Begin VB.Label lblvalor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   855
         Width           =   1995
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   225
         TabIndex        =   5
         Top             =   855
         Width           =   390
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Password"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   225
         TabIndex        =   4
         Top             =   1350
         Width           =   750
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   225
         TabIndex        =   3
         Top             =   405
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmAprobacionDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private indRespuesta As Integer
Private strCodUsuario As String
Private strAutogenerado As String
Private strCodPermiso As String
Private intento As Byte

'indEvaluacion = 0 ERROR        indEvaluacion = 1 OK
Public Sub MostrarForm(ByVal strVarCodPermiso As String, ByRef indEvaluacion As Integer, ByRef strCodUsuarioAutorizacion As String, ByRef StrMsgError As String)

intento = 0

If traerCampo("permisos", "estPermiso", "idPermiso", strVarCodPermiso, False, " CodSistema = '" & StrcodSistema & "' ") = "INA" Then
    indEvaluacion = 1 'Se asume q la validacion es OK ya q no se validara pq el permiso esta inactivo
    strCodUsuarioAutorizacion = ""
    Unload Me
    Exit Sub
End If

strCodPermiso = strVarCodPermiso

Load Me

'txtusuario.Text = glsUser

Me.Show 1

indEvaluacion = indRespuesta
strCodUsuarioAutorizacion = strCodUsuario

End Sub

Private Sub cmdopera_Click(Index As Integer)
Dim rstAprob As New ADODB.Recordset
Dim StrMsgError As String

On Error GoTo Err

    indRespuesta = 0

    Select Case Index
        Case 0
            intento = intento + 1
            
            'FALTA PONER Q ESTE ASIGNADO A LA SUCURSAL
            csql = "select idUsuario,autogenerado from usuarios where idEmpresa = '" & glsEmpresa & "' AND varUsuario='" & txtusuario.Text & "'"
            rstAprob.Open csql, Cn, adOpenStatic, adLockOptimistic
            If rstAprob.EOF Then
                If intento = 3 Then
                    MsgBox "Demasiados numeros de intentos", vbInformation, App.Title
                    Unload Me
                    Exit Sub
                End If
                
                StrMsgError = "Usuario Incorrecto"
                txtusuario.SetFocus
                GoTo Err
            End If
            
            strCodUsuario = Trim("" & rstAprob.Fields("idUsuario"))
            strAutogenerado = Trim("" & rstAprob.Fields("autogenerado"))
            
            rstAprob.Close
            
            If Len(strAutogenerado) = 0 Then
                StrMsgError = "Ud. no está Autorizado para realizar esta Operación"
                GoTo Err
            End If
            
            'Validamos si tiene el permiso
            If traerCampo("permisosusuarios", "idPermiso", "idUsuario", strCodUsuario, True, " idPermiso = '" & strCodPermiso & "' and CodSistema = '" & StrcodSistema & "' ") = "" Then
                StrMsgError = "Ud. no está Autorizado para realizar esta Operación, no tiene el permiso suficiente"
                GoTo Err
            End If
            
            'Verifica Password
            
            contraseña = numero()
            If Val(txtpassword.Value) <> Val(contraseña) Then
                If intento = 3 Then
                    MsgBox "Demasiados numeros de intentos", vbInformation, App.Title
                    Unload Me
                    Exit Sub
                End If
                
                StrMsgError = "Password Incorrecto"
                lblvalor.Caption = GeneraValor()
                txtpassword.Text = ""
                txtpassword.SetFocus
                GoTo Err
            End If
                      
            
            indRespuesta = 1 'OK
            Unload Me
            
        Case 1
            indRespuesta = 0
            Unload Me
    End Select
    
    If rstAprob.State = 1 Then rstAprob.Close
    Set rstAprob = Nothing
    Exit Sub
Err:
    If rstAprob.State = 1 Then rstAprob.Close
    Set rstAprob = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()

    indRespuesta = 0
    lblvalor.Caption = GeneraValor()
    
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdopera(0).Enabled Then
        cmdopera_Click 0
    End If
End If
End Sub

Private Sub txtusuario_GotFocus()
    
    txtusuario.SelStart = 0: txtusuario.SelLength = Len(txtusuario.Text)

End Sub

Private Sub txtpassword_Change()
    
    If Trim$(txtpassword.Text) = Empty Then
        cmdopera(0).Enabled = False
    Else
        cmdopera(0).Enabled = True
    End If

End Sub

Private Sub txtpassword_GotFocus()
    
    txtpassword.SelStart = 0: txtpassword.SelLength = Len(txtpassword.Text)

End Sub

Function numero()
    
    PRODUCTOX = ""
    Pass = ""
    For i = 1 To Len(strAutogenerado)
      Pass = Pass + Chr(Asc(Mid(strAutogenerado, i, 1)) - 5)
    Next
    Password = Pass
    X = 1
    Do While Len(Password) > 0
        valor = Val(Mid(Password, 1, 1))
        Select Case valor
            Case 1 To 6
                PRODUCTOX = PRODUCTOX & Val(Mid(lblvalor.Caption, valor, 1))
            Case 0
                If X = 1 Then
                    res = PRODUCTOX
                    
                    signo = Mid(Password, 1, 1)
                    X = X + 1
                Else
                    res = operacion(Val(res), signo, Val(PRODUCTOX))
                    signo = Mid(Password, 1, 1)
                End If
                PRODUCTOX = ""
        End Select
        Password = Mid(Password, 2)
    Loop
    res = operacion(Val(res), signo, Val(PRODUCTOX))
    numero = res
    
End Function

Function operacion(num1, signo, num2)
    
    Select Case signo
        Case "+"
            operacion = num1 + num2
        Case "-"
            operacion = num1 - num2
        Case "*"
            operacion = num1 * num2
    End Select
    
End Function

Public Function GeneraValor()
Dim valor As Long

    'Devuelve una cadena de 6 números aleatorios con rango de 1 a 9
    Randomize
    valor = Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1)
    GeneraValor = valor

End Function

