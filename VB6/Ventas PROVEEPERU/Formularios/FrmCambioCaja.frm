VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmCambioCaja 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Caja"
   ClientHeight    =   5685
   ClientLeft      =   4020
   ClientTop       =   2415
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7695
   Begin VB.CommandButton BtnSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3915
      TabIndex        =   8
      Top             =   5085
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   4920
      Left            =   90
      TabIndex        =   9
      Top             =   45
      Width           =   7485
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   840
         Index           =   9
         Left            =   180
         TabIndex        =   26
         Top             =   2430
         Width           =   7095
         Begin VB.CommandButton cmbAyudaUsuario 
            Height          =   315
            Left            =   6540
            Picture         =   "FrmCambioCaja.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   315
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Usuario 
            Height          =   315
            Left            =   1170
            TabIndex        =   4
            Tag             =   "TidMoneda"
            Top             =   315
            Width           =   1095
            _ExtentX        =   1931
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
            Container       =   "FrmCambioCaja.frx":038A
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Usuario 
            Height          =   315
            Left            =   2310
            TabIndex        =   28
            Top             =   315
            Width           =   4200
            _ExtentX        =   7408
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
            Container       =   "FrmCambioCaja.frx":03A6
            Vacio           =   -1  'True
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
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
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   555
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Documento "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   225
         Width           =   7095
         Begin VB.CommandButton cmbAyudaTipoDoc 
            Height          =   315
            Left            =   6540
            Picture         =   "FrmCambioCaja.frx":03C2
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   360
            Width           =   390
         End
         Begin VB.TextBox txtserie 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1170
            TabIndex        =   1
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtnumdoc 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5265
            TabIndex        =   2
            Top             =   720
            Width           =   1140
         End
         Begin CATControls.CATTextBox txtCod_Documento 
            Height          =   315
            Left            =   1170
            TabIndex        =   0
            Tag             =   "TidMoneda"
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
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
            Container       =   "FrmCambioCaja.frx":074C
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Documento 
            Height          =   315
            Left            =   2310
            TabIndex        =   22
            Top             =   360
            Width           =   4200
            _ExtentX        =   7408
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
            Container       =   "FrmCambioCaja.frx":0768
            Vacio           =   -1  'True
            CambiarConFoco  =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4590
            TabIndex        =   25
            Top             =   810
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Serie"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   270
            TabIndex        =   24
            Top             =   810
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "T/D"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   270
            TabIndex        =   23
            Top             =   405
            Width           =   240
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Caja "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1290
         Index           =   10
         Left            =   180
         TabIndex        =   14
         Top             =   3420
         Width           =   7095
         Begin VB.CommandButton cmbAyudaCaja 
            Height          =   315
            Left            =   6540
            Picture         =   "FrmCambioCaja.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   270
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Caja 
            Height          =   315
            Left            =   1170
            TabIndex        =   5
            Tag             =   "TidMoneda"
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
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
            Container       =   "FrmCambioCaja.frx":0B0E
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Caja 
            Height          =   315
            Left            =   2310
            TabIndex        =   16
            Top             =   270
            Width           =   4200
            _ExtentX        =   7408
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
            Container       =   "FrmCambioCaja.frx":0B2A
            Vacio           =   -1  'True
         End
         Begin MSComCtl2.DTPicker cbofecha 
            Height          =   330
            Left            =   1170
            TabIndex        =   6
            Top             =   765
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
            Format          =   103940097
            CurrentDate     =   39060
         End
         Begin VB.Label Label5 
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
            Height          =   210
            Left            =   225
            TabIndex        =   19
            Top             =   810
            Width           =   450
         End
         Begin VB.Label lblPrueba 
            Caption         =   "---"
            Height          =   135
            Left            =   5640
            TabIndex        =   18
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Caja"
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
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   4
         Left            =   180
         TabIndex        =   10
         Top             =   1530
         Width           =   7095
         Begin VB.CommandButton cmbAyudaSucursal 
            Height          =   315
            Left            =   6540
            Picture         =   "FrmCambioCaja.frx":0B46
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   315
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Sucursal 
            Height          =   315
            Left            =   1170
            TabIndex        =   3
            Tag             =   "TidMoneda"
            Top             =   315
            Width           =   1095
            _ExtentX        =   1931
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
            Container       =   "FrmCambioCaja.frx":0ED0
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Sucursal 
            Height          =   315
            Left            =   2310
            TabIndex        =   12
            Top             =   315
            Width           =   4200
            _ExtentX        =   7408
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
            Container       =   "FrmCambioCaja.frx":0EEC
            Vacio           =   -1  'True
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            Caption         =   "Sucursal"
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
            Height          =   240
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   765
         End
      End
   End
   Begin VB.CommandButton Btnaceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2565
      TabIndex        =   7
      Top             =   5085
      Width           =   1275
   End
End
Attribute VB_Name = "FrmCambioCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idmovcaja  As String

Private Sub Btnaceptar_Click()

    If Len(txtCod_Documento.Text) = 0 Or Len(txtCod_Documento.Text) > 2 Then
        MsgBox "Tipo de Documento Incorrecto,Verifque", vbInformation, App.Title
        txtCod_Documento.SetFocus
        Exit Sub
    End If

    If Len(txtserie.Text) = 0 Or Len(txtserie.Text) > 3 Then
         MsgBox "Numero de Serie Incorrecto,Verifque", vbInformation, App.Title
         txtserie.SetFocus
         Exit Sub
    End If

    If (txtserie.Text) > 0 And Len(txtserie.Text) = 1 Then
         txtserie.Text = "00" & Val(txtserie.Text)
    End If
    
    If Len(txtnumdoc.Text) = 0 Or Len(txtnumdoc.Text) > 8 Then
         MsgBox "Numero de Documento Incorrecto,Verifque", vbInformation, App.Title
         txtnumdoc.SetFocus
        Exit Sub
    End If
    
        If Len(txtCod_Sucursal.Text) = 0 Or Len(txtnumdoc.Text) > 8 Then
         MsgBox "Debe de Ingresar una Sucursal,Verifque", vbInformation, App.Title
         txtnumdoc.SetFocus
        Exit Sub
    End If
    
    If Len(txtCod_Usuario.Text) = 0 Then
         MsgBox "Debe de Ingresar un Usuario,Verifque", vbInformation, App.Title
         txtnumdoc.SetFocus
        Exit Sub
    End If
    
    If Len(txtCod_Caja.Text) = 0 Then
         MsgBox "Debe de Ingresar una Caja,Verifque", vbInformation, App.Title
         txtnumdoc.SetFocus
        Exit Sub
    End If
    validar "A"

End Sub

Private Sub btnSalir_Click()
    
    Unload Me

End Sub

Private Sub cmbAyudaCaja_Click()
    
    mostrarAyuda "CAJASUSUARIOFILTRO", txtCod_Caja, txtGls_Caja, "AND u.idUsuario = '" & txtCod_Usuario.Text & "'"

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmbAyudaTipoDoc_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento

End Sub

Private Sub cmbAyudaUsuario_Click()
    
    mostrarAyuda "USUARIO", txtCod_Usuario, txtGls_Usuario

End Sub

Private Sub Form_Load()
    
    cbofecha.Value = Format(Date, "DD/MM/YYYY")

End Sub

Private Sub txtCod_Documento_Click()
    
    txtGls_Documento.Text = traerCampo("Documentos", "GlsDocumento", "idDocumento", txtCod_Documento.Text, False)

End Sub

Private Sub actualizar_mov()

    csql = "Update MovCajasDet Set idmovcaja = '" & idmovcaja & "' " & _
           "Where iddocventas = '" & txtnumdoc.Text & "' And idserie = '" & txtserie.Text & "' And iddocumento = '" & txtCod_Documento.Text & "' And idEmpresa = '" & glsEmpresa & "' "
    Cn.Execute csql
        
End Sub

Private Sub actualizar_doc_ventas()

    csql = "Update Docventas Set idmovcaja = '" & idmovcaja & "' " & _
            " WHERE iddocventas = '" & txtnumdoc.Text & "' And idserie = '" & txtserie.Text & "' And iddocumento = '" & txtCod_Documento.Text & "'  And idEmpresa = '" & glsEmpresa & "' "
    Cn.Execute csql
        
End Sub

Private Sub cogercodigo_movcaja(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim fecha As String

    fecha = Format(cbofecha.Value, "yyyy-mm-dd")
    csql = "SELECT idmovcaja " & _
           "FROM movcajas " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
           "AND idcaja = '" & txtCod_Caja.Text & "' " & _
           "AND idusuario ='" & txtCod_Usuario.Text & "' " & _
           "AND feccaja = '" & fecha & "' " & _
           "AND idsucursal = '" & txtCod_Sucursal.Text & "'  "

    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        Do While Not rst.EOF
            idmovcaja = "" & rst.Fields("idmovcaja")
            rst.MoveNext
        Loop
    Else
        StrMsgError = "Algunos datos no coinciden. Verifique."
        Exit Sub
    End If
    
Err:

End Sub
 
Private Sub actualizar_caja()
On Error GoTo Err
Dim StrMsgError As String

    If MsgBox("¿Seguro de Actualizar el Estado?", vbInformation + vbYesNo, App.Title) = vbYes Then
        cogercodigo_movcaja StrMsgError
        If StrMsgError <> "" Then GoTo Err
        actualizar_mov
        actualizar_doc_ventas
        MsgBox ("EL Documento se ha modificado de Caja satisfactoriamente.")
    Else
        MsgBox ("Proceso Cancelado")
    End If
    
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
 
 Private Sub limpiar()
 
    txtCod_Documento.Text = ""
    txtGls_Documento.Text = ""
    txtserie.Text = ""
    txtnumdoc.Text = ""
    txtCod_Sucursal.Text = ""
    txtGls_Sucursal.Text = ""
    txtCod_Usuario.Text = ""
    txtGls_Usuario.Text = ""
    txtCod_Caja.Text = ""
    txtGls_Caja.Text = ""
 
End Sub

Private Sub txtCod_Documento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtserie.SetFocus
    End If

End Sub

Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtCod_Usuario.SetFocus
    End If
    
End Sub

Private Sub txtCod_Sucursal_LostFocus()
    
    txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)

End Sub

Private Sub txtnumdoc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtCod_Sucursal.SetFocus
    End If

End Sub

Private Sub txtnumdoc_LostFocus()

    txtnumdoc.Text = Format("" & txtnumdoc.Text, "00000000")

End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtnumdoc.SetFocus
    End If

End Sub

Public Sub validar(StrTipo As String)
On Error GoTo Err
Dim indEvaluacion As Integer
Dim strCodUsuarioAutorizacion As String
Dim StrMsgError As String
    
    If StrTipo = "A" Then
        indEvaluacion = 0
        frmAprobacion.MostrarForm "06", indEvaluacion, strCodUsuarioAutorizacion, StrMsgError
        If intento = 3 Or indRespuesta = 0 Then
            StrMsgError = Err.Description
            StrMsgError = "Intentelo de nuevo."
            Exit Sub
        Else
            actualizar_caja
        End If
    End If
    
Err:
End Sub

Private Sub txtserie_LostFocus()
    
    txtserie.Text = Format("" & txtserie.Text, "000")

End Sub
