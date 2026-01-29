VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantClaves 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Claves"
   ClientHeight    =   7740
   ClientLeft      =   4050
   ClientTop       =   1575
   ClientWidth     =   9585
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
   ScaleHeight     =   7740
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   9465
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   135
         TabIndex        =   4
         Top             =   135
         Width           =   9195
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            Caption         =   " Dígitos "
            ForeColor       =   &H80000008&
            Height          =   915
            Left            =   1500
            TabIndex        =   10
            Top             =   735
            Width           =   4725
            Begin VB.CommandButton cmd1 
               Caption         =   "1"
               Height          =   435
               Left            =   240
               TabIndex        =   16
               Top             =   300
               Width           =   675
            End
            Begin VB.CommandButton cmd5 
               Caption         =   "5"
               Height          =   435
               Left            =   3120
               TabIndex        =   15
               Top             =   300
               Width           =   675
            End
            Begin VB.CommandButton cmd4 
               Caption         =   "4"
               Height          =   435
               Left            =   2400
               TabIndex        =   14
               Top             =   300
               Width           =   675
            End
            Begin VB.CommandButton cmd3 
               Caption         =   "3"
               Height          =   435
               Left            =   1680
               TabIndex        =   13
               Top             =   300
               Width           =   675
            End
            Begin VB.CommandButton cmd2 
               Caption         =   "2"
               Height          =   435
               Left            =   960
               TabIndex        =   12
               Top             =   300
               Width           =   675
            End
            Begin VB.CommandButton cmd6 
               Caption         =   "6"
               Height          =   435
               Left            =   3840
               TabIndex        =   11
               Top             =   300
               Width           =   675
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            Caption         =   " Signos "
            ForeColor       =   &H80000008&
            Height          =   870
            Left            =   1500
            TabIndex        =   6
            Top             =   1755
            Width           =   4725
            Begin VB.CommandButton cmdmas 
               Caption         =   "+"
               Height          =   495
               Left            =   900
               TabIndex        =   9
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmdmenos 
               Caption         =   "-"
               Height          =   495
               Left            =   2040
               TabIndex        =   8
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmdpor 
               Caption         =   "x"
               Height          =   495
               Left            =   3180
               TabIndex        =   7
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.CommandButton cmbAyudaPersona 
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
            Left            =   7260
            Picture         =   "frmMantClaves.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   270
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Usuario 
            Height          =   315
            Left            =   1500
            TabIndex        =   17
            Tag             =   "TidPersona"
            Top             =   270
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
            Container       =   "frmMantClaves.frx":038A
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Usuario 
            Height          =   315
            Left            =   2460
            TabIndex        =   18
            Top             =   270
            Width           =   4755
            _ExtentX        =   8387
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
            Container       =   "frmMantClaves.frx":03A6
         End
         Begin CATControls.CATTextBox txtclave 
            Height          =   315
            Left            =   1500
            TabIndex        =   19
            Top             =   2910
            Width           =   1425
            _ExtentX        =   2514
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
            PasswordChar    =   "*"
            Container       =   "frmMantClaves.frx":03C2
         End
         Begin CATControls.CATTextBox txtgenerado 
            Height          =   315
            Left            =   5820
            TabIndex        =   20
            Top             =   2910
            Width           =   1335
            _ExtentX        =   2355
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
            PasswordChar    =   "*"
            Container       =   "frmMantClaves.frx":03DE
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Operación"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   720
            TabIndex        =   23
            Top             =   2970
            Width           =   750
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Generado"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4995
            TabIndex        =   22
            Top             =   2970
            Width           =   720
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   720
            TabIndex        =   21
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   " Permisos "
         ForeColor       =   &H00000000&
         Height          =   3270
         Left            =   150
         TabIndex        =   3
         Top             =   3630
         Width           =   9210
         Begin DXDBGRIDLibCtl.dxDBGrid gPermisos 
            Height          =   2865
            Left            =   120
            OleObjectBlob   =   "frmMantClaves.frx":03FA
            TabIndex        =   0
            Top             =   270
            Width           =   9030
         End
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   300
      Top             =   3840
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
            Picture         =   "frmMantClaves.frx":203B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":23D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":2827
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":2BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":2F5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":32F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":368F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":3A29
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":3DC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":415D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":44F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantClaves.frx":51B9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   1164
      ButtonWidth     =   2223
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "       Grabar       "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMantClaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim n1 As Variant, n2 As Variant, n3 As Variant, n4 As Variant, n5 As Variant, n6 As Variant
Dim signo1 As Variant, signo2 As Variant, signo3 As Variant
Dim swvalor As Boolean
Const ENCRYPT = 1
Const DECRYPT = 2

Private Sub cmbAyudaPersona_Click()
    
    mostrarAyuda "USUARIOS", txtCod_usuario, txtGls_usuario

End Sub

Private Sub cmd1_Click()
    
    n1 = ""
    n1 = 1
    txtclave.Text = txtclave.Text & n1

End Sub

Private Sub cmd2_Click()
    
    n2 = ""
    n2 = 2
    txtclave.Text = txtclave.Text & n2

End Sub

Private Sub cmd3_Click()
    
    n3 = ""
    n3 = 3
    txtclave.Text = txtclave.Text & n3

End Sub

Private Sub cmd4_Click()
    
    n4 = ""
    n4 = 4
    txtclave.Text = txtclave.Text & n4

End Sub

Private Sub cmd5_Click()
    
    n5 = ""
    n5 = 5
    txtclave.Text = txtclave.Text & n5

End Sub

Private Sub cmd6_Click()
    
    n6 = ""
    n6 = 6
    txtclave.Text = txtclave.Text & n6

End Sub

Private Sub cmdmas_Click()
    
    signo1 = ""
    signo1 = "+"
    txtclave.Text = txtclave.Text & signo1

End Sub

Private Sub cmdmenos_Click()
    
    signo2 = ""
    signo2 = "-"
    txtclave.Text = txtclave.Text & signo2

End Sub

Private Sub graba(ByRef StrMsgError As String)
On Error GoTo Err
Dim indIniTrans As Boolean

    indIniTrans = False
    eliminaNulosGrilla

    csql = "UPDATE usuarios SET Autogenerado='" & txtgenerado.Text & "' WHERE idUsuario ='" & Trim(txtCod_usuario.Text) & "' "
    Cn.Execute (csql)

    gPermisos.Dataset.First
    Cn.BeginTrans
    indIniTrans = True
    
    Cn.Execute "DELETE FROM permisosusuarios WHERE idEmpresa = '" & glsEmpresa & "' AND idUsuario = '" & Trim(txtCod_usuario.Text) & "' and CodSistema = '" & StrcodSistema & "'  "
    Do While Not gPermisos.Dataset.EOF
        Cn.Execute "INSERT INTO permisosusuarios (idUsuario,idPermiso,item,idEmpresa,CodSistema) VALUES(" & _
                   "'" & Trim(txtCod_usuario.Text) & "','" & gPermisos.Columns.ColumnByFieldName("idPermiso").Value & "'," & gPermisos.Columns.ColumnByFieldName("item").Value & ",'" & glsEmpresa & "','" & StrcodSistema & "')"
        gPermisos.Dataset.Next
    Loop
    
    Cn.CommitTrans
    MsgBox "Se registro satisfactoriamente", vbInformation, App.Title

    Exit Sub
    
Err:
    If indIniTrans = True Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Function encripta(texto As String, Tipo As String) As String
Dim Temp     As Integer
Dim i        As Integer
Dim j        As Integer
Dim n        As Integer
Dim rtn      As String, Password As String

    n = 6
    ReDim UserKeyASCIIS(1 To n)
    ReDim TextASCIIS(Len(texto)) As Integer
    If Tipo = "2" Then  'desencripta
        Pass = ""
        For i = 1 To Len(Trim(texto))
          Pass = Pass + Chr(Asc(Mid(texto, i, 1)) - 5)
        Next
        Password = Pass
    ElseIf Tipo = "1" Then 'encripta
        For i = 1 To Len(texto)
            rtn = rtn + Chr(Asc(Mid(texto, i, 1)) + 5)
        Next i
        Password = rtn
    End If
    encripta = Password

End Function

Private Sub cmdpor_Click()
    
    signo3 = ""
    signo3 = "*"
    txtclave.Text = txtclave.Text & signo3

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim csql As String

    ConfGrid gPermisos, True, False, False, False
    txtCod_usuario.Text = glsUser
    swvalor = False

End Sub

Private Sub gPermisos_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gPermisos.Columns.ColumnByFieldName("item").Value = gPermisos.Count
        gPermisos.Dataset.Post
    End If

End Sub

Private Sub gPermisos_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If (gPermisos.Columns.ColumnByFieldName("idPermiso").Value = "") Then
            Allow = False
        Else
            gPermisos.Columns.FocusedIndex = gPermisos.Columns.ColumnByFieldName("idPermiso").Index
        End If
    End If

End Sub

Private Sub gPermisos_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim StrDes As String
    
    Select Case Column.Index
        Case gPermisos.Columns.ColumnByFieldName("idPermiso").Index
            strCod = gPermisos.Columns.ColumnByFieldName("idPermiso").Value
            StrDes = gPermisos.Columns.ColumnByFieldName("GlsPermiso").Value
            
            mostrarAyudaTexto "PERMISOS", strCod, StrDes, " And  CodSistema = '" & StrcodSistema & "' "
            
            If existeEnGrilla(gPermisos, "idPermiso", strCod) = False Then
                gPermisos.Dataset.Edit
                gPermisos.Columns.ColumnByFieldName("idPermiso").Value = strCod
                gPermisos.Columns.ColumnByFieldName("GlsPermiso").Value = StrDes
                gPermisos.Dataset.Post
                gPermisos.SetFocus
            Else
                MsgBox "El Permiso ya fue registrado.", vbInformation, App.Title
            End If
    End Select
    
End Sub

Private Sub gPermisos_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gPermisos.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gPermisos.Count = 1 Then
                    gPermisos.Dataset.Edit
                    gPermisos.Columns.ColumnByFieldName("Item").Value = 1
                    gPermisos.Columns.ColumnByFieldName("idPermiso").Value = ""
                    gPermisos.Columns.ColumnByFieldName("GlPermiso").Value = ""
                    gPermisos.Dataset.Post
                
                Else
                    gPermisos.Dataset.Delete
                    gPermisos.Dataset.First
                    Do While Not gPermisos.Dataset.EOF
                        i = i + 1
                        gPermisos.Dataset.Edit
                        gPermisos.Columns.ColumnByFieldName("Item").Value = i
                        gPermisos.Dataset.Post
                        gPermisos.Dataset.Next
                    Loop
                    If gPermisos.Dataset.State = dsEdit Or gPermisos.Dataset.State = dsInsert Then
                        gPermisos.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gPermisos.Dataset.State = dsEdit Or gPermisos.Dataset.State = dsInsert Then
              gPermisos.Dataset.Post
        End If
    End If

End Sub

Private Sub gPermisos_OnKeyPress(Key As Integer)
Dim strCod As String
Dim StrDes As String

    Select Case gPermisos.Columns.FocusedColumn.Index
        Case gPermisos.Columns.ColumnByFieldName("idPermiso").Index
            strCod = gPermisos.Columns.ColumnByFieldName("idPermiso").Value
            StrDes = gPermisos.Columns.ColumnByFieldName("GlsPermiso").Value
            
            mostrarAyudaKeyasciiTexto Key, "PERMISOS", strCod, StrDes
            Key = 0
    
            If existeEnGrilla(gPermisos, "idPermiso", strCod) = False Then
                gPermisos.Dataset.Edit
                gPermisos.Columns.ColumnByFieldName("idPermiso").Value = strCod
                gPermisos.Columns.ColumnByFieldName("GlsPermiso").Value = StrDes
                gPermisos.Dataset.Post
                gPermisos.SetFocus
            Else
                MsgBox "El Permiso ya fue ingresado", vbInformation, App.Title
            End If
    End Select

End Sub

Private Sub txtCod_Usuario_GotFocus()
    
    txtCod_usuario.SelStart = 0: txtCod_usuario.SelLength = Len(txtCod_usuario.Text)

End Sub

Private Sub txtCod_Usuario_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_usuario.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_usuario.Text, False)
    If txtGls_usuario.Text <> "" Then
        txtclave.Text = encripta(traerCampo("usuarios", "Autogenerado", "idUsuario", txtCod_usuario.Text, True), DECRYPT)
    Else
        txtclave.Text = ""
    End If
    
    mostrarPermisos txtCod_usuario.Text, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Usuario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "USUARIOS", txtCod_usuario, txtGls_usuario
        KeyAscii = 0
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1
            txtgenerado.Text = encripta(txtclave.Text, ENCRYPT)
            graba StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2
            txtclave.Text = ""
            txtgenerado.Text = ""
        Case 3
            Unload Me
    End Select
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarPermisos(ByVal strCodUsuario As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsu As New ADODB.Recordset
Dim rst As New ADODB.Recordset

    csql = "SELECT u.item, u.idPermiso, p.GlsPermiso " & _
           "FROM permisosusuarios u, permisos p " & _
           "WHERE u.idEmpresa = '" & glsEmpresa & "' AND u.idPermiso = p.idPermiso AND u.CodSistema = p.CodSistema AND u.idUsuario = '" & strCodUsuario & "' And U.CodSistema = '" & StrcodSistema & "' "
    
    rsu.Open csql, Cn, adOpenForwardOnl, adLockReadOnly
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "idPermiso", adVarChar, 2, adFldIsNullable
    rst.Fields.Append "GlsPermiso", adVarChar, 180, adFldIsNullable
    rst.Open
    
    If rsu.EOF Then
        rst.AddNew
        rst.Fields("Item") = 1
        rst.Fields("idPermiso") = ""
        rst.Fields("GlsPermiso") = ""
    Else
        Do While Not rsu.EOF
            rst.AddNew
            rst.Fields("Item") = rsu.Fields("Item")
            rst.Fields("idPermiso") = "" & rsu.Fields("idPermiso")
            rst.Fields("GlsPermiso") = "" & rsu.Fields("GlsPermiso")
            rsu.MoveNext
        Loop
    End If
    
    mostrarDatosGridSQL gPermisos, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gPermisos.Columns.FocusedIndex = gPermisos.Columns.ColumnByFieldName("idPermiso").Index
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminaNulosGrilla()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer

    indWhile = True
    Do While indWhile = True
        If gPermisos.Count >= 1 Then
            gPermisos.Dataset.First
            indEntro = False
            Do While Not gPermisos.Dataset.EOF
                If Trim(gPermisos.Columns.ColumnByFieldName("idPermiso").Value) = "" Then
                    gPermisos.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gPermisos.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gPermisos.Count >= 1 Then
        gPermisos.Dataset.First
        i = 0
        Do While Not gPermisos.Dataset.EOF
            i = i + 1
            gPermisos.Dataset.Edit
            gPermisos.Columns.ColumnByFieldName("item").Value = i
            If gPermisos.Dataset.State = dsEdit Then gPermisos.Dataset.Post
            gPermisos.Dataset.Next
        Loop
    Else
        indInserta = True
        gPermisos.Dataset.Append
        indInserta = False
    End If
    
End Sub
