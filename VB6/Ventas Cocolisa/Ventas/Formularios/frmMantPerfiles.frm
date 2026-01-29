VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantPerfiles 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Perfiles"
   ClientHeight    =   7500
   ClientLeft      =   3525
   ClientTop       =   1605
   ClientWidth     =   8085
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
   ScaleHeight     =   7500
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmMantPerfiles.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantPerfiles.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
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
      Height          =   6735
      Left            =   60
      TabIndex        =   6
      Top             =   660
      Width           =   7965
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
         Height          =   750
         Left            =   120
         TabIndex        =   7
         Top             =   150
         Width           =   7725
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   990
            TabIndex        =   8
            Top             =   255
            Width           =   6600
            _ExtentX        =   11642
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
            MaxLength       =   255
            Container       =   "frmMantPerfiles.frx":3518
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Búsqueda"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   9
            Top             =   300
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   5565
         Left            =   135
         OleObjectBlob   =   "frmMantPerfiles.frx":3534
         TabIndex        =   10
         Top             =   1035
         Width           =   7725
      End
   End
   Begin VB.Frame fraGeneral 
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
      Height          =   6720
      Left            =   60
      TabIndex        =   1
      Top             =   675
      Width           =   7965
      Begin CATControls.CATTextBox txtCod_Perfil 
         Height          =   315
         Left            =   6870
         TabIndex        =   2
         Tag             =   "TidPerfil"
         Top             =   225
         Width           =   1005
         _ExtentX        =   1773
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
         Container       =   "frmMantPerfiles.frx":4C50
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Perfil 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Tag             =   "TGlsPerfil"
         Top             =   690
         Width           =   6765
         _ExtentX        =   11933
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
         MaxLength       =   80
         Container       =   "frmMantPerfiles.frx":4C6C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin MSComctlLib.TreeView TV 
         Height          =   5460
         Left            =   135
         TabIndex        =   11
         Top             =   1125
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   9631
         _Version        =   393217
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblApePaterno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   135
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6285
         TabIndex        =   4
         Top             =   270
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   2170
      ButtonWidth     =   2540
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Nuevo         "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMantPerfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo As String
Dim strMsg      As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaHomonimia "perfil", "GlsPerfil", "idPerfil", txtGls_Perfil.Text, txtCod_Perfil.Text, True, StrMsgError, " CodSistema = '" & StrcodSistema & "' "
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_Perfil.Text = "" Then 'graba
        txtCod_Perfil.Text = GeneraCorrelativoAnoMes("perfil", "idPerfil")
        
        EjecutaSQLForm_1 Me, 0, True, "perfil", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        actualizaOpciones StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        strMsg = "Grabo"
    
    Else 'modifica
        EjecutaSQLForm_1 Me, 1, True, "perfil", StrMsgError, "idperfil"
        If StrMsgError <> "" Then GoTo Err
        
        actualizaOpciones StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        strMsg = "Modifico"
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    fraGeneral.Enabled = False
    listaPerfil StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLista, False, False, False, False
    listaPerfil StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 6
    nuevo
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub nuevo()
On Error GoTo Err
Dim StrMsgError As String

    limpiaForm Me
    MuestraOpciones StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarPerfil gLista.Columns.ColumnByName("idPerfil").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            nuevo
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            fraGeneral.Enabled = True
        Case 4, 6 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        Case 5 'Imprimir
            gLista.m.ExportToXLS App.Path & "\Temporales\Mantenimiento_Perfiles.xls"
            ShellEx App.Path & "\Temporales\Mantenimiento_Perfiles.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 7 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub habilitaBotones(indexBoton As Integer)
Dim indHabilitar As Boolean

    Select Case indexBoton
        Case 1, 2, 3 'Nuevo, Grabar, Modificar
            If indexBoton = 2 Then indHabilitar = True
            Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
            Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
            Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
            Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
            Toolbar1.Buttons(5).Visible = indHabilitar 'Imprimir
            Toolbar1.Buttons(6).Visible = indHabilitar 'Lista
        Case 4, 6 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = True
            Toolbar1.Buttons(6).Visible = False
    End Select

End Sub

Private Sub TV_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim i As Integer
Dim contNodos As Integer
Dim contMarca As Integer

    For i = 1 To TV.Nodes.Count
        If Node.Key = left(TV.Nodes(i).Key, Len(Node.Key)) Then
            TV.Nodes(i).Checked = Node.Checked
        End If
    Next
    
    If Len(Node.Key) > 5 Then
        contNodos = 0
        contMarca = 0
        
        For i = 1 To TV.Nodes.Count
            If left(Node.Key, Len(Node.Key) - 2) = left(TV.Nodes(i).Key, Len(Node.Key) - 2) And Len(Node.Key) = Len(TV.Nodes(i).Key) Then
                contNodos = contNodos + 1
                If TV.Nodes(i).Checked = Node.Checked Then
                    contMarca = contMarca + 1
                End If
            End If
        Next
        
        If (((contMarca = contNodos) And (Node.Checked = False)) Or ((contMarca >= 1) And (Node.Checked = True))) And contNodos <> 0 Then
            
            Node.Parent.Checked = Node.Checked
            If Node.Checked Then
                If Len(Node.Key) >= 9 Then
                    Node.Parent.Parent.Checked = Node.Checked
                End If
        
                If Len(Node.Key) = 11 Then
                    Node.Parent.Parent.Parent.Checked = Node.Checked
                End If
            End If
        End If
    End If

End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaPerfil StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaPerfil(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND GlsPerfil LIKE '%" & strCond & "%' "
    End If
    
    csql = "SELECT idPerfil ,GlsPerfil " & _
           "FROM perfil " & _
           "WHERE idEmpresa = '" & glsEmpresa & "'  AND cODsISTEMA = '" & StrcodSistema & "' " & strCond & _
           "ORDER BY idPerfil"
           

    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
        
    Set gLista.DataSource = rsdatos

'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idPerfil"
'    End With
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarPerfil(strCodPer As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim no As Node

    csql = "SELECT idPerfil,GlsPerfil " & _
           "FROM perfil " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idperfil = '" & strCodPer & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    MuestraOpciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If rst.State = 1 Then rst.Close
    rst.Open "select opmNum from opcionesperfil where idEmpresa = '" & glsEmpresa & "' AND idPerfil = '" & txtCod_Perfil.Text & "' AND CodSistema = '" & StrcodSistema & "' ", Cn, adOpenForwardOnly, adLockReadOnly
    Do While Not rst.EOF
        For Each no In TV.Nodes
            If no.Key = rst.Fields("opmNum") Then
                no.Checked = True
                Exit For
            End If
        Next
        rst.MoveNext
    Loop
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub actualizaOpciones(ByRef StrMsgError As String)
On Error GoTo Err
Dim no As Node

    Cn.Execute "delete from opcionesperfil where idEmpresa = '" & glsEmpresa & "' AND idPerfil = '" & txtCod_Perfil.Text & "' and CodSistema = '" & StrcodSistema & "' "
    For Each no In TV.Nodes
        If no.Checked Then
            csql = "INSERT INTO opcionesperfil (idPerfil,opmNum,idEmpresa,CodSistema) VALUES('" & txtCod_Perfil.Text & "','" & no.Key & "','" & glsEmpresa & "','" & StrcodSistema & "')"
            Cn.Execute csql
        End If
    Next
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub MuestraOpciones(ByRef StrMsgError As String)
On Error GoTo Err
Dim rs      As New ADODB.Recordset
Dim nodX    As Node

    csql = "SELECT opmNum,opmDes,opmCod " & _
           "FROM opcionesmenu WHERE opmEstado = 'S' and CodSistema = '" & StrcodSistema & _
           "' order by opmnum"
    rs.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    TV.Nodes.Clear
    Set nodX = TV.Nodes.Add(, , , "Marcar Todos")
    
    Do While Not rs.EOF
        If Len(rs.Fields(0)) = 5 Then
            Set nodX = TV.Nodes.Add(1, , rs.Fields(0), rs.Fields(1))
        End If
        If Len(rs.Fields(0)) > 5 Then
            Set nodX = TV.Nodes.Add(left(rs.Fields(0), Len(rs.Fields(0)) - 2), tvwChild, rs.Fields(0), rs.Fields(1))
        End If
        nodX.Expanded = True
        rs.MoveNext
    Loop
    TV.Nodes.Remove (1)
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
