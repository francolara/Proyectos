VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmCajasUsuario 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cajas por Usuario"
   ClientHeight    =   4860
   ClientLeft      =   2175
   ClientTop       =   1290
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   4050
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   8235
      Begin VB.CommandButton cmbAyudaPersona 
         Height          =   315
         Left            =   7725
         Picture         =   "frmCajasUsuario.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   390
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gCajas 
         Height          =   3090
         Left            =   120
         OleObjectBlob   =   "frmCajasUsuario.frx":038A
         TabIndex        =   1
         Top             =   840
         Width           =   8055
      End
      Begin CATControls.CATTextBox txtCod_Usuario 
         Height          =   315
         Left            =   930
         TabIndex        =   3
         Tag             =   "TidPersona"
         Top             =   360
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
         Container       =   "frmCajasUsuario.frx":1FBF
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Usuario 
         Height          =   315
         Left            =   1875
         TabIndex        =   4
         Top             =   360
         Width           =   5820
         _ExtentX        =   10266
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
         Container       =   "frmCajasUsuario.frx":1FDB
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   390
         Width           =   555
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   480
      Top             =   500
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
            Picture         =   "frmCajasUsuario.frx":1FF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":2391
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":27E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":2B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":2F17
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":32B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":364B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":39E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":3D7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":4119
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":44B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasUsuario.frx":5175
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
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
Attribute VB_Name = "frmCajasUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indInserta As Boolean

Private Sub cmbAyudaPersona_Click()
    
    mostrarAyuda "USUARIOS", txtCod_Usuario, txtGls_Usuario

End Sub

Private Sub Form_Load()

    ConfGrid gCajas, True, False, False, False
    indInserta = False
    nuevo

End Sub

Private Sub txtCod_Usuario_Change()
On Error GoTo Err
Dim StrMsgError As String

    If Trim(txtCod_Usuario.Text) <> "" Then
        mostrarCajas Trim(txtCod_Usuario.Text), StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Usuario_KeyPress(KeyAscii As Integer)
 
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "USUARIOS", txtCod_Usuario, txtGls_Usuario
        KeyAscii = 0
        If txtCod_Usuario.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub nuevo()
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim StrMsgError As String
    
    limpiaForm Me
    '--- FORMATO GRILLA
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "idCaja", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsCaja", adVarChar, 185, adFldIsNullable
    rst.Open
    
    rst.AddNew
    rst.Fields("Item") = 1
    rst.Fields("idCaja") = ""
    rst.Fields("GlsCaja") = ""
        
    mostrarDatosGridSQL gCajas, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    gCajas.Columns.FocusedIndex = gCajas.Columns.ColumnByFieldName("idCaja").Index
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarCajas(strCodUsu As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset

    '--- TRAE EL LISTADO DE CAJAS Y LO ALMACENA EN UN RECORSET
    csql = "SELECT s.item,s.idCaja,c.GlsCaja " & _
            "FROM cajasusuario s,cajas c " & _
            "WHERE s.idCaja = c.idCaja " & _
             "AND s.idUsuario = '" & strCodUsu & "' " & _
             "AND s.idEmpresa = '" & glsEmpresa & "' " & _
             "AND c.idEmpresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idCaja", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsCaja", adVarChar, 185, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idCaja") = ""
        rsg.Fields("GlsCaja") = ""
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idCaja") = "" & rst.Fields("idCaja")
            rsg.Fields("GlsCaja") = "" & rst.Fields("GlsCaja")
            rst.MoveNext
        Loop
    End If
    
    rst.Close
    Set rst = Nothing
    
    mostrarDatosGridSQL gCajas, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gCajas_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gCajas.Columns.ColumnByFieldName("item").Value = gCajas.Count
        gCajas.Dataset.Post
    End If

End Sub

Private Sub gCajas_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If (gCajas.Columns.ColumnByFieldName("idCaja").Value = "") And indInserta = False Then
            Allow = False
        Else
            gCajas.Columns.FocusedIndex = gCajas.Columns.ColumnByFieldName("idCaja").Index
        End If
    End If

End Sub

Private Sub gCajas_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrCod As String
Dim StrDes As String
    
    Select Case Column.Index
        Case gCajas.Columns.ColumnByFieldName("idCaja").Index
            StrCod = gCajas.Columns.ColumnByFieldName("idCaja").Value
            StrDes = gCajas.Columns.ColumnByFieldName("GlsCaja").Value
            
            mostrarAyudaTexto "CAJAS", StrCod, StrDes
            If existeEnGrilla(gCajas, "idCaja", StrCod) = False Then
                gCajas.Dataset.Edit
                gCajas.Columns.ColumnByFieldName("idCaja").Value = StrCod
                gCajas.Columns.ColumnByFieldName("GlsCaja").Value = StrDes
                gCajas.Dataset.Post
                gCajas.SetFocus
            Else
                MsgBox "La caja ya fue ingresada", vbInformation, App.Title
            End If
    End Select
    
End Sub

Private Sub gCajas_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gCajas.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                           
                If gCajas.Count = 1 Then
                    gCajas.Dataset.Edit
                    gCajas.Columns.ColumnByFieldName("Item").Value = 1
                    gCajas.Columns.ColumnByFieldName("idCaja").Value = ""
                    gCajas.Columns.ColumnByFieldName("GlsCaja").Value = ""
                    gCajas.Dataset.Post
                Else
                    gCajas.Dataset.Delete
                    
                    gCajas.Dataset.First
                    Do While Not gCajas.Dataset.EOF
                        i = i + 1
                        
                        gCajas.Dataset.Edit
                        gCajas.Columns.ColumnByFieldName("Item").Value = i
                        gCajas.Dataset.Post
                        
                        gCajas.Dataset.Next
                    Loop
                    If gCajas.Dataset.State = dsEdit Or gCajas.Dataset.State = dsInsert Then
                        gCajas.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gCajas.Dataset.State = dsEdit Or gCajas.Dataset.State = dsInsert Then
              gCajas.Dataset.Post
        End If
    End If

End Sub

Private Sub gCajas_OnKeyPress(Key As Integer)
Dim StrCod As String
Dim StrDes As String

    Select Case gCajas.Columns.FocusedColumn.Index
        Case gCajas.Columns.ColumnByFieldName("idCaja").Index
            StrCod = gCajas.Columns.ColumnByFieldName("idCaja").Value
            StrDes = gCajas.Columns.ColumnByFieldName("GlsCaja").Value
            
            mostrarAyudaKeyasciiTexto Key, "CAJAS", StrCod, StrDes
            Key = 0
            
    
            If existeEnGrilla(gCajas, "idCaja", StrCod) = False Then
                gCajas.Dataset.Edit
                gCajas.Columns.ColumnByFieldName("idCaja").Value = StrCod
                gCajas.Columns.ColumnByFieldName("GlsCaja").Value = StrDes
                gCajas.Dataset.Post
                gCajas.SetFocus
            Else
                MsgBox "La caja ya fue ingresada", vbInformation, App.Title
            End If
    End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 2 'Cancelar
           nuevo
        Case 3 'Salir
            Unload Me
    End Select
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indIniTrans As Boolean
Dim strSerieOri As String

    indIniTrans = False
    eliminaNulosGrilla
    
    If gCajas.Count >= 1 Then
        If gCajas.Count = 1 And gCajas.Columns.ColumnByFieldName("idCaja").Value = "" Then
            StrMsgError = "Falta Ingresar Cajas"
            GoTo Err
        End If
    End If
    
    gCajas.Dataset.First
    Cn.BeginTrans
    indIniTrans = True
    
    Cn.Execute "DELETE FROM cajasusuario WHERE idUsuario = '" & Trim(txtCod_Usuario.Text) & "' AND idEmpresa = '" & glsEmpresa & "'"
    
    Do While Not gCajas.Dataset.EOF
        Cn.Execute "INSERT INTO cajasusuario (idUsuario,idCaja,item,idEmpresa) VALUES(" & _
                   "'" & Trim(txtCod_Usuario.Text) & "','" & gCajas.Columns.ColumnByFieldName("idCaja").Value & "'," & gCajas.Columns.ColumnByFieldName("item").Value & ",'" & glsEmpresa & "')"
                  
        gCajas.Dataset.Next
    Loop
    Cn.CommitTrans
    MsgBox "Se grabo Satisfactoriamente", vbInformation, App.Title
    
    Exit Sub
    
Err:
    If indIniTrans = True Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminaNulosGrilla()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer
    
    indWhile = True
    Do While indWhile = True
        If gCajas.Count >= 1 Then
            gCajas.Dataset.First
            indEntro = False
            Do While Not gCajas.Dataset.EOF
                If Trim(gCajas.Columns.ColumnByFieldName("idCaja").Value) = "" Then
                    gCajas.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gCajas.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gCajas.Count >= 1 Then
        gCajas.Dataset.First
        i = 0
        Do While Not gCajas.Dataset.EOF
            i = i + 1
            gCajas.Dataset.Edit
            gCajas.Columns.ColumnByFieldName("item").Value = i
            If gCajas.Dataset.State = dsEdit Then gCajas.Dataset.Post
            gCajas.Dataset.Next
        Loop
    Else
        indInserta = True
        gCajas.Dataset.Append
        indInserta = False
    End If
    
End Sub
