VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmDocumentoExportar 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos a Importar por Documento"
   ClientHeight    =   4860
   ClientLeft      =   3360
   ClientTop       =   2010
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
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   8235
      Begin VB.CommandButton cmbAyudaDocumento 
         Height          =   315
         Left            =   7725
         Picture         =   "frmDocumentoExportar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   255
         Width           =   390
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gDocumentos 
         Height          =   3270
         Left            =   120
         OleObjectBlob   =   "frmDocumentoExportar.frx":038A
         TabIndex        =   1
         Top             =   705
         Width           =   8055
      End
      Begin CATControls.CATTextBox txtCod_Documento 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Tag             =   "TidPersona"
         Top             =   255
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
         Container       =   "frmDocumentoExportar.frx":1FE7
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Documento 
         Height          =   315
         Left            =   2010
         TabIndex        =   4
         Top             =   255
         Width           =   5685
         _ExtentX        =   10028
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
         Container       =   "frmDocumentoExportar.frx":2003
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Documento"
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
         Left            =   180
         TabIndex        =   5
         Top             =   285
         Width           =   810
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
            Picture         =   "frmDocumentoExportar.frx":201F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":23B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":280B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":2BA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":2F3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":32D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":3673
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":3A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":3DA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":4141
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":44DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentoExportar.frx":519D
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
      ButtonWidth     =   2143
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "       Grabar      "
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
Attribute VB_Name = "frmDocumentoExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indInserta As Boolean

Private Sub cmbAyudaDocumento_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid GDocumentos, True, False, False, False
    indInserta = False
    nuevo StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Documento_Change()
On Error GoTo Err
Dim StrMsgError As String

    If Trim(txtCod_Documento.Text) <> "" Then
        mostrarDocumentos Trim(txtCod_Documento.Text), StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Documento_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DOCUMENTOS", txtCod_Documento, txtGls_Documento
        KeyAscii = 0
        If txtCod_Documento.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub nuevo(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
    
    limpiaForm Me
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "idDocumento", adVarChar, 2, adFldIsNullable
    rst.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    rst.Open
    
    rst.AddNew
    rst.Fields("Item") = 1
    rst.Fields("idDocumento") = ""
    rst.Fields("GlsDocumento") = ""
        
    mostrarDatosGridSQL GDocumentos, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    GDocumentos.Columns.FocusedIndex = GDocumentos.Columns.ColumnByFieldName("idDocumento").Index
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarDocumentos(strCod As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset

    csql = "SELECT e.item,e.idDocumentoExp,c.GlsDocumento " & _
            "FROM documentosexportar e,documentos c " & _
            "WHERE e.idDocumentoExp = c.idDocumento " & _
            "AND e.idDocumento = '" & strCod & "'"
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idDocumento", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "GlsDocumento", adVarChar, 185, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idDocumento") = ""
        rsg.Fields("GlsDocumento") = ""
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idDocumento") = "" & rst.Fields("idDocumentoExp")
            rsg.Fields("GlsDocumento") = "" & rst.Fields("GlsDocumento")
            
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    mostrarDatosGridSQL GDocumentos, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gDocumentos_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        GDocumentos.Columns.ColumnByFieldName("item").Value = GDocumentos.Count
        GDocumentos.Dataset.Post
    End If

End Sub

Private Sub gDocumentos_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If (GDocumentos.Columns.ColumnByFieldName("idDocumento").Value = "") And indInserta = False Then
            Allow = False
        Else
            GDocumentos.Columns.FocusedIndex = GDocumentos.Columns.ColumnByFieldName("idDocumento").Index
        End If
    End If

End Sub

Private Sub gDocumentos_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim StrDes As String
    
    Select Case Column.Index
        Case GDocumentos.Columns.ColumnByFieldName("idDocumento").Index
            strCod = GDocumentos.Columns.ColumnByFieldName("idDocumento").Value
            StrDes = GDocumentos.Columns.ColumnByFieldName("GlsDocumento").Value
            
            mostrarAyudaTexto "DOCUMENTOS", strCod, StrDes
            If existeEnGrilla(GDocumentos, "idDocumento", strCod) = False Then
                GDocumentos.Dataset.Edit
                GDocumentos.Columns.ColumnByFieldName("idDocumento").Value = strCod
                GDocumentos.Columns.ColumnByFieldName("GlsDocumento").Value = StrDes
                GDocumentos.Dataset.Post
                GDocumentos.SetFocus
            Else
                MsgBox "El Documento ya fue ingresado", vbInformation, App.Title
            End If
    End Select
    
End Sub

Private Sub gDocumentos_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If GDocumentos.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If GDocumentos.Count = 1 Then
                    GDocumentos.Dataset.Edit
                    GDocumentos.Columns.ColumnByFieldName("Item").Value = 1
                    GDocumentos.Columns.ColumnByFieldName("idDocumento").Value = ""
                    GDocumentos.Columns.ColumnByFieldName("GlsDocumento").Value = ""
                    GDocumentos.Dataset.Post
                
                Else
                    GDocumentos.Dataset.Delete
                    GDocumentos.Dataset.First
                    Do While Not GDocumentos.Dataset.EOF
                        i = i + 1
                        GDocumentos.Dataset.Edit
                        GDocumentos.Columns.ColumnByFieldName("Item").Value = i
                        GDocumentos.Dataset.Post
                        GDocumentos.Dataset.Next
                    Loop
                    If GDocumentos.Dataset.State = dsEdit Or GDocumentos.Dataset.State = dsInsert Then
                        GDocumentos.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If GDocumentos.Dataset.State = dsEdit Or GDocumentos.Dataset.State = dsInsert Then
              GDocumentos.Dataset.Post
        End If
    End If
    
End Sub

Private Sub gDocumentos_OnKeyPress(Key As Integer)
Dim strCod As String
Dim StrDes As String

    Select Case GDocumentos.Columns.FocusedColumn.Index
        Case GDocumentos.Columns.ColumnByFieldName("idDocumento").Index
            strCod = GDocumentos.Columns.ColumnByFieldName("idDocumento").Value
            StrDes = GDocumentos.Columns.ColumnByFieldName("GlsDocumento").Value
            
            mostrarAyudaKeyasciiTexto Key, "DOCUMENTOS", strCod, StrDes
            Key = 0
            If existeEnGrilla(GDocumentos, "idCaja", strCod) = False Then
                GDocumentos.Dataset.Edit
                GDocumentos.Columns.ColumnByFieldName("idDocumento").Value = strCod
                GDocumentos.Columns.ColumnByFieldName("GlsDocumento").Value = StrDes
                GDocumentos.Dataset.Post
                GDocumentos.SetFocus
            Else
                MsgBox "El documento ya fue ingresado", vbInformation, App.Title
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
           nuevo StrMsgError
            If StrMsgError <> "" Then GoTo Err
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
    
    If GDocumentos.Count >= 1 Then
        If GDocumentos.Count = 1 And GDocumentos.Columns.ColumnByFieldName("idDocumento").Value = "" Then
            StrMsgError = "Falta Ingresar Documentos"
            GoTo Err
        End If
    End If
    
    GDocumentos.Dataset.First
    Cn.BeginTrans
    indIniTrans = True
    
    Cn.Execute "DELETE FROM documentosexportar WHERE idDocumento = '" & Trim(txtCod_Documento.Text) & "'"
    Do While Not GDocumentos.Dataset.EOF
        Cn.Execute "INSERT INTO documentosexportar (idDocumento,idDocumentoExp,item) VALUES(" & _
                   "'" & Trim(txtCod_Documento.Text) & "','" & GDocumentos.Columns.ColumnByFieldName("idDocumento").Value & "'," & GDocumentos.Columns.ColumnByFieldName("item").Value & ")"
        GDocumentos.Dataset.Next
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
        If GDocumentos.Count >= 1 Then
            GDocumentos.Dataset.First
            indEntro = False
            Do While Not GDocumentos.Dataset.EOF
                If Trim(GDocumentos.Columns.ColumnByFieldName("idDocumento").Value) = "" Then
                    GDocumentos.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                GDocumentos.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If GDocumentos.Count >= 1 Then
        GDocumentos.Dataset.First
        i = 0
        Do While Not GDocumentos.Dataset.EOF
            i = i + 1
            GDocumentos.Dataset.Edit
            GDocumentos.Columns.ColumnByFieldName("item").Value = i
            If GDocumentos.Dataset.State = dsEdit Then GDocumentos.Dataset.Post
            GDocumentos.Dataset.Next
        Loop
    Else
        indInserta = True
        GDocumentos.Dataset.Append
        indInserta = False
    End If
    
End Sub
