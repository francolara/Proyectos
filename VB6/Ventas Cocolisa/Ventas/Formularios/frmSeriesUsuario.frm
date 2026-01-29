VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmSeriesUsuario 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Series por Usuario"
   ClientHeight    =   6525
   ClientLeft      =   4410
   ClientTop       =   1845
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   7695
      Top             =   180
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
            Picture         =   "frmSeriesUsuario.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesUsuario.frx":317E
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
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   1164
      ButtonWidth     =   2619
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Grabar          "
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   5715
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   9495
      Begin VB.CommandButton cmbAyudaPersona 
         Height          =   315
         Left            =   8985
         Picture         =   "frmSeriesUsuario.frx":3518
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   255
         Width           =   390
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gSeries 
         Height          =   4890
         Left            =   120
         OleObjectBlob   =   "frmSeriesUsuario.frx":38A2
         TabIndex        =   1
         Top             =   705
         Width           =   9270
      End
      Begin CATControls.CATTextBox txtCod_Usuario 
         Height          =   315
         Left            =   720
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
         Container       =   "frmSeriesUsuario.frx":59A2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Usuario 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   255
         Width           =   7305
         _ExtentX        =   12885
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
         Container       =   "frmSeriesUsuario.frx":59BE
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
         Left            =   105
         TabIndex        =   5
         Top             =   315
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmSeriesUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indInserta As Boolean

Private Sub cmbAyudaPersona_Click()
    
    mostrarAyuda "USUARIOS", txtCod_usuario, txtGls_usuario

End Sub

Private Sub Form_Load()

    ConfGrid gSeries, True, False, False, False
    indInserta = False
    nuevo

End Sub

Private Sub txtCod_Usuario_Change()
On Error GoTo Err
Dim StrMsgError As String

    If Trim(txtCod_usuario.Text) <> "" Then
        mostrarSeries Trim(txtCod_usuario.Text), StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Usuario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "USUARIOS", txtCod_usuario, txtGls_usuario
        KeyAscii = 0
        If txtCod_usuario.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub nuevo()
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim StrMsgError As String

    limpiaForm Me
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "idDocumento", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsDocumento", adVarChar, 250, adFldIsNullable
    rst.Fields.Append "idSerie", adVarChar, 4, adFldIsNullable
    rst.Open
    
    rst.AddNew
    rst.Fields("Item") = 1
    rst.Fields("idDocumento") = ""
    rst.Fields("GlsDocumento") = ""
    rst.Fields("idSerie") = ""
    
    mostrarDatosGridSQL gSeries, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gSeries.Columns.FocusedIndex = gSeries.Columns.ColumnByFieldName("idDocumento").Index
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarSeries(strCodUsu As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset

    csql = "SELECT s.item,s.idDocumento,d.GlsDocumento,s.idSerie " & _
            "FROM seriexusuario s,documentos d " & _
            "WHERE s.idDocumento = d.idDocumento " & _
            "AND s.idUsuario = '" & strCodUsu & "' " & _
            "AND s.idEmpresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idDocumento", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsDocumento", adVarChar, 250, adFldIsNullable
    rsg.Fields.Append "idSerie", adVarChar, 4, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idDocumento") = ""
        rsg.Fields("GlsDocumento") = ""
        rsg.Fields("idSerie") = ""
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idDocumento") = "" & rst.Fields("idDocumento")
            rsg.Fields("GlsDocumento") = "" & rst.Fields("GlsDocumento")
            rsg.Fields("idSerie") = "" & rst.Fields("idSerie")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gSeries, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gSeries_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gSeries.Columns.ColumnByFieldName("item").Value = gSeries.Count
        gSeries.Dataset.Post
    End If

End Sub

Private Sub gSeries_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If (gSeries.Columns.ColumnByFieldName("idDocumento").Value = "" Or gSeries.Columns.ColumnByFieldName("idSerie").Value = "") And indInserta = False Then
            Allow = False
        Else
            gSeries.Columns.FocusedIndex = gSeries.Columns.ColumnByFieldName("idDocumento").Index
        End If
    End If
    
End Sub

Private Sub gSeries_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim StrDes As String

    Select Case Column.Index
        Case gSeries.Columns.ColumnByFieldName("idDocumento").Index
            strCod = gSeries.Columns.ColumnByFieldName("idDocumento").Value
            StrDes = gSeries.Columns.ColumnByFieldName("GlsDocumento").Value
            mostrarAyudaTexto "DOCUMENTOS", strCod, StrDes
            
            If existeEnGrilla(gSeries, "idDocumento", strCod) = False Then
                gSeries.Dataset.Edit
                gSeries.Columns.ColumnByFieldName("idDocumento").Value = strCod
                gSeries.Columns.ColumnByFieldName("GlsDocumento").Value = StrDes
                gSeries.Dataset.Post
                gSeries.SetFocus
            Else
                MsgBox "El Documento ya fue ingresado", vbInformation, App.Title
            End If
    End Select
    
End Sub

Private Sub gSeries_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gSeries.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gSeries.Count = 1 Then
                    gSeries.Dataset.Edit
                    gSeries.Columns.ColumnByFieldName("Item").Value = 1
                    gSeries.Columns.ColumnByFieldName("idDocumento").Value = ""
                    gSeries.Columns.ColumnByFieldName("GlsDocumento").Value = ""
                    gSeries.Columns.ColumnByFieldName("idSerie").Value = ""
                    gSeries.Dataset.Post
                
                Else
                    gSeries.Dataset.Delete
                    gSeries.Dataset.First
                    Do While Not gSeries.Dataset.EOF
                        i = i + 1
                        gSeries.Dataset.Edit
                        gSeries.Columns.ColumnByFieldName("Item").Value = i
                        gSeries.Dataset.Post
                        gSeries.Dataset.Next
                    Loop
                    If gSeries.Dataset.State = dsEdit Or gSeries.Dataset.State = dsInsert Then
                        gSeries.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gSeries.Dataset.State = dsEdit Or gSeries.Dataset.State = dsInsert Then
              gSeries.Dataset.Post
        End If
    End If

End Sub

Private Sub gSeries_OnKeyPress(Key As Integer)
Dim strCod As String
Dim StrDes As String

    Select Case gSeries.Columns.FocusedColumn.Index
        Case gSeries.Columns.ColumnByFieldName("idDocumento").Index
            strCod = gSeries.Columns.ColumnByFieldName("idDocumento").Value
            StrDes = gSeries.Columns.ColumnByFieldName("GlsDocumento").Value
            
            mostrarAyudaKeyasciiTexto Key, "DOCUMENTOS", strCod, StrDes
            Key = 0
    
            If existeEnGrilla(gSeries, "idDocumento", strCod) = False Then
                gSeries.Dataset.Edit
                gSeries.Columns.ColumnByFieldName("idDocumento").Value = strCod
                gSeries.Columns.ColumnByFieldName("GlsDocumento").Value = StrDes
                gSeries.Dataset.Post
                gSeries.SetFocus
            Else
                MsgBox "El Documento ya fue ingresado", vbInformation, App.Title
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
Dim rst             As New ADODB.Recordset
Dim indIniTrans     As Boolean
Dim strSerieOri     As String
Dim glsCodEtiqueta  As String
Dim csql2           As String

    indIniTrans = False
    eliminaNulosGrilla
    
    If gSeries.Count >= 1 Then
        If gSeries.Count = 1 And gSeries.Columns.ColumnByFieldName("idDocumento").Value = "" Then
            StrMsgError = "Falta Ingresar Series"
            GoTo Err
        End If
    End If
    
    gSeries.Dataset.First
    Cn.BeginTrans
    indIniTrans = True
    
    Cn.Execute "DELETE FROM seriexusuario WHERE idEmpresa = '" & glsEmpresa & "' AND idUsuario = '" & Trim(txtCod_usuario.Text) & "'"
    
    Do While Not gSeries.Dataset.EOF
        Cn.Execute "INSERT INTO seriexusuario (idUsuario,idDocumento,idSerie,item,idEmpresa) VALUES(" & _
                   "'" & Trim(txtCod_usuario.Text) & "','" & gSeries.Columns.ColumnByFieldName("idDocumento").Value & "','" & gSeries.Columns.ColumnByFieldName("idSerie").Value & "'," & gSeries.Columns.ColumnByFieldName("item").Value & ",'" & glsEmpresa & "')"
                   
        If Trim("" & traerCampo("objdocventas", "idDocumento", "idDocumento", gSeries.Columns.ColumnByFieldName("idDocumento").Value, True, " idserie = '" & gSeries.Columns.ColumnByFieldName("idSerie").Value & "'")) = "" Then
            strSerieOri = traerCampo("documentos", "idSerie", "idDocumento", gSeries.Columns.ColumnByFieldName("idDocumento").Value, False)
            
'            csql = "INSERT INTO objdocventas (idEmpresa, idDocumento, idSerie, tipoObj, GlsObj, GlsCampo, indVisible, intLeft, intTop, tipoDato, Etiqueta, numCol, ancho, Decimales, indImprime, impX, impY, impLongitud,intTabIndex,GlsObs) " & _
'                   "SELECT '" & glsEmpresa & "','" & gSeries.Columns.ColumnByFieldName("idDocumento").Value & "', '" & gSeries.Columns.ColumnByFieldName("idSerie").Value & "', tipoObj, GlsObj, GlsCampo, indVisible, intLeft, intTop, tipoDato, Etiqueta, numCol, ancho, Decimales, indImprime, impX, impY, impLongitud,intTabIndex,GlsObs FROM objdocventas " & _
'                   "WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & gSeries.Columns.ColumnByFieldName("idDocumento").Value & "' and  idSerie = '" & strSerieOri & "'"
'            Cn.Execute csql
'
'            csql = "SELECT idObjEtiquetasDoc,idEmpresa,idDocumento,idSerie,Etiqueta,impX,impY,GlsObs FROM objetiquetasdoc WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & gSeries.Columns.ColumnByFieldName("idDocumento").Value & "' and  idSerie = '" & gSeries.Columns.ColumnByFieldName("idSerie").Value & "' Order By idObjEtiquetasDoc "
'            If rst.State = 1 Then rst.Close
'            rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
'            If Not rst.EOF Then
'                Do While Not rst.EOF
'                    csql = "Update objetiquetasdoc Set  Etiqueta = '" & rst.Fields("Etiqueta").Value & "', impX ='" & rst.Fields("impX").Value & "' ,impY= '" & rst.Fields("impY").Value & "',GlsObs= '" & rst.Fields("GlsObs").Value & "' " & _
'                           "Where idObjEtiquetasDoc = '" & rst.Fields("idObjEtiquetasDoc").Value & "' And idEmpresa = '" & glsEmpresa & "' And idDocumento = '" & gSeries.Columns.ColumnByFieldName("idDocumento").Value & "' And  idSerie = '" & gSeries.Columns.ColumnByFieldName("idSerie").Value & "' "
'                    rst.MoveNext
'                Loop
'            End If
        End If
        gSeries.Dataset.Next
    Loop
    Cn.CommitTrans
    
    MsgBox "Se registro satisfactoriamente", vbInformation, App.Title
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If indIniTrans = True Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminaNulosGrilla()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer
    
    indWhile = True
    Do While indWhile = True
        If gSeries.Count >= 1 Then
            gSeries.Dataset.First
            indEntro = False
            Do While Not gSeries.Dataset.EOF
                If Trim(gSeries.Columns.ColumnByFieldName("idDocumento").Value) = "" Or Trim(gSeries.Columns.ColumnByFieldName("idSerie").Value) = "" Then
                    gSeries.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gSeries.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gSeries.Count >= 1 Then
        gSeries.Dataset.First
        i = 0
        Do While Not gSeries.Dataset.EOF
            i = i + 1
            gSeries.Dataset.Edit
            gSeries.Columns.ColumnByFieldName("item").Value = i
            If gSeries.Dataset.State = dsEdit Then gSeries.Dataset.Post
            gSeries.Dataset.Next
        Loop
    Else
        indInserta = True
        gSeries.Dataset.Append
        indInserta = False
    End If
    
End Sub
