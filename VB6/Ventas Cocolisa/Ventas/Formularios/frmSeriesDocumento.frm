VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmSeriesDocumento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Número Máximo de Registros"
   ClientHeight    =   4875
   ClientLeft      =   1785
   ClientTop       =   2160
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   4050
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   8925
      Begin DXDBGRIDLibCtl.dxDBGrid gSeries 
         Height          =   3750
         Left            =   90
         OleObjectBlob   =   "frmSeriesDocumento.frx":0000
         TabIndex        =   1
         Top             =   180
         Width           =   8790
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
            Picture         =   "frmSeriesDocumento.frx":2AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":2E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":328C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":3626
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":39C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":3D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":40F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":448E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":4828
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":4BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":4F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeriesDocumento.frx":5C1E
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
      Width           =   9060
      _ExtentX        =   15981
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
Attribute VB_Name = "frmSeriesDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gSeries, True, False, False, False
    mostrarSeries StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarSeries(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset

    csql = "SELECT s.item,s.idDocumento,d.GlsDocumento,s.idSerie, s.numRegMaximo, s.espacioLineasImp " & _
            "FROM seriesdocumento s,documentos d " & _
            "WHERE s.idDocumento = d.idDocumento " & _
            "AND s.idEmpresa = '" & glsEmpresa & "'" & _
            "AND s.idSucursal = '" & glsSucursal & "'"
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idDocumento", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsDocumento", adVarChar, 250, adFldIsNullable
    rsg.Fields.Append "idSerie", adVarChar, 4, adFldIsNullable
    rsg.Fields.Append "numRegMaximo", adInteger, , adFldIsNullable
    rsg.Fields.Append "espacioLineasImp", adInteger, , adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idDocumento") = ""
        rsg.Fields("GlsDocumento") = ""
        rsg.Fields("idSerie") = ""
        rsg.Fields("numRegMaximo") = 0
        rsg.Fields("espacioLineasImp") = 0
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idDocumento") = "" & rst.Fields("idDocumento")
            rsg.Fields("GlsDocumento") = "" & rst.Fields("GlsDocumento")
            rsg.Fields("idSerie") = "" & rst.Fields("idSerie")
            rsg.Fields("numRegMaximo") = "" & rst.Fields("numRegMaximo")
            rsg.Fields("espacioLineasImp") = "" & rst.Fields("espacioLineasImp")
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
            mostrarSeries StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Salir
            Unload Me
    End Select
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indIniTrans As Boolean
Dim glsCodEtiqueta As String

    indIniTrans = False
    eliminaNulosGrilla
    gSeries.Dataset.First
    
    Cn.BeginTrans
    
    indIniTrans = True
    Cn.Execute "DELETE FROM seriesdocumento WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'"
    Do While Not gSeries.Dataset.EOF
        Cn.Execute "INSERT INTO seriesdocumento (idSucursal,idDocumento,idSerie,item,idEmpresa,numRegMaximo,espacioLineasImp) VALUES(" & _
                   "'" & glsSucursal & "','" & gSeries.Columns.ColumnByFieldName("idDocumento").Value & "','" & gSeries.Columns.ColumnByFieldName("idSerie").Value & "'," & gSeries.Columns.ColumnByFieldName("item").Value & ",'" & glsEmpresa & "'," & gSeries.Columns.ColumnByFieldName("numRegMaximo").Value & "," & gSeries.Columns.ColumnByFieldName("espacioLineasImp").Value & ")"
        gSeries.Dataset.Next
    Loop
    Cn.CommitTrans
    
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
