VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmMantAlmacenesVtas 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Almacenes de Ventas por Sucursal"
   ClientHeight    =   4785
   ClientLeft      =   2700
   ClientTop       =   1995
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   240
      Top             =   840
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
            Picture         =   "frmMantAlmacenesVtas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantAlmacenesVtas.frx":317E
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
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   1164
      ButtonWidth     =   2461
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "        Grabar         "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
      Height          =   4050
      Left            =   60
      TabIndex        =   0
      Top             =   690
      Width           =   10125
      Begin DXDBGRIDLibCtl.dxDBGrid gAlm 
         Height          =   3675
         Left            =   120
         OleObjectBlob   =   "frmMantAlmacenesVtas.frx":3518
         TabIndex        =   1
         Top             =   300
         Width           =   9915
      End
   End
End
Attribute VB_Name = "frmMantAlmacenesVtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indInserta As Boolean

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    Me.left = 0
    Me.top = 0
    ConfGrid gAlm, True, False, False, False
    indInserta = False
    nuevo
    mostrarAlm StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub nuevo()
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim StrMsgError As String
    
    limpiaForm Me
    
    rst.Fields.Append "Item", adInteger, , adFldRowID
    rst.Fields.Append "idSucursal", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsSucursal", adVarChar, 185, adFldIsNullable
    rst.Fields.Append "idAlmacen", adVarChar, 8, adFldIsNullable
    rst.Fields.Append "GlsAlmacen", adVarChar, 185, adFldIsNullable
    rst.Open
    
    rst.AddNew
    rst.Fields("Item") = 1
    rst.Fields("idSucursal") = ""
    rst.Fields("GlsSucursal") = ""
    rst.Fields("idAlmacen") = ""
    rst.Fields("GlsAlmacen") = ""
    
    mostrarDatosGridSQL gAlm, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gAlm.Columns.FocusedIndex = gAlm.Columns.ColumnByFieldName("idSucursal").Index
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarAlm(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
    
    csql = "SELECT v.item, v.idSucursal, s.GlsPersona AS GlsSucursal, v.idAlmacen, a.GlsAlmacen " & _
            "FROM AlmacenesVtas v,personas s,almacenes a " & _
            "WHERE v.idSucursal = s.idPersona " & _
            "AND v.idAlmacen = a.idAlmacen " & _
            "AND v.idEmpresa = '" & glsEmpresa & "' " & _
            "AND a.idEmpresa = '" & glsEmpresa & "' " & _
            "AND v.idEmpresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
        
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idSucursal", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsSucursal", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "idAlmacen", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsAlmacen", adVarChar, 185, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("idSucursal") = ""
        rsg.Fields("GlsSucursal") = ""
        rsg.Fields("idAlmacen") = ""
        rsg.Fields("GlsAlmacen") = ""
    Else
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = "" & rst.Fields("Item")
            rsg.Fields("idSucursal") = "" & rst.Fields("idSucursal")
            rsg.Fields("GlsSucursal") = "" & rst.Fields("GlsSucursal")
            rsg.Fields("idAlmacen") = "" & rst.Fields("idAlmacen")
            rsg.Fields("GlsAlmacen") = "" & rst.Fields("GlsAlmacen")
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gAlm, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gAlm_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
Dim i As Integer

    If Action = daInsert Then
        gAlm.Columns.ColumnByFieldName("item").Value = gAlm.Count
        gAlm.Dataset.Post
    End If

End Sub

Private Sub gAlm_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If Action = daInsert Then
        If (gAlm.Columns.ColumnByFieldName("idSucursal").Value = "" And gAlm.Columns.ColumnByFieldName("idAlmacen").Value = "") And indInserta = False Then
            Allow = False
        Else
            gAlm.Columns.FocusedIndex = gAlm.Columns.ColumnByFieldName("idSucursal").Index
        End If
    End If

End Sub

Private Sub gAlm_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strCod As String
Dim strDes As String

    Select Case Column.Index
        Case gAlm.Columns.ColumnByFieldName("idSucursal").Index
            strCod = gAlm.Columns.ColumnByFieldName("idSucursal").Value
            strDes = gAlm.Columns.ColumnByFieldName("GlsSucursal").Value
            mostrarAyudaTexto "SUCURSAL", strCod, strDes
            If existeEnGrilla(gAlm, "idSucursal", strCod) = False Then
                gAlm.Dataset.Edit
                gAlm.Columns.ColumnByFieldName("idSucursal").Value = strCod
                gAlm.Columns.ColumnByFieldName("GlsSucursal").Value = strDes
                gAlm.Dataset.Post
                gAlm.SetFocus
            Else
                MsgBox "La sucursal ya fue ingresada", vbInformation, App.Title
            End If
        
        Case gAlm.Columns.ColumnByFieldName("idAlmacen").Index
            strCod = gAlm.Columns.ColumnByFieldName("idAlmacen").Value
            strDes = gAlm.Columns.ColumnByFieldName("GlsAlmacen").Value
            mostrarAyudaTexto "ALMACENVTA", strCod, strDes, " AND idSucursal = '" & gAlm.Columns.ColumnByFieldName("idSucursal").Value & "'"
            gAlm.Dataset.Edit
            gAlm.Columns.ColumnByFieldName("idAlmacen").Value = strCod
            gAlm.Columns.ColumnByFieldName("GlsAlmacen").Value = strDes
            gAlm.Dataset.Post
            gAlm.SetFocus
    End Select
    
End Sub

Private Sub gAlm_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer

    If KeyCode = 46 Then
        If gAlm.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                If gAlm.Count = 1 Then
                    gAlm.Dataset.Edit
                    gAlm.Columns.ColumnByFieldName("Item").Value = 1
                    gAlm.Columns.ColumnByFieldName("idSucursal").Value = ""
                    gAlm.Columns.ColumnByFieldName("GlsSucursal").Value = ""
                    gAlm.Columns.ColumnByFieldName("idAlmacen").Value = ""
                    gAlm.Columns.ColumnByFieldName("GlsAlmacen").Value = ""
                    gAlm.Dataset.Post
                
                Else
                    gAlm.Dataset.Delete
                    gAlm.Dataset.First
                    Do While Not gAlm.Dataset.EOF
                        i = i + 1
                        gAlm.Dataset.Edit
                        gAlm.Columns.ColumnByFieldName("Item").Value = i
                        gAlm.Dataset.Post
                        gAlm.Dataset.Next
                    Loop
                    If gAlm.Dataset.State = dsEdit Or gAlm.Dataset.State = dsInsert Then
                        gAlm.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gAlm.Dataset.State = dsEdit Or gAlm.Dataset.State = dsInsert Then
            gAlm.Dataset.Post
        End If
    End If

End Sub

Private Sub gAlm_OnKeyPress(Key As Integer)
Dim strCod As String
Dim strDes As String

    Select Case gAlm.Columns.FocusedColumn.Index
        Case gAlm.Columns.ColumnByFieldName("idSucursal").Index
            strCod = gAlm.Columns.ColumnByFieldName("idSucursal").Value
            strDes = gAlm.Columns.ColumnByFieldName("GlsSucursal").Value
            mostrarAyudaKeyasciiTexto Key, "SUCURSAL", strCod, strDes
            Key = 0
            If existeEnGrilla(gAlm, "idSucursal", strCod) = False Then
                gAlm.Dataset.Edit
                gAlm.Columns.ColumnByFieldName("idSucursal").Value = strCod
                gAlm.Columns.ColumnByFieldName("GlsSucursal").Value = strDes
                gAlm.Dataset.Post
                gAlm.SetFocus
            Else
                MsgBox "La sucursal ya fue ingresada", vbInformation, App.Title
            End If
        
        Case gAlm.Columns.ColumnByFieldName("idAlmacen").Index
            strCod = gAlm.Columns.ColumnByFieldName("idAlmacen").Value
            strDes = gAlm.Columns.ColumnByFieldName("GlsAlmacen").Value
            mostrarAyudaKeyasciiTexto Key, "ALMACENVTA", strCod, strDes, " AND idSucursal = '" & gAlm.Columns.ColumnByFieldName("idSucursal").Value & "'"
            Key = 0
            gAlm.Dataset.Edit
            gAlm.Columns.ColumnByFieldName("idAlmacen").Value = strCod
            gAlm.Columns.ColumnByFieldName("GlsAlmacen").Value = strDes
            gAlm.Dataset.Post
            gAlm.SetFocus
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
    If gAlm.Count >= 1 Then
        If gAlm.Count = 1 And gAlm.Columns.ColumnByFieldName("idSucursal").Value = "" And gAlm.Columns.ColumnByFieldName("idAlmacen").Value = "" Then
            StrMsgError = "Falta Ingresar Datos"
            GoTo Err
        End If
    End If
    gAlm.Dataset.First
    
    Cn.BeginTrans
    indIniTrans = True
    
    Cn.Execute "DELETE FROM AlmacenesVtas WHERE idEmpresa = '" & glsEmpresa & "'"
    Do While Not gAlm.Dataset.EOF
        Cn.Execute "INSERT INTO AlmacenesVtas (idSucursal,idAlmacen,item,idEmpresa) VALUES(" & _
                   "'" & gAlm.Columns.ColumnByFieldName("idSucursal").Value & "','" & gAlm.Columns.ColumnByFieldName("idAlmacen").Value & "'," & gAlm.Columns.ColumnByFieldName("item").Value & ",'" & glsEmpresa & "')"
        gAlm.Dataset.Next
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
        If gAlm.Count >= 1 Then
            gAlm.Dataset.First
            indEntro = False
            Do While Not gAlm.Dataset.EOF
                If Trim(gAlm.Columns.ColumnByFieldName("idSucursal").Value) = "" Or Trim(gAlm.Columns.ColumnByFieldName("idAlmacen").Value) = "" Then
                    gAlm.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gAlm.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gAlm.Count >= 1 Then
        gAlm.Dataset.First
        i = 0
        Do While Not gAlm.Dataset.EOF
            i = i + 1
            gAlm.Dataset.Edit
            gAlm.Columns.ColumnByFieldName("item").Value = i
            If gAlm.Dataset.State = dsEdit Then gAlm.Dataset.Post
            gAlm.Dataset.Next
        Loop
    Else
        indInserta = True
        gAlm.Dataset.Append
        indInserta = False
    End If
    
End Sub
