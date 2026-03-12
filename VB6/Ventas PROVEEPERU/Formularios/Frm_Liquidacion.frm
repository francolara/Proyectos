VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Frm_Liquidacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidaciones"
   ClientHeight    =   4350
   ClientLeft      =   4035
   ClientTop       =   2640
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDetalle 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   3645
      Left            =   30
      TabIndex        =   1
      Top             =   690
      Width           =   9420
      Begin DXDBGRIDLibCtl.dxDBGrid gLiquidaciones 
         Height          =   3375
         Left            =   45
         OleObjectBlob   =   "Frm_Liquidacion.frx":0000
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   135
         Width           =   9345
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   3735
      Top             =   90
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
            Picture         =   "Frm_Liquidacion.frx":2F45
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":32DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":3731
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":3ACB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":3E65
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":41FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":4599
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":4933
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":4CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":5067
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":5401
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Liquidacion.frx":60C3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
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
            Caption         =   "Eiminar"
            Key             =   "Eiminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "Frm_Liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strserieguia As String
Dim strnumguia As String

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLiquidaciones, True, False, False, False
    mostrarLiquidaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    gLiquidaciones.Columns.ColumnByFieldName("IdLiquidacion").DisableEditor = False
     
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLiquidaciones_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
     
     If Action = daInsert Then
        gLiquidaciones.Columns.ColumnByFieldName("item").Value = gLiquidaciones.Count
        gLiquidaciones.Columns.ColumnByFieldName("IdLiquidacion").Value = ""
        gLiquidaciones.Columns.ColumnByFieldName("idGranjaOri").Value = ""
        gLiquidaciones.Columns.ColumnByFieldName("idCamal").Value = ""
        gLiquidaciones.Columns.ColumnByFieldName("FechaLiq").Value = 0
    End If

End Sub

Private Sub gLiquidaciones_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String
Dim rsRpt       As New ADODB.Recordset
Dim indEvaluaEstado     As Boolean

    indEvaluaEstado = False
    Select Case Column.Index
        Case gLiquidaciones.Columns.ColumnByFieldName("IdLiquidacion").Index
            If gLiquidaciones.Columns.ColumnByFieldName("IdLiquidacion").ReadOnly = False Then
                Frm_Ayuda_Liquidacion.MostrarForm rsRpt, indEvaluaEstado, StrMsgError
                If StrMsgError <> "" Then GoTo Err
                 If indEvaluaEstado = True Then
                    procesaDocumentos rsRpt, StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                 End If
            End If
    End Select

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLiquidaciones_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim i As Integer
    
    If KeyCode = 46 Then
        If gLiquidaciones.Count > 0 Then
            If MsgBox("¿Seguro de eliminar el registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
                                               
                If gLiquidaciones.Count = 1 Then
                    gLiquidaciones.Dataset.Edit
                    gLiquidaciones.Columns.ColumnByFieldName("item").Value = gLiquidaciones.Count
                    gLiquidaciones.Columns.ColumnByFieldName("IdLiquidacion").Value = ""
                    gLiquidaciones.Columns.ColumnByFieldName("idGranjaOri").Value = ""
                    gLiquidaciones.Columns.ColumnByFieldName("idCamal").Value = ""
                    gLiquidaciones.Columns.ColumnByFieldName("FechaLiq").Value = ""
                    gLiquidaciones.Columns.ColumnByFieldName("glsCamal").Value = ""
                    gLiquidaciones.Columns.ColumnByFieldName("glsGranjaOri").Value = ""
                    
                    gLiquidaciones.Dataset.Post
                Else
                    gLiquidaciones.Dataset.Delete
                    gLiquidaciones.Dataset.First
                    Do While Not gLiquidaciones.Dataset.EOF
                        i = i + 1
                        gLiquidaciones.Dataset.Edit
                        gLiquidaciones.Columns.ColumnByFieldName("Item").Value = i
                        gLiquidaciones.Dataset.Post
                        gLiquidaciones.Dataset.Next
                    Loop
                    If gLiquidaciones.Dataset.State = dsEdit Or gLiquidaciones.Dataset.State = dsInsert Then
                        gLiquidaciones.Dataset.Post
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = 13 Then
        If gLiquidaciones.Dataset.State = dsEdit Or gLiquidaciones.Dataset.State = dsInsert Then
              gLiquidaciones.Dataset.Post
        End If
    End If

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
        Case 3 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 4 'Salir
            Unload Me
    End Select
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub eliminaNulosGrilla()
Dim indWhile As Boolean
Dim indEntro As Boolean
Dim i As Integer
    
    indWhile = True
    Do While indWhile = True
        If gLiquidaciones.Count >= 1 Then
            gLiquidaciones.Dataset.First
            indEntro = False
            Do While Not gLiquidaciones.Dataset.EOF
                If Trim(gLiquidaciones.Columns.ColumnByFieldName("IdLiquidacion").Value) = "" Then
                    gLiquidaciones.Dataset.Delete
                    indEntro = True
                    Exit Do
                End If
                gLiquidaciones.Dataset.Next
            Loop
            indWhile = indEntro
        Else
            indWhile = False
        End If
    Loop
    
    If gLiquidaciones.Count >= 1 Then
        gLiquidaciones.Dataset.First
        i = 0
        Do While Not gLiquidaciones.Dataset.EOF
            i = i + 1
            gLiquidaciones.Dataset.Edit
            gLiquidaciones.Columns.ColumnByFieldName("item").Value = i
            If gLiquidaciones.Dataset.State = dsEdit Then gLiquidaciones.Dataset.Post
            gLiquidaciones.Dataset.Next
        Loop
    Else
        indInserta = True
        gLiquidaciones.Dataset.Append
        indInserta = False
    End If
    
End Sub

Private Sub procesaDocumentos(ByVal rsDoc As ADODB.Recordset, ByRef StrMsgError As String)
On Error GoTo Err
Dim i As Integer
Dim rsDetalle As New ADODB.Recordset

    rsDetalle.Fields.Append "Item", adInteger, , adFldRowID
    rsDetalle.Fields.Append "IdLiquidacion", adChar, 20, adFldIsNullable
    rsDetalle.Fields.Append "idGranjaOri", adVarChar, 50, adFldIsNullable
    rsDetalle.Fields.Append "idCamal", adVarChar, 10, adFldIsNullable
    rsDetalle.Fields.Append "FechaLiq", adVarChar, 14, adFldIsNullable
    rsDetalle.Fields.Append "glsGranjaOri", adVarChar, 200, adFldIsNullable
    rsDetalle.Fields.Append "glsCamal", adVarChar, 200, adFldIsNullable
    rsDetalle.Open
    
    gLiquidaciones.Dataset.First
    Do While Not gLiquidaciones.Dataset.EOF
       If Trim(gLiquidaciones.Columns.ColumnByFieldName("IdLiquidacion").Value) <> "" Then
            i = i + 1
            rsDetalle.AddNew
            rsDetalle.Fields("Item") = i
            rsDetalle.Fields("IdLiquidacion") = Trim(gLiquidaciones.Columns.ColumnByFieldName("IdLiquidacion").Value)
            rsDetalle.Fields("idGranjaOri") = Trim(gLiquidaciones.Columns.ColumnByFieldName("idGranjaOri").Value)
            rsDetalle.Fields("idCamal") = Trim(gLiquidaciones.Columns.ColumnByFieldName("idCamal").Value)
            rsDetalle.Fields("FechaLiq") = Trim(gLiquidaciones.Columns.ColumnByFieldName("FechaLiq").Value)
            rsDetalle.Fields("glsGranjaOri") = Trim(gLiquidaciones.Columns.ColumnByFieldName("glsGranjaOri").Value)
            rsDetalle.Fields("glsCamal") = Trim(gLiquidaciones.Columns.ColumnByFieldName("glsCamal").Value)
        End If
        gLiquidaciones.Dataset.Next
    Loop
    
    rsDoc.MoveFirst
    Do While Not rsDoc.EOF
        If Trim("" & rsDoc("CHK").Value) = "S" Then
            i = i + 1
            rsDetalle.AddNew
            rsDetalle.Fields("Item") = i
            rsDetalle.Fields("IdLiquidacion") = Trim(rsDoc("IdLiquidacion").Value & "")
            rsDetalle.Fields("idGranjaOri") = Trim(rsDoc("IdUPP") & "")
            rsDetalle.Fields("idCamal") = Trim(rsDoc("idCamal") & "")
            rsDetalle.Fields("FechaLiq") = Trim(rsDoc("FechaLiq") & "")
            rsDetalle.Fields("glsGranjaOri") = Trim(rsDoc("glsGranjaOri") & "")
            rsDetalle.Fields("glsCamal") = Trim(rsDoc("glsCamal") & "")
        End If
        rsDoc.MoveNext
    Loop
    
    Set rs_planilla_det = rsDetalle
    mostrarDatosGridSQL gLiquidaciones, rsDetalle, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo(ByRef StrMsgError As String)
On Error GoTo Err
Dim rsg As New ADODB.Recordset
Dim rsd As New ADODB.Recordset
Dim strAno As String
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "IdLiquidacion", adChar, 20, adFldIsNullable
    rsg.Fields.Append "idGranjaOri", adVarChar, 50, adFldIsNullable
    rsg.Fields.Append "idCamal", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "FechaLiq", adVarChar, 14, adFldIsNullable
    rsg.Fields.Append "glsGranjaOri", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "glsCamal", adVarChar, 14, adFldIsNullable
    rsg.Open
    
    rsg.AddNew
    rsg.Fields("Item") = 1
    rsg.Fields("IdLiquidacion") = ""
    rsg.Fields("idGranjaOri") = ""
    rsg.Fields("idCamal") = ""
    rsg.Fields("FechaLiq") = ""
    rsg.Fields("glsGranjaOri") = ""
    rsg.Fields("glsCamal") = ""
    
    Set gLiquidaciones.DataSource = Nothing
    mostrarDatosGridSQL gLiquidaciones, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
   
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
  
    If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    
    csql = "Delete  From docventasdetliquidacion Where IdLiquidacion ='" & Trim(gLiquidaciones.Columns.ColumnByFieldName("IdLiquidacion").Value) & "'  And idEmpresa = '" & glsEmpresa & "' And idSucursal ='" & glsSucursal & "'AND NumGuia ='" & strnumguia & "' AND   SerieGuia ='" & strserieguia & "' "
    Cn.Execute (csql)
    
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    
    mostrarLiquidaciones StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
      
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim item As Integer

    item = 0
    item = traerCampo("docventasdetliquidacion", "ifnull(max(item),0)", "idEmpresa", glsEmpresa, True) + 1
    
    With gLiquidaciones
        If Len(Trim(.Columns.ColumnByFieldName("IdLiquidacion").Value)) > 0 Then
            .Dataset.First
            Do While Not .Dataset.EOF
                csql = "Delete  From  docventasdetliquidacion Where  IdLiquidacion ='" & Trim(.Columns.ColumnByFieldName("IdLiquidacion").Value) & "' And idEmpresa=  '" & glsEmpresa & "' And idSucursal = '" & glsSucursal & "' And SerieGuia= '" & strserieguia & "' And  NumGuia= '" & strnumguia & "'"
                Cn.Execute (csql)
           
                csql = "Insert Into docventasdetliquidacion(IdEmpresa, IdSucursal, SerieGuia, NumGuia, IdLiquidacion, idCamal, idGranjaOri, " & _
                        "FechaLiq, FecRegistro,HoraRegistro,IdUsuarioRegistro,Item) " & _
                        "Values('" & glsEmpresa & "','" & glsSucursal & "','" & strserieguia & "','" & strnumguia & "','" & Trim(.Columns.ColumnByFieldName("IdLiquidacion").Value) & "','" & Trim(.Columns.ColumnByFieldName("idCamal").Value) & "', " & _
                        "'" & Trim(.Columns.ColumnByFieldName("idGranjaOri").Value) & "' ,'" & Format(Trim(.Columns.ColumnByFieldName("FechaLiq").Value), "yyyy-mm-dd") & "','" & Format(getFechaSistema, "yyyy-mm-dd") & "', '" & Time & "' ,'" & glsUser & "','" & item & "'  ) "
                Cn.Execute (csql)
              
                .Dataset.Next
                item = item + 1
            Loop
            MsgBox "Se Grabo Satisfactoriamente", vbInformation, App.Title
        
        Else
            MsgBox "No hay Registros", vbInformation, App.Title
        End If
    End With
        
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub MostrarForm(ByVal strVarSerie As String, ByVal strVarNum As String, ByRef StrMsgError As String)
On Error GoTo Err

    strserieguia = strVarSerie
    strnumguia = strVarNum
    Me.Show 1
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarLiquidaciones(ByVal StrMsgError As String)
On Error GoTo Err
Dim rsg As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim item As Integer

    item = 0
    csql = "Select d.IdLiquidacion,d.idGranjaOri,d.idCamal,p.glsPersona as glsCamal ,u.DescUnidad  as glsGranjaOri , d.FechaLiq From docventasdetliquidacion d " & _
           "Inner Join UnidadProduccion u    On d.idGranjaOri = u.CodUnidProd  And d.idEmpresa = u.idEmpresa " & _
           "Inner Join Personas p   On   p.idPersona = d.idCamal " & _
           "Where d.NumGuia ='" & strnumguia & "' AND   d.SerieGuia ='" & strserieguia & "'  AND d.idEmpresa = '" & glsEmpresa & "' And d.idSucursal = '" & glsSucursal & "' "
    
    If rst.State = 1 Then rst.Close
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "IdLiquidacion", adChar, 20, adFldIsNullable
    rsg.Fields.Append "idGranjaOri", adVarChar, 50, adFldIsNullable
    rsg.Fields.Append "idCamal", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "FechaLiq", adVarChar, 14, adFldIsNullable
    rsg.Fields.Append "glsGranjaOri", adVarChar, 200, adFldIsNullable
    rsg.Fields.Append "glsCamal", adVarChar, 200, adFldIsNullable
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("IdLiquidacion") = ""
        rsg.Fields("idGranjaOri") = ""
        rsg.Fields("idCamal") = ""
        rsg.Fields("FechaLiq") = ""
        rsg.Fields("glsGranjaOri") = ""
        rsg.Fields("glsCamal") = ""
    Else
        If Not rst.EOF Then
          rst.MoveFirst
            Do While Not rst.EOF
                rsg.AddNew
                rsg.Fields("Item") = item
                rsg.Fields("IdLiquidacion") = Trim("" & rst.Fields("IdLiquidacion"))
                rsg.Fields("idGranjaOri") = "" & rst.Fields("idGranjaOri")
                rsg.Fields("idCamal") = "" & rst.Fields("idCamal")
                rsg.Fields("FechaLiq") = "" & Format(rst.Fields("FechaLiq"), "dd/mm/yyyy")
                rsg.Fields("glsGranjaOri") = "" & rst.Fields("glsGranjaOri")
                rsg.Fields("glsCamal") = "" & rst.Fields("glsCamal")
                rst.MoveNext
                item = item + 1
            Loop
        End If
    End If
    rst.Close: Set rst = Nothing
    
    mostrarDatosGridSQL gLiquidaciones, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
