VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmListaDocExportarGuia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar Guia"
   ClientHeight    =   8880
   ClientLeft      =   1140
   ClientTop       =   870
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   1164
      ButtonWidth     =   1455
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refrescar"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Listado"
      ForeColor       =   &H00C00000&
      Height          =   8115
      Left            =   0
      TabIndex        =   0
      Top             =   675
      Width           =   12510
      Begin VB.CommandButton cmbAyudaMotivoTraslado 
         Height          =   315
         Left            =   7350
         Picture         =   "frmListaDocExportarGuia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   390
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3420
         Left            =   75
         OleObjectBlob   =   "frmListaDocExportarGuia.frx":038A
         TabIndex        =   1
         Top             =   900
         Width           =   12315
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   3375
         Left            =   75
         OleObjectBlob   =   "frmListaDocExportarGuia.frx":3C3B
         TabIndex        =   2
         Top             =   4635
         Width           =   12315
      End
      Begin CATControls.CATTextBox txtCod_MotivoTraslado 
         Height          =   285
         Left            =   900
         TabIndex        =   5
         Tag             =   "TidMotivoTraslado"
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
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
         Locked          =   -1  'True
         MaxLength       =   8
         Container       =   "frmListaDocExportarGuia.frx":6B5E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_MotivoTraslado 
         Height          =   285
         Left            =   1875
         TabIndex        =   6
         Top             =   300
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   503
         BackColor       =   16775664
         Enabled         =   0   'False
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
         Container       =   "frmListaDocExportarGuia.frx":6B7A
         Vacio           =   -1  'True
      End
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
            NumListImages   =   16
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":6B96
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":6F30
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":7382
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":771C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":7AB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":7E50
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":81EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":8584
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":891E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":8CB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":9052
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":9D14
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":A0AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":A500
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":A89A
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaDocExportarGuia.frx":B2AC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl_MotivoTraslado 
         Appearance      =   0  'Flat
         Caption         =   "Motivo:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   150
         TabIndex        =   7
         Top             =   375
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmListaDocExportarGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsg As New ADODB.Recordset
Private rsd As New ADODB.Recordset

Private strTDExportar As String

Dim indNuevoDoc As Boolean

Private Sub cmbAyudaMotivoTraslado_Click()
    mostrarAyuda "MOTIVOTRASLADO", txtCod_MotivoTraslado, txtGls_MotivoTraslado, " AND idMotivoTraslado IN ('06090006','08050001')"
    If txtCod_MotivoTraslado.Text <> "" Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
Dim strMsgError As String
On Error GoTo ERR

strRptNum = ""
strRptSerie = ""

ConfGrid gLista, True, False, False, False
ConfGrid gListaDetalle, True, False, False, False

Exit Sub
ERR:
If strMsgError = "" Then strMsgError = ERR.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub listaDocVentas(ByRef strMsgError As String)
Dim rst As New ADODB.Recordset
Dim strCond As String
On Error GoTo ERR

 '********FORMATO GRILLA
    Set gLista.DataSource = Nothing
 
    If rsg.State = 1 Then rsg.Close
    Set rsg = Nothing
    
    If rsd.State = 1 Then rsd.Close
    Set rsd = Nothing
    
    'Formato cabecera****************************************************
    rsg.Fields.Append "Item", adChar, 13, adFldRowID
    rsg.Fields.Append "chkMarca", adChar, 1, adFldIsNullable
    rsg.Fields.Append "idDocVentas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerie", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idPerVendedor", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsVendedor", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "FecEmision", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "estDocVentas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idSucursal", adChar, 8, adFldIsNullable
    
    rsg.Open , , adOpenKeyset, adLockOptimistic
    
    '********************************************************************
    
    
    'Formato Detalle****************************************************
    rsd.Fields.Append "Item", adVarChar, 20, adFldRowID
    rsd.Fields.Append "chkMarca", adChar, 1, adFldIsNullable
    rsd.Fields.Append "idProducto", adChar, 8, adFldIsNullable
    
    rsd.Fields.Append "idCodFabricante", adVarChar, 20, adFldIsNullable
    rsd.Fields.Append "GlsProducto", adVarChar, 185, adFldIsNullable
    rsd.Fields.Append "idMarca", adChar, 8, adFldIsNullable
    rsd.Fields.Append "GlsMarca", adVarChar, 185, adFldIsNullable
    rsd.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsd.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsd.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsd.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable

    
    rsd.Fields.Append "idDocVentas", adChar, 8, adFldIsNullable
    rsd.Fields.Append "idSerie", adChar, 3, adFldIsNullable
    
    rsd.Fields.Append "NumLote", adVarChar, 30, adFldIsNullable
    rsd.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    
    rsd.Open , , adOpenKeyset, adLockOptimistic
    '********************************************************************
    
    If Trim(txtCod_MotivoTraslado.Text) = "" Then
        
        mostrarDatosGridSQL gLista, rsg, strMsgError
        If strMsgError <> "" Then GoTo ERR
        
        Exit Sub
        
    End If
    
    strCond = ""
    
    csql = "SELECT concat(idDocumento,idDocVentas,idSerie) as Item ,idSucursal, idDocVentas,idSerie,idPerVendedor,GlsVendedor,FecEmision,estDocVentas,idMoneda,TotalValorVenta,TotalIGVVenta,TotalPrecioVenta " & _
            "FROM docventas " & _
            "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal <> '" & glsSucursal & "' AND idDocumento = '86' AND estDocventas <> 'ANU' AND idPerCliente = '" & glsSucursal & "' AND idMotivoTraslado = '" & txtCod_MotivoTraslado.Text & "' AND estGuiaImportado <> 'S'"
            
            
    '''cn    strcn
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF And Not rst.BOF Then
        Do While Not rst.EOF
            
            rsg.AddNew
            
            rsg.Fields("Item") = rst.Fields("Item")
            rsg.Fields("chkMarca") = 0
            rsg.Fields("idDocVentas") = rst.Fields("idDocVentas")
            rsg.Fields("idSerie") = rst.Fields("idSerie")
            rsg.Fields("idPerVendedor") = rst.Fields("idPerVendedor")
            rsg.Fields("GlsVendedor") = rst.Fields("GlsVendedor")
            rsg.Fields("FecEmision") = Format(rst.Fields("FecEmision"), "dd/mm/yyyy")
            rsg.Fields("estDocVentas") = rst.Fields("estDocVentas")
            rsg.Fields("idSucursal") = rst.Fields("idSucursal")

        
            rst.MoveNext
        Loop

        mostrarDatosGridSQL gLista, rsg, strMsgError
        If strMsgError <> "" Then GoTo ERR
    End If
    
Me.Refresh
If rst.State = 1 Then rst.Close
Set rst = Nothing
Exit Sub
ERR:
If rst.State = 1 Then rst.Close
Set rst = Nothing
If strMsgError = "" Then strMsgError = ERR.Description
End Sub

Private Sub listaDetalle(ByRef strMsgError As String)
Dim rst As New ADODB.Recordset
Dim strCond As String
Dim indExisteDoc As Boolean

Dim strNumDoc As String
Dim strSerie As String
Dim stridSucursal As String

On Error GoTo ERR

 '********FORMATO GRILLA
    
    strNumDoc = gLista.Columns.ColumnByFieldName("idDocVentas").Value
    strSerie = gLista.Columns.ColumnByFieldName("idSerie").Value
    stridSucursal = gLista.Columns.ColumnByFieldName("idSucursal").Value
    
    'Validamos si ya adicionamos el detalle******************************
    gListaDetalle.Dataset.Filter = ""
    gListaDetalle.Dataset.Filtered = True
    indExisteDoc = False
    
    Set gListaDetalle.DataSource = Nothing
    
    gListaDetalle.Dataset.DisableControls
    
    If rsd.RecordCount > 0 Then rsd.MoveFirst
    Do While Not rsd.EOF
        If rsd.Fields("idDocVentas") = strNumDoc And rsd.Fields("idSerie") = strSerie Then
            indExisteDoc = True
            Exit Do
        End If
        rsd.MoveNext
    Loop
    '********************************************************************
    
    If indExisteDoc = False Then
        strCond = ""
'        If Trim(txt_TextoBuscar.Text) <> "" Then
'            strCond = Trim(txt_TextoBuscar.Text)
'            strCond = " AND GlsCliente LIKE '%" & strCond & "%'"
'        End If
        
        csql = "SELECT item, idProducto,idCodFabricante, GlsProducto,idMarca, GlsMarca,idUM, GlsUM,Factor,Afecto, Cantidad,VVUnit,IGVUnit, PVUnit,TotalVVBruto,TotalPVBruto, PorDcto,DctoVV,DctoPV,TotalVVNeto,TotalIGVNeto, TotalPVNeto,idTipoProducto,idMoneda,NumLote,FecVencProd FROM docventasdet " & _
                "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & stridSucursal & "' AND idDocumento = '86' AND idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
                
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
        Do While Not rst.EOF
    
            rsd.AddNew
            
            rsd.Fields("Item") = strNumDoc & strSerie & CStr(rst.Fields("Item"))
            rsd.Fields("chkMarca") = 0
            rsd.Fields("idProducto") = rst.Fields("idProducto")
            
            rsd.Fields("idCodFabricante") = "" & rst.Fields("idCodFabricante")
            rsd.Fields("GlsProducto") = "" & rst.Fields("GlsProducto")
            rsd.Fields("idMarca") = "" & rst.Fields("idMarca")
            rsd.Fields("GlsMarca") = "" & rst.Fields("GlsMarca")
            rsd.Fields("idUM") = "" & rst.Fields("idUM")
            rsd.Fields("GlsUM") = "" & rst.Fields("GlsUM")
            rsd.Fields("Factor") = "" & rst.Fields("Factor")
            rsd.Fields("Afecto") = "" & rst.Fields("Afecto")
            rsd.Fields("Cantidad") = "" & rst.Fields("Cantidad")
            rsd.Fields("idDocVentas") = strNumDoc
            rsd.Fields("idSerie") = strSerie
            
            rsd.Fields("NumLote") = "" & rst.Fields("NumLote")
            rsd.Fields("FecVencProd") = "" & rst.Fields("FecVencProd")
            
            rst.MoveNext
        Loop
    End If
    
    If rsd.RecordCount > 0 Then
        mostrarDatosGridSQL gListaDetalle, rsd, strMsgError
        If strMsgError <> "" Then GoTo ERR
    End If
    
    gListaDetalle.Dataset.Filter = " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    gListaDetalle.Dataset.Filtered = True
    gListaDetalle.Dataset.EnableControls
    
Me.Refresh
If rst.State = 1 Then rst.Close
Set rst = Nothing
Exit Sub
ERR:
If rst.State = 1 Then rst.Close
Set rst = Nothing
gListaDetalle.Dataset.EnableControls
If strMsgError = "" Then strMsgError = ERR.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

If rsg.State = 1 Then rsg.Close
Set rsg = Nothing

If rsd.State = 1 Then rsd.Close
Set rsd = Nothing

End Sub

Private Sub gLista_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    If gLista.Dataset.State = dsEdit Then
        gLista.Dataset.Post
    End If
    
    If gListaDetalle.Count = 0 Then Exit Sub
    
    gListaDetalle.Dataset.First
    
    Do While Not gListaDetalle.Dataset.EOF
        gListaDetalle.Dataset.Edit
        gListaDetalle.Columns.ColumnByFieldName("chkMarca").Value = Column.Value
        gListaDetalle.Dataset.Post
        
        gListaDetalle.Dataset.Next
    Loop
End Sub

Private Sub gListaDetalle_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    If gListaDetalle.Dataset.State = dsEdit Then
        gListaDetalle.Dataset.Post
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strMsgError As String
On Error GoTo ERR
Select Case Button.Index
    Case 1 'Nuevo
        If gLista.Count > 0 Then
            Me.Hide
        End If
    Case 2
        listaDocVentas strMsgError
        If strMsgError <> "" Then GoTo ERR
    Case 4
        strRptNum = ""
        strRptSerie = ""
        Me.Hide
End Select
Exit Sub
ERR:
If strMsgError = "" Then strMsgError = ERR.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim strMsgError As String

On Error GoTo ERR

    listaDetalle strMsgError
    If strMsgError <> "" Then GoTo ERR
    
Exit Sub
ERR:
If strMsgError = "" Then strMsgError = ERR.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnDblClick()
Dim strMsgError As String
On Error GoTo ERR

If gLista.Count > 0 Then
            
    Me.Hide

End If

Exit Sub
ERR:
MsgBox strMsgError, vbInformation, App.Title
End Sub

Public Sub mostrarForm(ByRef rscd As ADODB.Recordset, ByRef rsdd As ADODB.Recordset, ByRef strNumDocImportado As String, ByRef strMotivo As String, ByRef strMsgError As String)

On Error GoTo ERR

    indNuevoDoc = True
    
    Set gLista.DataSource = Nothing
    Set gListaDetalle.DataSource = Nothing
    
    strTDExportar = "86"
    
    txtCod_MotivoTraslado.Text = "06090006"
    
    indNuevoDoc = False
   
    listaDocVentas strMsgError
    If strMsgError <> "" Then GoTo ERR
    
    frmListaDocExportarGuia.Show 1
    
    'Quitamos Filtros existentes
    gLista.Dataset.Filter = ""
    gLista.Dataset.Filtered = True
    
    gListaDetalle.Dataset.Filter = ""
    gListaDetalle.Dataset.Filtered = True
    
    Set gLista.DataSource = Nothing
    Set gListaDetalle.DataSource = Nothing
    
    If TypeName(rsg) = "Nothing" Then
        Exit Sub
    Else
        If rsg.State = 0 Then
            Exit Sub
        End If
        
        If rsg.EOF Or rsg.BOF Then
            Exit Sub
        End If
    End If
    
    'Eliminamos los registros q no estan marcados
    If rsg.RecordCount > 0 Then
        rsg.MoveFirst
        Do While Not rsg.EOF
            If rsg.Fields("chkMarca") = "0" Or IsNull(rsg.Fields("chkMarca")) = True Then
                rsg.Delete adAffectCurrent
                rsg.Update
            End If
            rsg.MoveNext
        Loop
    End If
    
    If rsd.RecordCount > 0 Then
        rsd.MoveFirst
        Do While Not rsd.EOF
            If rsd.Fields("chkMarca") = "0" Then
                rsd.Delete adAffectCurrent
                rsd.Update
            End If
            rsd.MoveNext
        Loop
    End If
        
    'Devolvemos valores seleccionados
    strMotivo = ""
    If rsg.RecordCount > 0 Then
        strNumDocImportado = "S"
        strMotivo = Trim(txtCod_MotivoTraslado.Text)
    End If
       
    Set rscd = rsg.Clone(adLockReadOnly)
    Set rsdd = rsd.Clone(adLockReadOnly)
    
    
    If rsg.State = 1 Then rsg.Close
    Set rsg = Nothing
    
    If rsd.State = 1 Then rsd.Close
    Set rsd = Nothing
    
    Unload Me

Exit Sub
ERR:
If strMsgError = "" Then strMsgError = ERR.Description
End Sub

Private Sub txtCod_MotivoTraslado_Change()
Dim strMsgError As String

On Error GoTo ERR

    txtGls_MotivoTraslado.Text = traerCampo("motivostraslados", "GlsMotivoTraslado", "idMotivoTraslado", txtCod_MotivoTraslado.Text, False)
    
    If indNuevoDoc Then Exit Sub
    
    listaDocVentas strMsgError
    If strMsgError <> "" Then GoTo ERR
    
Exit Sub
ERR:
If strMsgError = "" Then strMsgError = ERR.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_MotivoTraslado_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    mostrarAyudaKeyascii KeyAscii, "MOTIVOTRASLADO", txtCod_MotivoTraslado, txtGls_MotivoTraslado, " AND idMotivoTraslado IN ('06090006','08050001')"
    KeyAscii = 0
    If txtCod_MotivoTraslado.Text <> "" Then SendKeys "{tab}"
End If
End Sub
