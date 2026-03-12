VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmAyudaLiquidacion 
   Caption         =   "Ayuda de Liquidaciones"
   ClientHeight    =   5235
   ClientLeft      =   2265
   ClientTop       =   3270
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   14085
   Begin VB.Frame Frame2 
      Height          =   4515
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   14010
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   3375
         Left            =   45
         OleObjectBlob   =   "FrmAyudaLiquidacion.frx":0000
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   585
         Width           =   13845
      End
      Begin CATControls.CATTextBox txt_CantSolictada 
         Height          =   285
         Left            =   9585
         TabIndex        =   6
         Top             =   4095
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "FrmAyudaLiquidacion.frx":6669
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_PesoSolicitado 
         Height          =   285
         Left            =   12600
         TabIndex        =   7
         Top             =   4095
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "FrmAyudaLiquidacion.frx":6685
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Peso Solicitado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10755
         TabIndex        =   5
         Top             =   4095
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad Solicitada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7245
         TabIndex        =   4
         Top             =   4095
         Width           =   2355
      End
      Begin VB.Label lblGlsProducto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Descripcion del Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   30
         TabIndex        =   3
         Top             =   225
         Width           =   13830
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   60
      Top             =   1020
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
            Picture         =   "FrmAyudaLiquidacion.frx":66A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":6A3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":6E8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":7227
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":75C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":795B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":7CF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":808F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":8429
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":87C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":8B5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":981F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":9BB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":A00B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":A3A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLiquidacion.frx":ADB7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   1164
      ButtonWidth     =   2064
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Grabar      "
            Object.ToolTipText     =   "Procesar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            ImageIndex      =   6
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
Attribute VB_Name = "FrmAyudaLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw_limpia            As Boolean
Dim codigo               As String
Dim descripcion          As String
Dim CantProDocVenta      As String
Dim CodDocumento         As String
Dim CodSerie             As String
Dim CodDocventas         As String
Dim CodGlsProducto       As String
Dim rsg                  As New ADODB.Recordset
Dim CodProductoDocVentas As String
Dim ItemDocVentas        As String
Dim indEstLiquix         As Integer

Private Sub mostrarLiquidaciones(unidadprod As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim item As Integer

    item = 0
    csql = "Select Item, IdLiquidacion, IdUPP, idCamal, FecRegistro, Glspersona, DescUnidad, ValPesoVivo, CantidadAnt, PesoAnt,RePesoAnt,KgEnvAnt, KgOri, ValPeso, UnidadOri, ValCantidad " & _
           "From(" & _
           "Select dd.Item,d.IdLiquidacion,dd.IdUPP,d.idCamal,FecRegistro,p.DescUnidad as Glspersona,u.DescUnidad,d.ValPesoVivo, " & _
           "ddl.Unidad as CantidadAnt,ddl.kg as PesoAnt,ddl.ReKg as RePesoAnt,ddl.KgEnv KgEnvAnt, " & _
           "d.ValPeso as KgOri  ,(d.ValPeso - PesoImp) as ValPeso,sum(dd.ValCantidad) as UnidadOri,(sum(dd.ValCantidad)- CantidadImp) as ValCantidad From docventasliqcab d " & _
           "Inner Join  docventasliqdet dd  " & _
           "On d.IdLiquidacion = dd.IdLiquidacion And d.IdEmpresa = dd.IdEmpresa And d.IdSucursal = dd.IdSucursal " & _
           "Inner Join UnidadProduccion u  On dd.IdUPP = u.CodUnidProd And dd.idEmpresa = u.idEmpresa " & _
           "Inner Join UnidadProduccion p On d.idCamal = p.CodUnidProd And d.idEmpresa = p.idEmpresa " & _
           "left join (Select IdDocumento, IdSerie, IdDocventas, IdProducto, IdLiquidacion, IdUPP, idCamal, FechaLiq, UnidadSaldo, Unidad, KgSaldo, Kg, KgVivo, IdEmpresa, Idsucursal, Item,ReKg,KgEnv " & _
           "from docventasdetliquidacion  " & _
           "where IdDocumento = '" & CodDocumento & "' and IdSerie = '" & CodSerie & "' and IdDocventas = '" & CodDocventas & "' and IdProducto = '" & CodProductoDocVentas & "' ) ddl " & _
           "On d.IdLiquidacion = ddl.IdLiquidacion And d.IdEmpresa = ddl.IdEmpresa And d.IdSucursal = ddl.IdSucursal " & _
           "where d.idCamal = '" & unidadprod & "' " & _
           "and d.idempresa = '" & glsEmpresa & "' " & _
           "Group By d.IdLiquidacion " & _
           "Order By d.IdLiquidacion ) x where (valpeso > 0 and valcantidad > 0)"
           
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "IdLiquidacion", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdUPP", adVarChar, 20, adFldIsNullable
    rsg.Fields.Append "idCamal", adVarChar, 50, adFldIsNullable
    rsg.Fields.Append "FechaLiq", adVarChar, 50, adFldIsNullable
    rsg.Fields.Append "glsGranjaOri", adVarChar, 200, adFldIsNullable
    rsg.Fields.Append "glsCamal", adVarChar, 200, adFldIsNullable
    rsg.Fields.Append "UnidadOri", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "UnidadSaldo", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Unidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "KgOri", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "KgSaldo", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Kg", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "KgVivo", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idProducto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "ItemDocVentas", adVarChar, 3, adFldIsNullable
    rsg.Fields.Append "CantidadAnt", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "PesoAnt", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "RePesoAnt", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "ReKg", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "KgEnv", adDouble, 14, adFldIsNullable
    
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("IdLiquidacion") = ""
        rsg.Fields("IdUPP") = ""
        rsg.Fields("idCamal") = ""
        rsg.Fields("FechaLiq") = ""
        rsg.Fields("glsGranjaOri") = ""
        rsg.Fields("glsCamal") = ""
        rsg.Fields("UnidadSaldo") = 0#
        rsg.Fields("Unidad") = 0#
        rsg.Fields("KgSaldo") = 0#
        rsg.Fields("Kg") = 0#
        rsg.Fields("KgVivo") = 0#
        rsg.Fields("idproducto") = ""
        rsg.Fields("ItemDocVentas") = ""
        rsg.Fields("UnidadOri") = 0#
        rsg.Fields("KgOri") = 0#
        rsg.Fields("CantidadAnt") = 0#
        rsg.Fields("PesoAnt") = 0#
        rsg.Fields("RePesoAnt") = 0#
        rsg.Fields("ReKg") = 0#
        rsg.Fields("KgEnv") = 0#
        
        
        
    Else
        Do While Not rst.EOF
         item = item + 1
            rsg.AddNew
            rsg.Fields("Item") = item
            rsg.Fields("IdLiquidacion") = Trim("" & rst.Fields("IdLiquidacion"))
            rsg.Fields("IdUPP") = Trim("" & rst.Fields("IdUPP"))
            rsg.Fields("idCamal") = Trim("" & rst.Fields("idCamal"))
            rsg.Fields("FechaLiq") = Trim("" & Format(rst.Fields("FecRegistro"), "dd/mm/yyyy"))
            rsg.Fields("glsCamal") = Trim("" & rst.Fields("Glspersona"))
            rsg.Fields("glsGranjaOri") = Trim("" & rst.Fields("DescUnidad"))
            rsg.Fields("UnidadOri") = Val(Format(rst.Fields("UnidadOri"), "0.00"))
            rsg.Fields("UnidadSaldo") = Val(Format(rst.Fields("ValCantidad"), "0.00"))
            rsg.Fields("KgOri") = Val(Format(rst.Fields("KgOri"), "0.00"))
            
            rsg.Fields("Unidad") = Val(Format(rst.Fields("CantidadAnt"), "0.00"))
            rsg.Fields("Kg") = Val(Format(rst.Fields("PesoAnt"), "0.00"))
            rsg.Fields("ReKg") = Val(Format(rst.Fields("RePesoAnt"), "0.00"))
            rsg.Fields("KgEnv") = Val(Format(rst.Fields("KgEnvAnt"), "0.00"))
            
            
            rsg.Fields("KgSaldo") = Val(Format(rst.Fields("ValPeso"), "0.00"))
            rsg.Fields("KgVivo") = Val(Format(rst.Fields("ValPesoVivo"), "0.00"))
'            rsg.Fields("Kg") = 0#
'            rsg.Fields("Unidad") = 0#
'            rsg.Fields("ReKg") = 0#
            rsg.Fields("idproducto") = Trim("" & CodProductoDocVentas)
            rsg.Fields("ItemDocVentas") = Trim("" & ItemDocVentas)
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    mostrarDatosGridSQL dxDBGrid1, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError         As String
Dim sw_nc               As Boolean
Dim DblDifRepeso        As Double

    Select Case dxDBGrid1.Columns.FocusedColumn.Index
    
'        Case dxDBGrid1.Columns.ColumnByFieldName("KgEnv").Index
'            dxDBGrid1.Dataset.Edit
'            dxDBGrid1.Columns.ColumnByFieldName("kg").Value = 0#
'            dxDBGrid1.Columns.ColumnByFieldName("Rekg").Value = 0#
'            dxDBGrid1.Dataset.Post
            
        Case dxDBGrid1.Columns.ColumnByFieldName("Unidad").Index
        
            If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("Unidad").Value, "0.00")) > Val(Format(dxDBGrid1.Columns.ColumnByFieldName("UnidadSaldo").Value, "0.00")) Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("Unidad").Value = 0#
                dxDBGrid1.Dataset.Post
                MsgBox ("Usted no cuenta con las Unidades Solicitadas ...  Verifique"), vbInformation, App.Title
            ElseIf Val(Format(dxDBGrid1.Columns.ColumnByFieldName("Unidad").Value, "0.00")) > Val(Format(txt_CantSolictada.Text, "0.00")) Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("Unidad").Value = 0#
                dxDBGrid1.Dataset.Post
                MsgBox ("La Cantidad es Mayor a los requerido ...  Verifique"), vbInformation, App.Title
            End If
            
        Case dxDBGrid1.Columns.ColumnByFieldName("kg").Index
            
            DblDifRepeso = 0#
            If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("KgEnv").Value, "0.00")) > 0 Then
            
                DblDifRepeso = Val(Format((dxDBGrid1.Columns.ColumnByFieldName("KgEnv").Value - dxDBGrid1.Columns.ColumnByFieldName("kg").Value), "0.00"))
                
                If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("kg").Value, "0.00")) > Val(Format(dxDBGrid1.Columns.ColumnByFieldName("KgEnv").Value, "0.00")) Then
                
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Columns.ColumnByFieldName("kg").Value = 0#
                    dxDBGrid1.Columns.ColumnByFieldName("Rekg").Value = 0#
                    dxDBGrid1.Dataset.Post
                    
                    StrMsgError = "Los Kilos Recibidos no puede ser mayor a lo Enviado.. Verifique"
                    GoTo Err
                
                ElseIf Val(Format(dxDBGrid1.Columns.ColumnByFieldName("kg").Value + DblDifRepeso, "0.00")) > Val(Format(dxDBGrid1.Columns.ColumnByFieldName("kgSaldo").Value, "0.00")) Then
                
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Columns.ColumnByFieldName("kg").Value = 0#
                    dxDBGrid1.Columns.ColumnByFieldName("Rekg").Value = 0#
                    dxDBGrid1.Dataset.Post
                    
                    StrMsgError = "Usted no cuenta con los Kilos Solicitados ...  Verifique"
                    GoTo Err
                    
                ElseIf Val(Format(dxDBGrid1.Columns.ColumnByFieldName("kg").Value, "0.00")) > Val(Format(txt_PesoSolicitado.Text, "0.00")) Then
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Columns.ColumnByFieldName("kg").Value = 0#
                    dxDBGrid1.Columns.ColumnByFieldName("Rekg").Value = 0#
                    dxDBGrid1.Dataset.Post
                    
                    StrMsgError = "Los Kilos Es mayor a lo Requerido ...  Verifique"
                    GoTo Err
                End If
                
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("ReKg").Value = Val(Format((dxDBGrid1.Columns.ColumnByFieldName("KgEnv").Value - dxDBGrid1.Columns.ColumnByFieldName("kg").Value), "0.00"))
                dxDBGrid1.Dataset.Post
                
            Else
            
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("kg").Value = 0#
                dxDBGrid1.Columns.ColumnByFieldName("Rekg").Value = 0#
                dxDBGrid1.Dataset.Post
                
                MsgBox ("Debe Ingresas los Kg. Enviados ...  Verifique"), vbInformation, App.Title
            End If
            
                        
    End Select
    dxDBGrid1.Dataset.Edit
    dxDBGrid1.Dataset.Post
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub Form_Activate()
    
    dxDBGrid1.Option = egoAutoSearch
    dxDBGrid1.OptionEnabled = 0
    
    dxDBGrid1.Columns.FocusedIndex = 1
    dxDBGrid1.SetFocus
    
    dxDBGrid1.OptionEnabled = 1
    ConfGrid dxDBGrid1, True, True, False, False
    
End Sub

Public Sub mostrar_from(codigo_lote As String, codproducto As String, unidadprod As String, ByRef rscd As ADODB.Recordset, item As String, StrTipo As String, strSerie As String, strDocventas As String, strGlsproducto As String, indEstLiqui As Integer, DblCantidadLiq As Double, DblPesoLiq As Double, StrMsgError As String)
On Error GoTo Err

    indEstLiquix = 0
    CodProductoDocVentas = codproducto
    ItemDocVentas = item
    
    CodDocumento = StrTipo
    CodSerie = strSerie
    CodDocventas = strDocventas
    CodGlsProducto = strGlsproducto
    
    mostrarLiquidaciones unidadprod, StrMsgError
    
    lblGlsProducto = CodGlsProducto
    
    txt_CantSolictada.Text = DblCantidadLiq
    txt_PesoSolicitado.Text = DblPesoLiq
    
    FrmAyudaLiquidacion.Show 1
    codigo_lote = "" & codigo
    indEstLiqui = indEstLiquix
                
    dxDBGrid1.Dataset.Filter = ""
    dxDBGrid1.Dataset.Filtered = True

    Set dxDBGrid1.DataSource = Nothing
    
    If TypeName(rsg) = "Nothing" Then
        Exit Sub
    Else
        If rsg.State = 0 Then
            Exit Sub
        End If
    End If
                        
    If rsg.RecordCount > 0 Then
        rsg.MoveFirst
        Do While Not rsg.EOF
            If rsg.Fields("Unidad") = 0# Or rsg.Fields("kg") = 0# Then
                rsg.Delete adAffectCurrent
                rsg.Update
            End If
            rsg.MoveNext
        Loop
    End If
    
    Set rscd = rsg.Clone(adLockReadOnly)
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    Me.Hide
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError         As String
Dim CantSolicitada      As Double
Dim PesoSolicitado      As Double


    Select Case Button.Index
        Case 1 '---- Procesar}
        
            CantSolicitada = 0#
            PesoSolicitado = 0#
            
            'Valida Cantidad / Peso
            dxDBGrid1.Dataset.First
            If Not dxDBGrid1.Dataset.EOF Then
                dxDBGrid1.Dataset.First
                Do While Not dxDBGrid1.Dataset.EOF
                    CantSolicitada = CantSolicitada + Val(Format(dxDBGrid1.Columns.ColumnByFieldName("Unidad").Value, "0.00"))
                    PesoSolicitado = PesoSolicitado + Val(Format(dxDBGrid1.Columns.ColumnByFieldName("kg").Value, "0.00"))
                    dxDBGrid1.Dataset.Next
                Loop
            End If
            'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            If CantSolicitada <> txt_CantSolictada.Text Then
                StrMsgError = "Las Cantidad no coincide con lo Requerido"
                GoTo Err
            End If
            
            If PesoSolicitado <> txt_PesoSolicitado.Text Then
                StrMsgError = "Los Kilos no coincide con lo Requerido"
                GoTo Err
            End If
            
            indEstLiquix = 1
            If dxDBGrid1.Dataset.State = dsEdit Then dxDBGrid1.Dataset.Post
            If dxDBGrid1.Count > 0 Then
                Me.Hide
            End If
        Case 2
            indEstLiquix = 2
            If dxDBGrid1.Dataset.State = dsEdit Then dxDBGrid1.Dataset.Post
            If dxDBGrid1.Count > 0 Then
                Me.Hide
            End If
            
        Case 3 '---- Salir
            indEstLiquix = 3
            Unload Me
    End Select

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub
