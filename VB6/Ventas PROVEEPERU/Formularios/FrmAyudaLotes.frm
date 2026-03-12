VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmAyudaLotes 
   Caption         =   "Lista de Tallas"
   ClientHeight    =   5610
   ClientLeft      =   6600
   ClientTop       =   4395
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   7575
   Begin VB.Frame Frame2 
      Height          =   4830
      Left            =   0
      TabIndex        =   0
      Top             =   675
      Width           =   7575
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   3990
         Left            =   135
         OleObjectBlob   =   "FrmAyudaLotes.frx":0000
         TabIndex        =   1
         Top             =   270
         Width           =   7335
      End
      Begin VB.Label lblCantidad 
         Alignment       =   2  'Center
         Caption         =   "lblCantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   3
         Top             =   4320
         Width           =   7350
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
            Picture         =   "FrmAyudaLotes.frx":33E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":377A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":3BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":3F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":4300
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":469A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":4A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":4DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":5168
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":5502
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":589C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":655E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":68F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":6D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":70E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAyudaLotes.frx":7AF6
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
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1164
      ButtonWidth     =   2487
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "       Procesar       "
            Object.ToolTipText     =   "Procesar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmAyudaLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw_limpia            As Boolean
Dim codigo               As String
Dim DesLote              As String
Dim Descripcion          As String
Dim CantProDocVenta      As String
Dim rsg                  As New ADODB.Recordset
Dim CodProductoDocVentas As String
Dim ItemDocVentas        As String

Private Sub fill(codproducto As String, codalmacen As String, StrMsgError As String)
On Error GoTo Err
Dim CSqlC                   As String
Dim rst                     As New ADODB.Recordset
Dim item                    As Integer

''    csql = "select l.idlote,l.glslote,if(pl.idalmacen = '" & codalmacen & "' And pl.idproducto = '" & codproducto & "',pl.Cantidadstock,0) AS Cantidadstock from LOTES  l left join productosalmacenporlote pl " & _
''           "on l.idlote = pl.idlote " & _
''           "and l.idempresa = pl.idempresa " & _
''           "and l.idsucursal = pl.idsucursal " & _
''           "and l.estado = 'ACT' " & _
''           "where L.IDSUCURSAL = '" & glsSucursal & "' " & _
''           "and l.idempresa = '" & glsEmpresa & "' " & _
''           "And pl.idproducto = '" & codproducto & "' " & _
''           "and pl.idalmacen = '" & Trim(codalmacen) & _
''           "' group by l.idlote,pl.idproducto"
           
           
    'CSqlC = "Select X.IdLote,L.GlsLote,Sum(X.CantidadStock) SaldoTotal " & _
            "From(" & _
                "Select VD.IdSucursal,VD.IdEmpresa,If(Length(Trim(IfNull(VD.idlote,''))) = 0 And VD.IdSucursal = '08090001','10120001', " & _
                "If(VD.IdSucursal = '08090002','10120002',VD.idlote)) IdLote," & _
                "If(VC.IdAlmacen = '" & CodAlmacen & "' And VD.IdProducto = '" & codproducto & "',Sum(If(VD.TipoVale = 'I',VD.Cantidad,VD.Cantidad * -1)),0) CantidadStock " & _
                "From ValesCab Vc " & _
                "Inner Join PeriodosInv P " & _
                    "On Vc.IdEmpresa = P.IdEmpresa And Vc.IdSucursal = P.IdSucursal And Vc.IdPeriodoInv = P.IdPeriodoInv " & _
                "Inner Join Almacenes A " & _
                    "On Vc.IdEmpresa = A.IdEmpresa And Vc.IdAlmacen = A.IdAlmacen " & _
                "Inner Join ValesDet Vd " & _
                    "On Vc.IdValesCab = Vd.IdValesCab And Vc.IdEmpresa = Vd.IdEmpresa And Vc.IdSucursal = Vd.IdSucursal And Vc.TipoVale = Vd.TipoVale " & _
                "Where Vd.IdEmpresa = '" & glsEmpresa & "' And Vd.IdSucursal = '" & glsSucursal & "' And P.EstPeriodoInv = 'ACT' " & _
                "And Vd.IdProducto = '" & codproducto & "' And Vc.IdAlmacen = '" & Trim(CodAlmacen) & "' And Vc.EstValeCab <> 'ANU' " & _
                "Group By Vd.IdProducto,Vd.IdLote" & _
            ") X " & _
            "Inner Join Lotes L " & _
                "On X.IdEmpresa = L.IdEmpresa And X.IdSucursal = L.IdSucursal And X.IdLote = L.IdLote And L.Estado = 'ACT' " & _
            "Group By X.IdLote "
             
             
    'A.IdPeriodoInv = '" & glsCodPeriodoINV & "'
             
    CSqlC = "Select B.IdLote,C.GlsLote,C.FechaLote,Sum(B.Cantidad * CASE WHEN B.TipoVale = 'I' THEN 1 ELSE -1 END) SaldoTotal " & _
            "From ValesCab A " & _
            "Inner Join ValesDetLotes B " & _
                "On A.IdEmpresa = B.IdEmpresa And A.IdSucursal = B.IdSucursal And A.TipoVale = B.TipoVale And A.IdValesCab = B.IdValesCab " & _
            "Inner Join Lotes C " & _
                "On B.IdEmpresa = C.IdEmpresa And B.IdLote = C.IdLote " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' " & _
            "And A.EstValeCab <> 'ANU' " & _
            "And A.IdAlmacen = '" & codalmacen & "' And B.IdProducto = '" & codproducto & "' And C.Estado = 'ACT' " & _
            "Group By B.IdLote,C.GlsLote,C.FechaLote " & _
            "Order By C.FechaLote"
             
    rst.Open CSqlC, Cn, adOpenKeyset, adLockOptimistic
    
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "idLote", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "GlsLote", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "CantidadStock", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "ok", adVarChar, 1, adFldIsNullable
    rsg.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "idProducto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "ItemDocVentas", adVarChar, 3, adFldIsNullable
    rsg.Open
    
    item = 0
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
    
        rsg.Fields("Item") = 1
        rsg.Fields("idLote") = ""
        rsg.Fields("GlsLote") = ""
        rsg.Fields("CantidadStock") = 0
        rsg.Fields("ok") = "0"
        rsg.Fields("Cantidad") = 0
        rsg.Fields("idProducto") = ""
        rsg.Fields("ItemDocVentas") = ""
        
    Else
    
        rst.MoveFirst
        Do While Not rst.EOF
        
            rsg.AddNew
            item = item + 1
            rsg.Fields("Item") = item
            rsg.Fields("idLote") = Trim(rst.Fields("idLote") & "")
            rsg.Fields("GlsLote") = Trim(rst.Fields("GlsLote") & "")
            rsg.Fields("ok") = "0"
            rsg.Fields("CantidadStock") = Val(Format((rst.Fields("SaldoTotal")), "0.00"))
            rsg.Fields("Cantidad") = 0#
            rsg.Fields("idproducto") = Trim("" & CodProductoDocVentas)
            rsg.Fields("ItemDocVentas") = Trim("" & ItemDocVentas)
            
            rst.MoveNext
            
        Loop
    End If

    mostrarDatosGridSQL dxDBGrid1, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
On Error GoTo Err
Dim StrMsgError As String

    If dxDBGrid1.Dataset.State = dsEdit Then
        dxDBGrid1.Dataset.Post
    End If
    
    If State = cbsChecked Then
    
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Value = dxDBGrid1.Columns.ColumnByFieldName("CantidadStock").Value
        dxDBGrid1.Dataset.Post
        
        If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("cantidad").Value, "0.00")) > Val(Format(CantProDocVenta, "0.00")) Then
        
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Value = "0"
            dxDBGrid1.Columns.ColumnByFieldName("OK").Value = "0"
            dxDBGrid1.Dataset.Post
            
            dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Index
            
            MsgBox ("la Cantidad Ingresada es mayor a la Cantidad de la Venta ..  Verifique"), vbInformation, App.Title
            
            Exit Sub
        End If
    
    Else
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Value = "0"
        dxDBGrid1.Dataset.Post
    
    End If
            
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
    
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String
Dim sw_nc   As Boolean

    Select Case dxDBGrid1.Columns.FocusedColumn.Index
        Case dxDBGrid1.Columns.ColumnByFieldName("cantidad").Index
            If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("cantidad").Value, "0.00")) > Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CantidadStock").Value, "0.00")) Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Value = "0"
                dxDBGrid1.Dataset.Post
                MsgBox ("Usted no cuenta con la Cantidad Solicitada ..  Verifique"), vbInformation, App.Title
            End If
                
            If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("cantidad").Value, "0.00")) > Val(Format(CantProDocVenta, "0.00")) Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Value = "0"
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Index
                MsgBox ("la Cantidad Ingresada es mayor a la Cantidad de la Venta ..  Verifique"), vbInformation, App.Title
                Exit Sub
            End If
    End Select

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

Public Sub mostrar_from(XDesLote As String, codigo_lote As String, codproducto As String, codalmacen As String, cantidad As String, ByRef rscd As ADODB.Recordset, item As String, StrMsgError As String)
On Error GoTo Err
    codigo = ""
    DesLote = ""
    CantProDocVenta = cantidad
    CodProductoDocVentas = codproducto
    ItemDocVentas = item
    
    lblCantidad.Caption = "Cantidad Solicitada : " & " " & cantidad
    
    fill codproducto, codalmacen, StrMsgError
    FrmAyudaLotes.Show 1
    codigo_lote = "" & codigo
    XDesLote = "" & DesLote
        
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
            If rsg.Fields("cantidad") = "0" Then
                rsg.Delete adAffectCurrent
                rsg.Update
            End If
            rsg.MoveNext
        Loop
    End If
                
    Set rscd = rsg.Clone(adLockReadOnly)
    
    If rsg.State = 1 Then rsg.Close
    Set rsg = Nothing
        
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
Dim StrMsgError As String
Dim cantGRid    As Double
Dim StockGrid   As String
Dim item        As Integer

    Select Case Button.Index
        Case 1 '---- Procesar
            cantGRid = 0#
            
''            dxDBGrid1.Dataset.First
''            dxDBGrid1.Dataset.Filter = ""
''            dxDBGrid1.Dataset.Filter = "  Cantidad <> 0 "
''            If Not dxDBGrid1.Dataset.EOF Then
''                If dxDBGrid1.Dataset.RecordCount > 1 Then
''                    MsgBox "Solo puede seleccionar una talla...", vbInformation, App.Title
''                    Exit Sub
''                End If
''            End If
''
''            xDBGrid1.Dataset.Filter = ""

            dxDBGrid1.Dataset.First
            If Not dxDBGrid1.Dataset.EOF Then
                dxDBGrid1.Dataset.First
                item = 0
                Do While Not dxDBGrid1.Dataset.EOF
                    If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Value, "0.00")) > 0 Then
                        item = item + 1
                    End If

                    If Val(item) > 1 Then
                        MsgBox "Solo se puede elegir una Talla ...", vbInformation, App.Title
                        Exit Sub
                    End If
            
                    dxDBGrid1.Dataset.Next
                Loop
            End If
                        
            dxDBGrid1.Dataset.First
            If Not dxDBGrid1.Dataset.EOF Then
                dxDBGrid1.Dataset.First
                Do While Not dxDBGrid1.Dataset.EOF
                    cantGRid = cantGRid + Val(Format(dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Value, "0.00"))
                    dxDBGrid1.Dataset.Next
                Loop
            End If
                            
                            
            dxDBGrid1.Dataset.First
            If Not dxDBGrid1.Dataset.EOF Then
                dxDBGrid1.Dataset.First
                item = 0
                Do While Not dxDBGrid1.Dataset.EOF
                    If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("Cantidad").Value, "0.00")) > 0 Then
                        codigo = ("" & dxDBGrid1.Columns.ColumnByFieldName("idlote").Value)
                        DesLote = ("" & dxDBGrid1.Columns.ColumnByFieldName("GlsLote").Value)
                    End If
                    dxDBGrid1.Dataset.Next
                Loop
            End If
            
            
            If Val(Format(cantGRid, "0.00")) <> Val(Format(CantProDocVenta, "0.00")) Then
                MsgBox "El Total de La Cantidad Digitada es Diferente a la Cantidad de la Venta Verifique ...", vbInformation, App.Title
                Exit Sub
            Else
                If dxDBGrid1.Count > 0 Then
                    Me.Hide
                End If
            End If
        
        Case 2 '---- Salir
            Unload Me
    End Select

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub
