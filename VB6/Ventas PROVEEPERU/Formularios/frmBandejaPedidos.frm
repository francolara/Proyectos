VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmBandejaPedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bandeja de Pedidos"
   ClientHeight    =   9705
   ClientLeft      =   2610
   ClientTop       =   1065
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   1164
      ButtonWidth     =   2090
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aprobar"
            Object.ToolTipText     =   "Aprobar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quitar Aprob."
            Object.ToolTipText     =   "Quitar Aprob."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refrescar"
            Object.ToolTipText     =   "Refrescar"
            ImageIndex      =   14
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
      Caption         =   "Listado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   9000
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   12900
      Begin MSComctlLib.ImageList imgDocVentas 
         Left            =   1125
         Top             =   3150
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
               Picture         =   "frmBandejaPedidos.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":039A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":07EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":0B86
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":0F20
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":12BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":1654
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":19EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":1D88
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":2122
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":24BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":317E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":3518
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":396A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":3D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBandejaPedidos.frx":4716
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   75
         TabIndex        =   5
         Top             =   600
         Width           =   12690
         Begin VB.ComboBox cbx_Mes 
            Height          =   315
            ItemData        =   "frmBandejaPedidos.frx":4DE8
            Left            =   8400
            List            =   "frmBandejaPedidos.frx":4E10
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   225
            Width           =   1665
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   285
            Left            =   1500
            TabIndex        =   7
            Top             =   210
            Width           =   5340
            _ExtentX        =   9419
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
            Container       =   "frmBandejaPedidos.frx":4E79
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Ano 
            Height          =   285
            Left            =   11175
            TabIndex        =   8
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmBandejaPedidos.frx":4E95
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Busqueda:"
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
            Left            =   120
            TabIndex        =   12
            Top             =   285
            Width           =   780
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Mes:"
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
            Left            =   7800
            TabIndex        =   11
            Top             =   285
            Width           =   345
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Año:"
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
            Left            =   10725
            TabIndex        =   10
            Top             =   285
            Width           =   345
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "(Filtra por Razon social del cliente o Numero del documento )"
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
            Left            =   1890
            TabIndex        =   9
            Top             =   480
            Width           =   4395
         End
      End
      Begin VB.ComboBox cbxLista 
         Height          =   315
         ItemData        =   "frmBandejaPedidos.frx":4EB1
         Left            =   10125
         List            =   "frmBandejaPedidos.frx":4EBB
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   2640
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   4260
         Left            =   75
         OleObjectBlob   =   "frmBandejaPedidos.frx":4EE1
         TabIndex        =   4
         Top             =   1425
         Width           =   12735
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   3165
         Left            =   75
         OleObjectBlob   =   "frmBandejaPedidos.frx":903A
         TabIndex        =   13
         Top             =   5775
         Width           =   12735
      End
      Begin VB.Label Label1 
         Caption         =   "Mostrar los Pedidos:"
         Height          =   240
         Left            =   8325
         TabIndex        =   2
         Top             =   255
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmBandejaPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indCargando As Boolean

Private Sub cbx_Mes_Click()
Dim StrMsgError As String
On Error GoTo Err

If indCargando Then Exit Sub

listaPedidos StrMsgError
If StrMsgError <> "" Then GoTo Err

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cbxLista_Click()
Dim StrMsgError As String
On Error GoTo Err

If cbxLista.ListIndex = 0 Then
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
Else
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
End If

If indCargando Then Exit Sub

listaPedidos StrMsgError
If StrMsgError <> "" Then GoTo Err

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

Me.top = 0
Me.left = 0

indCargando = True

cbxLista.ListIndex = 0
txt_Ano.Text = Year(getFechaSistema)
cbx_Mes.ListIndex = Month(getFechaSistema) - 1

ConfGrid gLista, False, False, False, False
ConfGrid gListaDetalle, False, False, False, False

listaPedidos StrMsgError
If StrMsgError <> "" Then GoTo Err

indCargando = False

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub listaPedidos(ByRef StrMsgError As String)
Dim strCond As String
Dim strFiltro As String
On Error GoTo Err
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND (GlsCliente LIKE '%" & strCond & "%' or idDocVentas LIKE '%" & strCond & "%') "
    End If
    
    strFiltro = " AND (indAprobado = '0' or IfNull(indAprobado,'') = '') "
    If cbxLista.ListIndex = 1 Then strFiltro = " AND indAprobado = '1' "
    
    csql = "SELECT concat(idSucursal,idDocumento,idDocVentas,idSerie) as Item ,idSucursal,personas.GlsPersona as GlsSucursal, idDocVentas,idSerie,idPerCliente,GlsCliente,RUCCliente,DATE_FORMAT(FecEmision,GET_FORMAT(DATE, 'EUR')) as FecEmision,estDocVentas,Format(TotalPrecioVenta,2) AS TotalPrecioVenta " & _
            "FROM docventas, personas " & _
            "WHERE docventas.idSucursal = personas.idPersona AND idEmpresa = '" & glsEmpresa & "' AND idDocumento = '40' AND year(FecEmision) = " & Val(txt_Ano.Text) & " AND Month(FecEmision) = " & cbx_Mes.ListIndex + 1 & strFiltro
           
    If strCond <> "" Then csql = csql + strCond

    csql = csql + " ORDER BY idSerie,idDocVentas,FecEmision"
    
    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With
    
    If gLista.Ex.GroupColumnCount = 0 Then gLista.Columns.ColumnByFieldName("GlsSucursal").GroupIndex = 0
    
    'DETALLE
    
    ListaDetalle
    
Me.Refresh
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub


Private Sub ListaDetalle()
    csql = "SELECT item, idProducto, GlsProducto, GlsMarca, GlsUM, Format(Cantidad,2) AS Cantidad, Format(PVUnit,2) AS PVUnit, FORMAT(PorDcto,2) AS PorDcto, Format(TotalPVNeto,2) AS TotalPVNeto " & _
           "FROM docventasdet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & gLista.Columns.ColumnByFieldName("idSucursal").Value & "' AND idDocumento = '40' AND idDocVentas = '" & gLista.Columns.ColumnByFieldName("idDocVentas").Value & "' AND idSerie = '" & gLista.Columns.ColumnByFieldName("idSerie").Value & "'"
    
    With gListaDetalle
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With
End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    ListaDetalle
End Sub

Private Sub gLista_OnDblClick()
    frmSituacionCre.MostrarFrom gLista.Columns.ColumnByFieldName("idPerCliente").Value, gLista.Columns.ColumnByFieldName("GlsCliente").Value, gLista.Columns.ColumnByFieldName("RucCliente").Value
End Sub

Private Sub gLista_OnReloadGroupList()
  gLista.m.FullExpand
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
On Error GoTo Err

Select Case Button.Index
    Case 1 'Aprobar
    
        If MsgBox("¿Seguro de Aprobar el pedido?", vbQuestion + vbYesNo, App.Title) = vbYes Then
    
            csql = "UPDATE docventas SET indAprobado = '1' " & _
                   "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & gLista.Columns.ColumnByFieldName("idSucursal").Value & "' AND idDocumento = '40' AND idDocVentas = '" & gLista.Columns.ColumnByFieldName("idDocVentas").Value & "' AND idSerie = '" & gLista.Columns.ColumnByFieldName("idSerie").Value & "'"
                   
            Cn.Execute csql
            
            listaPedidos StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        End If
    Case 2 'Quitar Aprobacion
        If MsgBox("¿Seguro de Quitar la Aprobacion del pedido?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        
            'falta validar q cuando quitan aprobacion ya no este importado
        
            csql = "UPDATE docventas SET indAprobado = '0' " & _
                   "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & gLista.Columns.ColumnByFieldName("idSucursal").Value & "' AND idDocumento = '40' AND idDocVentas = '" & gLista.Columns.ColumnByFieldName("idDocVentas").Value & "' AND idSerie = '" & gLista.Columns.ColumnByFieldName("idSerie").Value & "'"
                   
            Cn.Execute csql
            
            listaPedidos StrMsgError
            If StrMsgError <> "" Then GoTo Err
        End If
    Case 3 'Refrescar
        listaPedidos StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Case 4 'Salir
        Unload Me
End Select

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_Ano_Change()
Dim StrMsgError As String
On Error GoTo Err

If indCargando Then Exit Sub

listaPedidos StrMsgError
If StrMsgError <> "" Then GoTo Err

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub


Private Sub txt_TextoBuscar_Change()
Dim StrMsgError As String
On Error GoTo Err

If indCargando Then Exit Sub

listaPedidos StrMsgError
If StrMsgError <> "" Then GoTo Err

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub
