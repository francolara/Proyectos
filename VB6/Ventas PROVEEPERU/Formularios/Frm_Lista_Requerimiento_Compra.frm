VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form Frm_Lista_Requerimiento_Compra 
   Caption         =   "Requerimiento de Compra"
   ClientHeight    =   7635
   ClientLeft      =   2940
   ClientTop       =   1920
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   9990
   Begin VB.Frame Fra_General 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7530
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   9945
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3300
         Left            =   135
         OleObjectBlob   =   "Frm_Lista_Requerimiento_Compra.frx":0000
         TabIndex        =   3
         Top             =   495
         Width           =   9705
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gProductos 
         Height          =   3210
         Left            =   135
         OleObjectBlob   =   "Frm_Lista_Requerimiento_Compra.frx":4598
         TabIndex        =   4
         Top             =   4155
         Width           =   9705
      End
      Begin VB.Label Label1 
         Caption         =   "Productos Pendientes por Atender"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Width           =   4035
      End
      Begin VB.Label lblnumero_Requerimiento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         TabIndex        =   2
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label lbletiqueta_requerimiento 
         Caption         =   "Requerimiento  de Compra Nº"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   4035
      End
   End
End
Attribute VB_Name = "Frm_Lista_Requerimiento_Compra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError         As String

    ConfGrid GLista, False, False, False, False
    ConfGrid gProductos, False, False, False, False
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub MostrarFrom(strTipoDoc As String, strnum As String, strSerie As String)
Dim rst  As New ADODB.Recordset
    
    csql = "Select tipoDocOrigen, numDocOrigen, serieDocOrigen, tipoDocReferencia, " & _
           "numDocReferencia, serieDocReferencia, item, idEmpresa, idSucursal, idSucursalReferencia " & _
           "From docreferencia " & _
            "Where IdEmpresa='" & glsEmpresa & "' And idSucursal = '" & glsSucursal & "' And tipoDocReferencia ='" & strTipoDoc & "' And tipoDocOrigen='87' " & _
            "And serieDocReferencia= '" & strSerie & "' And numDocReferencia= '" & strnum & "' Order by numDocOrigen Desc "
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        If rst.Fields("tipoDocOrigen").Value = "87" Then
            csql = "Select d.* " & _
                    "From DocVentas c " & _
                    "Inner Join  DocVentasDet d " & _
                         "On c.IdDocumento =d.IdDocumento  " & _
                         "And c.IdDocVentas = d.IdDocVentas " & _
                         "And c.IdSerie = d.IdSerie  " & _
                         "And c.IdEmpresa = d.IdEmpresa  " & _
                         "And c.IdSucursal =d.IdSucursal " & _
                    "Where d.IdEmpresa='" & glsEmpresa & "'" & _
                     "And d.idSucursal = '" & glsSucursal & "' " & _
                     "And d.IdDocumento ='" & rst.Fields("tipoDocOrigen").Value & "' " & _
                     "And d.IdSerie= '" & rst.Fields("serieDocOrigen").Value & "' " & _
                     "And d.IdDocVentas= '" & rst.Fields("numDocOrigen").Value & "' "
    
            With GLista
                .DefaultFields = False
                .Dataset.ADODataset.ConnectionString = strcn
                .Dataset.ADODataset.CursorLocation = clUseClient
                .Dataset.Active = False
                .Dataset.ADODataset.CommandText = csql
                .Dataset.DisableControls
                .Dataset.Active = True
                .KeyField = "item"
                .Dataset.Refresh
            End With
       
            csql = "Select idDocumento, idDocVentas, idSerie, idProducto, glsProducto, idMarca, idUM, Factor, Afecto, Cantidad, VVUnit, IGVUnit, PVUnit, TotalVVBruto, TotalPVBruto, PorDcto, DctoVV, DctoPV, TotalVVNeto, TotalIGVNeto, TotalPVNeto, item, GlsMarca, GlsUM, idTipoProducto, idMoneda, idCodFabricante, idEmpresa, idSucursal, estDocImportado, idDocumentoImp, idDocVentasImp, idSerieImp, NumLote, FecVencProd, idUsuarioDcto, VVUnitLista, PVUnitLista, CantidadImp, VVUnitNeto, PVUnitNeto, Cantidad2, CodigoRapido, idTallaPeso, CantidadAnt, Simbolo1, Simbolo2, Simbolo3, itemPro, PorcUtilidad, IdCentroCosto, IdSucursalPres, IdDocumentoPres, IdSeriePres, IdDocVentasPres, GlsPlaca, IdUPCliente From docVentasDet " & _
                    "Where idDocventas ='" & GLista.Columns.ColumnByFieldName("idDocventas").Value & "' " & _
                    "And idSerie ='" & GLista.Columns.ColumnByFieldName("idSerie").Value & "' " & _
                    "And idDocumento  ='" & GLista.Columns.ColumnByFieldName("IdDocumento").Value & "' " & _
                    "And IdEmpresa='" & glsEmpresa & "'" & _
                    "And idSucursal = '" & glsSucursal & "' " & _
                    "And( estDocImportado= '' or estDocImportado='N')"
            
            With gProductos
                 .DefaultFields = False
                 .Dataset.ADODataset.ConnectionString = strcn
                 .Dataset.ADODataset.CursorLocation = clUseClient
                 .Dataset.Active = False
                 .Dataset.ADODataset.CommandText = csql
                 .Dataset.DisableControls
                 .Dataset.Active = True
                 .KeyField = "item"
                 .Dataset.Refresh
            End With
            
            rst.Close: Set rst = Nothing
            lblnumero_Requerimiento.Caption = GLista.Columns.ColumnByFieldName("idDocventas").Value
        End If
    End If
    Me.Show 1
    
End Sub
