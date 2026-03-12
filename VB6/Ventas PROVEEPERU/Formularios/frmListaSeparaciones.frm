VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmListaSeparaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Separaciones"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   12960
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   840
      Top             =   960
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
            Picture         =   "frmListaSeparaciones.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":317E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":3518
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":396A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":3D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSeparaciones.frx":4716
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
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagos"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refrescar"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gen. Doc."
            ImageIndex      =   15
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Height          =   8865
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   12915
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   12690
         Begin VB.ComboBox cbx_Mes 
            Height          =   315
            ItemData        =   "frmListaSeparaciones.frx":4DE8
            Left            =   8400
            List            =   "frmListaSeparaciones.frx":4E10
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   225
            Width           =   1665
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   285
            Left            =   1500
            TabIndex        =   4
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
            Container       =   "frmListaSeparaciones.frx":4E79
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Ano 
            Height          =   285
            Left            =   11175
            TabIndex        =   5
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
            Container       =   "frmListaSeparaciones.frx":4E95
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "(Filtra por Razon social del cliente o Numero del documento )"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   1950
            TabIndex        =   9
            Top             =   480
            Width           =   4275
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Año:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   10725
            TabIndex        =   8
            Top             =   300
            Width           =   405
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Mes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7800
            TabIndex        =   7
            Top             =   300
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Busqueda:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   210
            Width           =   915
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   5025
         Left            =   60
         OleObjectBlob   =   "frmListaSeparaciones.frx":4EB1
         TabIndex        =   10
         Top             =   1125
         Width           =   12735
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   2370
         Left            =   75
         OleObjectBlob   =   "frmListaSeparaciones.frx":8FF2
         TabIndex        =   11
         Top             =   6300
         Width           =   12735
      End
   End
End
Attribute VB_Name = "frmListaSeparaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim indNuevoDoc As Boolean
Const strTipoDoc As String = "91"


Private Sub listaDocVentas(ByRef StrMsgError As String)
Dim strCond As String
On Error GoTo Err
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND (GlsCliente LIKE '%" & strCond & "%' or idDocVentas LIKE '%" & strCond & "%') "
    End If
    
    csql = "SELECT Item, idDocVentas,idSerie,idPerCliente,GlsCliente,RUCCliente," & _
                  "DATE_FORMAT(FecEmision,GET_FORMAT(DATE, 'EUR')) as FecEmision,estDocVentas,idMoneda,Format(TotalPrecioVenta,2) AS TotalPrecioVenta, " & _
                  "Format(Adelantos,2) AS Adelantos, Format(TotalPrecioVenta - Adelantos,2) AS SaldoSeparacion " & _
           "FROM " & _
           "(SELECT concat(idDocumento,idDocVentas,idSerie) as Item , idDocVentas,idSerie,idPerCliente,GlsCliente,RUCCliente," & _
                  "FecEmision,estDocVentas,idMoneda,TotalPrecioVenta, " & _
                  "(SELECT SUM(c.ValMonto) FROM movcajasdet c WHERE c.idEmpresa = '" & glsEmpresa & "' AND c.idSucursal = '" & glsSucursal & "' AND c.idDocumento = '" & strTipoDoc & "' AND c.idSerie = docventas.idSerie AND c.idDocVentas = docventas.idDocVentas) AS Adelantos " & _
           "FROM docventas " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTipoDoc & "' AND year(FecEmision) = " & Val(txt_Ano.Text) & " AND Month(FecEmision) = " & cbx_Mes.ListIndex + 1 & strCond & ") Separaciones"

    csql = csql + " ORDER BY idSerie,idDocVentas,FecEmision"
    
    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With
    
    'DETALLE
    
    listaDetalle
    
Me.Refresh
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub


Private Sub listaDetalle()
    csql = "SELECT item, idProducto, GlsProducto, GlsMarca, GlsUM, Format(Cantidad,2) AS Cantidad, Format(PVUnit,2) AS PVUnit, FORMAT(PorDcto,2) AS PorDcto, Format(TotalPVNeto,2) AS TotalPVNeto " & _
           "FROM docventasdet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & strTipoDoc & "' AND idDocVentas = '" & gLista.Columns.ColumnByFieldName("idDocVentas").Value & "' AND idSerie = '" & gLista.Columns.ColumnByFieldName("idSerie").Value & "'"
    
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

Private Sub cbx_Mes_Click()
Dim StrMsgError As String
On Error GoTo Err
If indNuevoDoc = False Then
    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err
End If
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

indNuevoDoc = True

Me.top = 0
Me.left = 0

txt_Ano.Text = Year(getFechaSistema)
cbx_Mes.ListIndex = Month(getFechaSistema) - 1

ConfGrid gLista, False, False, False, False
ConfGrid gListaDetalle, False, False, False, False

listaDocVentas StrMsgError
If StrMsgError <> "" Then GoTo Err

indNuevoDoc = False
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    listaDetalle
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
On Error GoTo Err
Select Case Button.Index
    Case 1 'Pagos
    
           frmPagosSeparacion.MostrarForm strTipoDoc, gLista.Columns.ColumnByFieldName("idDocVentas").Value, gLista.Columns.ColumnByFieldName("idSerie").Value, StrMsgError
           If StrMsgError <> "" Then GoTo Err
           
           listaDocVentas StrMsgError
           If StrMsgError <> "" Then GoTo Err
           
    Case 2 'Cancelar
        Unload Me
    Case 3 'Refrescar
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Case 4 'Gen. Doc.
    Case 5 'Salir
        Unload Me
End Select
Exit Sub
Err:
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_Ano_Change()
Dim StrMsgError As String
On Error GoTo Err
If indNuevoDoc = False Then
    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err
End If
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
Dim StrMsgError As String
On Error GoTo Err
listaDocVentas StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub


