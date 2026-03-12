VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmImportaCombos 
   Caption         =   "Importar Fórmulas"
   ClientHeight    =   8010
   ClientLeft      =   2010
   ClientTop       =   2295
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   11820
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      Caption         =   "Listado"
      ForeColor       =   &H00C00000&
      Height          =   7275
      Left            =   0
      TabIndex        =   0
      Top             =   675
      Width           =   11790
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   135
         TabIndex        =   1
         Top             =   195
         Width           =   10815
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   285
            Left            =   1500
            TabIndex        =   2
            Top             =   210
            Width           =   9240
            _ExtentX        =   16298
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
            MaxLength       =   255
            Container       =   "FrmImportaCombos.frx":0000
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
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
            TabIndex        =   3
            Top             =   210
            Width           =   915
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3015
         Left            =   180
         OleObjectBlob   =   "FrmImportaCombos.frx":001C
         TabIndex        =   4
         Top             =   870
         Width           =   11505
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   2955
         Left            =   150
         OleObjectBlob   =   "FrmImportaCombos.frx":2EB1
         TabIndex        =   5
         Top             =   4125
         Width           =   11520
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   10305
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":5881
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":5C1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":606D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":6407
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":67A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":6B3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":6ED5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":726F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":7609
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":79A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":7D3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":89FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportaCombos.frx":8D99
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
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1164
      ButtonWidth     =   767
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmImportaCombos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CIdCombo                        As String
Dim CIdComboAux                     As String

Public Sub MostrarForm(StrMsgError As String, PIdCombo As String)
On Error GoTo Err
    
    CIdComboAux = PIdCombo
    PIdCombo = ""
    
    Me.Show 1
    
    PIdCombo = CIdCombo
    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

    ConfGrid GLista, False, False, False, False
    ConfGrid gListaDetalle, False, False, False, False

    listaProducto StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    FraListado.Visible = True
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub listaProducto(ByRef StrMsgError As String)
Dim rst As New ADODB.Recordset
Dim strCond As String
Dim intNumNiveles As Integer

Dim strTabla As String
Dim strWhere As String
Dim strCampos As String
Dim strTablas As String
Dim strTablaAnt As String

Dim i As Integer
On Error GoTo Err

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND (glsCombo LIKE '%" & strCond & "%' or idComboCab LIKE '%" & strCond & "%') "
    End If

    csql = "SELECT idComboCab,glsCombo,DATE_FORMAT(FecEmision,GET_FORMAT(DATE, 'EUR')) as FecEmision,Format(TotalPrecioVenta,2) AS TotalPrecioVenta, idUM,idMoneda, idUsuario " & _
            "FROM combocab " & _
            "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' And IdComboCab Not In('" & CIdComboAux & "')"
    
    If strCond <> "" Then csql = csql + strCond

    csql = csql + " ORDER BY idComboCab, FecEmision"

    With GLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn

        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idComboCab"
    End With

    listaDetalle StrMsgError
    If StrMsgError <> "" Then GoTo Err

Me.Refresh
If rst.State = 1 Then rst.Close
Set rst = Nothing
Exit Sub
Err:
If rst.State = 1 Then rst.Close
Set rst = Nothing
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub listaDetalle(StrMsgError As String)
On Error GoTo Err

    csql = "SELECT item, idProducto, GlsProducto, idUM, Format(Cantidad,2) AS Cantidad, Format(PVUnit,2) AS PVUnit " & _
           "FROM combodet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idComboCab = '" & GLista.Columns.ColumnByFieldName("idComboCab").Value & "'"
    
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
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrMsgError                 As String
On Error GoTo Err
    
    listaDetalle StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnDblClick()
Dim StrMsgError                 As String
On Error GoTo Err
    
    CIdCombo = GLista.Columns.ColumnByFieldName("idComboCab").Value
    Me.Hide
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
Dim strCodUltProd As String
On Error GoTo Err

    Select Case Button.Index
    
        Case 1 'Salir
            CIdCombo = ""
            Me.Hide
            
    End Select

Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
Dim StrMsgError As String
On Error GoTo Err
listaProducto StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then GLista.SetFocus
End Sub
