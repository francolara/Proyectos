VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form FrmBusquedaClientes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Clientes"
   ClientHeight    =   5655
   ClientLeft      =   4770
   ClientTop       =   4170
   ClientWidth     =   10845
   Icon            =   "FrmBusquedaClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   10845
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   " Tiendas "
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
      Height          =   2115
      Left            =   3360
      TabIndex        =   8
      Top             =   7590
      Width           =   10650
      Begin DXDBGRIDLibCtl.dxDBGrid gTiendas 
         Height          =   1815
         Left            =   75
         OleObjectBlob   =   "FrmBusquedaClientes.frx":000C
         TabIndex        =   9
         Top             =   225
         Width           =   10470
      End
   End
   Begin VB.CommandButton Cmdbusq 
      DownPicture     =   "FrmBusquedaClientes.frx":2533
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10350
      Picture         =   "FrmBusquedaClientes.frx":28DE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   225
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox TxtBusq 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1140
      MaxLength       =   255
      TabIndex        =   0
      Top             =   300
      Width           =   9105
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   75
      TabIndex        =   6
      Top             =   675
      Width           =   10680
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   4125
         Left            =   75
         OleObjectBlob   =   "FrmBusquedaClientes.frx":2C89
         TabIndex        =   1
         Top             =   150
         Width           =   10500
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      MaxLength       =   255
      TabIndex        =   7
      Top             =   4575
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Presionar Enter en el registro para obtener el resultado "
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
      Height          =   465
      Left            =   90
      TabIndex        =   3
      Top             =   5100
      Width           =   3120
   End
   Begin VB.Label LblReg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "(0) Registros"
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
      Height          =   285
      Left            =   8850
      TabIndex        =   4
      Top             =   5100
      Width           =   1905
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Búsqueda"
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
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   315
      Width           =   915
   End
End
Attribute VB_Name = "FrmBusquedaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Dim cadRelacion As String
Dim cadCondicion As String
Dim cadBusqueda As String
Private SqlAdic As String
Private sqlBus As String
Private sqlCond As String
Private SRptBus(5) As String
Private strAyuda As String
Private EsNuevo As Boolean

Private Sub CmdBusq_Click()
Dim StrMsgError As String

On Error GoTo Err

    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then GoTo Err
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Activate()
'TxtBusq.SetFocus
End Sub

Private Sub Form_Deactivate()
'SqlAdic = ""
If EsNuevo = False Then
    TxtBusq.Text = ""
End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim codCliente As String
If KeyCode = 45 Then g_OnKeyDown KeyCode, 1

'If KeyCode = 45 Then
'    frmMantPersonaRapido.mostrarForm codCliente
'
'    If codCliente <> "" Then
'        SRptBus(0) = codCliente
'        SRptBus(1) = traerCampo("Personas", "GlsPersona", "idPersona", codCliente, False)
'
'        g.Dataset.Close
'        g.Dataset.Active = False
'        'Set g = Nothing
'        'Unload Me
'        Me.Hide
'    End If
'
'End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
    
    Me.Caption = "Búsqueda de Clientes"
    ConfGrid G, False, False, False, False
    ConfGrid gTiendas, False, False, False, False
    EsNuevo = True
    
    If leeParametro("VISUALIZA_NOMBRE_COMERCIAL") = "S" Then
        
        G.Columns.ColumnByFieldName("GlsNombreComercial").Visible = True
    
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'SqlAdic = ""
TxtBusq.Text = ""
End Sub

Private Sub G_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim StrMsgError As String
    
    listaTiendas StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub G_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
  
'  u = g.Columns.ColumnByFieldName("cod").Index
'
'  If Node.Index = 1 Then
'     Color = &HC0E0FF
'  Else
'     Color = &HC0FFFF
'  End If
End Sub

Private Sub g_OnDblClick()
    devolverResultado False
End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Dim codCliente As String
Dim rst As New ADODB.Recordset

On Error Resume Next

Select Case KeyCode
 Case 13:
      
    devolverResultado True

 Case 27
    Text1.SetFocus
 Case 45
    frmMantPersonaRapido.MostrarForm codCliente
    
    If codCliente <> "" Then
        SRptBus(0) = codCliente
        
        csql = "SELECT p.ruc,concat(p.direccion,' ',u.glsUbigeo,' ',d.glsUbigeo) as direccion,p.GlsPersona,p.direccionEntrega " & _
               "FROM personas p LEFT JOIN ubigeo u ON P.idDistrito = u.idDistrito LEFT JOIN ubigeo d ON left(u.idDistrito,2) = d.idDpto AND d.idProv = '00' AND d.idDist = '00' " & _
               "Where p.idPersona = '" & codCliente & "'"
               
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rst.EOF Then
            
            SRptBus(1) = "" & rst.Fields("GlsPersona")
            SRptBus(2) = "" & rst.Fields("ruc")
            SRptBus(3) = "" & rst.Fields("direccion")
        
        End If
        
        If rst.State = 1 Then rst.Close
        Set rst = Nothing
        
        G.Dataset.Close
        G.Dataset.Active = False
        'Set g = Nothing
        'Unload Me
        Me.Hide
    End If
    

End Select

If rst.State = 1 Then rst.Close
Set rst = Nothing
End Sub

Private Sub gTiendas_OnDblClick()
devolverResultado False
End Sub

Private Sub Text1_GotFocus()
Unload Me
End Sub

Private Sub TxtBusq_Change()
 If EsNuevo = False Then
    If glsEnterAyudaClientes = False Then
        CmdBusq_Click
    End If
 End If
End Sub

Private Sub TxtBusq_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDown Then G.SetFocus
 If EsNuevo = True Then TxtBusq.SelStart = Len(TxtBusq.Text) + 1
 If KeyCode = 27 Then
    Unload Me
 End If
End Sub


Private Sub TxtBusq_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    CmdBusq_Click
    If G.Count > 1 Then G.SetFocus
 End If
 EsNuevo = False
End Sub

Private Sub fill(ByRef StrMsgError As String)
Dim rsdatos                     As New ADODB.Recordset
On Error GoTo Err

'If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
'    If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
'        cadBusqueda = "v.ruc"
'    Else
'        cadBusqueda = "ruc"
'    End If
'Else
    cadBusqueda = "ruc"
'End If

If Trim(TxtBusq.Text) <> "" Then
    
    sqlCond = sqlBus + " like '%" & Trim(TxtBusq.Text) & "%' OR idClienteInterno like '%" & Trim(TxtBusq.Text) & "%' " & _
    "OR " & cadBusqueda & " like '%" & Trim(TxtBusq.Text) & "%' OR GlsNombreComercial like '%" & Trim(TxtBusq.Text) & "%' " & SqlAdic & " order by 1"
    
Else
'    If "" & traerCampo("Parametros", "ValParametro", "GlsParametro", "LISTA_AYUDA_CLIENTES", True) = "1" Then
'        sqlCond = sqlBus + " Like '%%'"
'    Else
        sqlCond = sqlBus + " = ''"
'    End If
End If

If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open sqlCond, Cn, adOpenStatic, adLockOptimistic
    
Set G.DataSource = rsdatos

'With G
'    .DefaultFields = False
'    .Dataset.ADODataset.ConnectionString = strcn '''Cn
'    .Dataset.ADODataset.CursorLocation = clUseClient
'    .Dataset.Active = False
'    .Dataset.ADODataset.CommandText = sqlCond
'    .Dataset.DisableControls
'    .Dataset.Active = True
'    .KeyField = "cod"
'End With

listaTiendas StrMsgError
If StrMsgError <> "" Then GoTo Err

LblReg.Caption = "(" + Format(G.Count, "0") + ")Registros"

Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Public Sub Execute(ByRef TextBox1 As Object, ByRef TextBox2 As Object, ByRef TextRUC As Object, ByRef TextDireccion As Object, ByRef TextCodtienda As Object, strParAdic As String, inddirll As Boolean)
    Dim StrMsgError As String
    Dim intI As Integer
    
    pblnAceptar = False
    MousePointer = 0
    
    SRptBus(0) = ""
    
'    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
'       If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
'          cadRelacion = "Inner Join personas v On c.idVendedorCampo = v.idPersona  "
'          cadCondicion = "c.idVendedorCampo ='" & glsUser & "' And "
'       Else
'          cadRelacion = ""
'          cadCondicion = ""
'       End If
'    Else
          cadRelacion = ""
          cadCondicion = ""
'    End If
    
    SqlAdic = strParAdic
    
    sqlBus = setSql()
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Me.Show vbModal
'    If SRptBus(0) <> "" Then
'        TextBox1.Text = SRptBus(0)
'        TextBox2.Text = SRptBus(1)
'        TextRUC.Text = SRptBus(2)
'        TextDireccion.Text = SRptBus(3)
'        TextCodtienda.Text = SRptBus(4)
'    End If
    
    If inddirll = True Then 'Para Apimas solo recupera el codigo de la tienda y la direcciomn la tienda
        If SRptBus(0) <> "" Then
             If Not TextDireccion Is Nothing Then TextDireccion.Text = SRptBus(3)
             TextCodtienda.Text = SRptBus(4)
        End If
    Else
        If SRptBus(0) <> "" Then
            TextBox1.Text = SRptBus(0)
            TextBox2.Text = SRptBus(1)
            TextRUC.Text = SRptBus(2)
            TextDireccion.Text = SRptBus(3)
            TextCodtienda.Text = SRptBus(4)
        End If
    End If
    
    Set G.DataSource = Nothing
   
    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Function setSql() As String
    
    setSql = ""
'setSql = "SELECT personas.idPersona cod ,personas.GlsPersona, des, personas.ruc, personas.Telefonos FROM personas,clientes WHERE personas.idPersona = clientes.idCliente AND Clientes.idEmpresa = '" & glsEmpresa & "' AND GlsPersona "
    
'    If Trim("" & traerCampo("parametros", "ValParametro", "GlsParametro", "VISUALIZA_UBIGEO", True)) = "N" Then
    
'        setSql = "SELECT p.idPersona cod ,p.GlsPersona des, p.ruc, p.Telefonos,p.direccion as direccion,P.GlsNombreComercial " & _
'                 "FROM personas p INNER JOIN clientes c ON c.idEmpresa = '" & glsEmpresa & "' And p.idPersona = c.idCliente " & _
'                 "Left Join ubigeo u On P.idDistrito = u.idDistrito  AND p.idPais = u.idPais " & _
'                 "Left Join ubigeo d On left(u.idDistrito,2) = d.idDpto And d.idProv = '00' And d.idDist = '00'   AND u.idPais = d.idPais " & _
'                  cadRelacion & _
'                 "WHERE " & strVacio & " " & cadCondicion & "  p.GlsPersona "
'
'    Else
    
        setSql = "SELECT p.idPersona cod,p.GlsPersona des,p.ruc,p.Telefonos," & _
                 "(p.direccion+' '+isnull(u.glsUbigeo,'')+' '+ isnull(d.glsUbigeo,'')) direccion,P.GlsNombreComercial " & _
                 "FROM personas p INNER JOIN clientes c ON c.idEmpresa = '" & glsEmpresa & "' And p.idPersona = c.idCliente " & _
                 "Left Join ubigeo u On P.idDistrito = u.idDistrito AND p.idPais = u.idPais  " & _
                 "Left Join ubigeo d On left(u.idDistrito,2) = d.idDpto And d.idProv = '00' And d.idDist = '00' AND u.idPais = d.idPais " & _
                  cadRelacion & _
                 "WHERE " & strVacio & " " & cadCondicion & "  p.GlsPersona "
                 
'    End If
    
End Function

Public Sub ExecuteReturnText(ByRef strCod As String, ByRef StrDes As String, strParAdic As String)
    Dim StrMsgError As String
    Dim intI As Integer
    
    On Error GoTo Err
    
    pblnAceptar = False
    MousePointer = 0
    
    SRptBus(0) = ""
    
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
        If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
           cadRelacion = "Inner Join personas v On c.idVendedorCampo = v.idPersona  "
           cadCondicion = "c.idVendedorCampo ='" & glsUser & "' And "
        Else
           cadRelacion = ""
           cadCondicion = ""
        End If
     Else
           cadRelacion = ""
           cadCondicion = ""
     End If
    
    SqlAdic = strParAdic
    
    sqlBus = setSql()
    
    fill StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        StrDes = SRptBus(1)
    End If
    
    Set G.DataSource = Nothing
   
    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub ExecuteKeyascii(ByVal KeyAscii As Integer, ByRef TextBox1 As Object, ByRef TextBox2 As Object, ByRef TextRUC As Object, ByRef TextDireccion As Object, strParAdic As String)
    Dim StrMsgError As String
    
    On Error GoTo Err
    
    MousePointer = 0
    
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    SRptBus(0) = ""
    
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
       If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
          cadRelacion = "Inner Join personas v On c.idVendedorCampo = v.idPersona  "
          cadCondicion = "c.idVendedorCampo ='" & glsUser & "' And "
       Else
          cadRelacion = ""
          cadCondicion = ""
       End If
    Else
          cadRelacion = ""
          cadCondicion = ""
    End If
    
    SqlAdic = strParAdic
    
    sqlBus = setSql()
    
''    fill strMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        TextBox1.Text = SRptBus(0)
        TextBox2.Text = SRptBus(1)
        TextRUC.Text = SRptBus(2)
        TextDireccion.Text = SRptBus(3)
    End If
    
    Set G.DataSource = Nothing
   
    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub ExecuteKeyasciiReturnText(ByVal KeyAscii As Integer, ByRef strCod As String, ByRef StrDes As String, strParAdic As String)
    Dim StrMsgError As String
    Dim intI As Integer
    
    On Error GoTo Err
    
    pblnAceptar = False
    MousePointer = 0
    
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    SRptBus(0) = ""
    
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
        If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
           cadRelacion = "Inner Join personas v On c.idVendedorCampo = v.idPersona  "
           cadCondicion = "c.idVendedorCampo ='" & glsUser & "' And "
        Else
           cadRelacion = ""
           cadCondicion = ""
        End If
     Else
           cadRelacion = ""
           cadCondicion = ""
     End If
    
    SqlAdic = strParAdic
    
    sqlBus = setSql()
    
    fill StrMsgError
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        StrDes = SRptBus(1)
    End If
    
    Set G.DataSource = Nothing
   
    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub


Private Sub listaTiendas(ByRef StrMsgError As String)
'Dim rsdatos                     As New ADODB.Recordset
'On Error GoTo Err

'csql = "SELECT item,GlsNombre,GlsDireccion,GlsTelefonos,idtdacli " & _
'           "FROM tiendascliente " & _
'           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
'             "AND idPersona = '" & G.Columns.ColumnByFieldName("cod").Value & "'"
'
'
'If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
'rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
'
'Set gTiendas.DataSource = rsdatos
'
''With gTiendas
''    .DefaultFields = False
''    .Dataset.ADODataset.ConnectionString = strcn '''Cn
''    .Dataset.ADODataset.CursorLocation = clUseClient
''    .Dataset.Active = False
''    .Dataset.ADODataset.CommandText = csql
''    .Dataset.DisableControls
''    .Dataset.Active = True
''    .KeyField = "item"
''End With
'
'Exit Sub
'Err:
'If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub devolverResultado(ByVal indVarCabecera As Boolean)

   SRptBus(0) = G.Columns.ColumnByFieldName("cod").Value
   SRptBus(1) = G.Columns.ColumnByFieldName("des").Value
   SRptBus(2) = G.Columns.ColumnByFieldName("ruc").Value
   SRptBus(3) = G.Columns.ColumnByFieldName("Direccion").Value
   If indVarCabecera Then
    'SRptBus(3) = g.Columns.ColumnByFieldName("Direccion").Value
    SRptBus(4) = ""
   Else
    'SRptBus(3) = gTiendas.Columns.ColumnByFieldName("GlsDireccion").Value
    SRptBus(4) = gTiendas.Columns.ColumnByFieldName("idtdacli").Value
   End If
   
   
   G.Dataset.Close
   G.Dataset.Active = False
   'Set g = Nothing
   'Unload Me
   Me.Hide


End Sub
