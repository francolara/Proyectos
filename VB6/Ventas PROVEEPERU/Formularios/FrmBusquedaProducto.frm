VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form FrmBusquedaProducto 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda"
   ClientHeight    =   5985
   ClientLeft      =   7725
   ClientTop       =   7755
   ClientWidth     =   11130
   Icon            =   "FrmBusquedaProducto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdbusq 
      Height          =   315
      Left            =   10620
      Picture         =   "FrmBusquedaProducto.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   180
      Width           =   390
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   75
      TabIndex        =   5
      Top             =   585
      Width           =   10935
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   4530
         Left            =   90
         OleObjectBlob   =   "FrmBusquedaProducto.frx":0396
         TabIndex        =   1
         Top             =   150
         Width           =   10710
      End
   End
   Begin VB.TextBox TxtBusq 
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
      Left            =   1125
      MaxLength       =   255
      TabIndex        =   0
      Top             =   165
      Width           =   9450
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
      TabIndex        =   6
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
      Top             =   5490
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
      Left            =   9090
      TabIndex        =   4
      Top             =   5490
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
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "FrmBusquedaProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Private SqlAdic As String
Private sqlBus As String
Private sqlCond As String
Private SRptBus(2) As String
Private strAyuda As String
Private EsNuevo As Boolean

Private Sub CmdBusq_Click()
    
    fill

End Sub

Private Sub Form_Deactivate()

    If EsNuevo = False Then
        TxtBusq.Text = ""
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    
    Me.Caption = "Búsqueda"
    Me.top = 0
    Me.left = 0
    ConfGrid g, False, False, False, False
    EsNuevo = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    TxtBusq.Text = ""

End Sub

Private Sub g_OnDblClick()
    
    g_OnKeyDown 13, 1

End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error Resume Next

    Select Case KeyCode
        Case 13:
            SRptBus(0) = g.Columns.ColumnByFieldName("cod").Value
            SRptBus(1) = g.Columns.ColumnByFieldName("des").Value
            g.Dataset.Close
            g.Dataset.Active = False
            Me.Hide
        Case 27
            Text1.SetFocus
    End Select

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
 
    If KeyCode = vbKeyDown Then g.SetFocus
    
    If EsNuevo = True Then TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    If KeyCode = 45 Then g_OnKeyDown KeyCode, 1
    
    If KeyCode = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub TxtBusq_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
        CmdBusq_Click
        If g.Count > 1 Then g.SetFocus
    End If
    EsNuevo = False
    
End Sub

Private Sub fill()
Dim rsdatos                     As New ADODB.Recordset

sqlCond = Replace(sqlBus, "%X%", "%" & Trim(TxtBusq.Text) & "%") & SqlAdic & " order by 1"
    
If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open sqlCond, Cn, adOpenStatic, adLockOptimistic
    
Set g.DataSource = rsdatos
    
    
'    With g
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = sqlCond
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "cod"
'    End With
    
    LblReg.Caption = "(" & Format(g.Count, "0") & ")Registros"

End Sub

Public Sub Execute(strParAyuda As String, ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
Dim StrMsgError As String
Dim intI As Integer
    
    pblnAceptar = False
    MousePointer = 0
    
    SRptBus(0) = ""
    SqlAdic = strParAdic
    sqlBus = setSql(strParAyuda)
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        TextBox1.Text = SRptBus(0)
        TextBox2.Text = SRptBus(1)
    End If
    Set g.DataSource = Nothing
    Unload Me

End Sub

Private Function setSql(strParAyuda As String) As String
Dim CCodProducto            As String

    CCodProducto = IIf(leeParametro("VIZUALIZA_CODIGO_RAPIDO") = "S", "CodigoRapido", "IdProducto")
    Select Case UCase(strParAyuda)
         Case "PRODUCTOS": setSql = "SELECT " & CCodProducto & " cod,GlsProducto des ,idFabricante,GlsUm From Productos p Inner Join  UnidadMedida u   On p.idUMCompra = u.idUm Where idEmpresa = '" & glsEmpresa & "' AND (GlsProducto like '%X%' or " & CCodProducto & " like '%X%' or idFabricante like '%X%' Or GlsUm like '%X%'  ) "
    End Select

End Function

Public Sub ExecuteReturnText(strParAyuda As String, ByRef strCod As String, ByRef strDes As String, strParAdic As String)
Dim StrMsgError As String
Dim intI As Integer
    
    pblnAceptar = False
    MousePointer = 0
    
    If glsVisualizaCodFab = "N" Then
        g.Columns.ColumnByFieldName("IdFabricante").Visible = False
    End If
    
    SRptBus(0) = ""
    SqlAdic = strParAdic
    sqlBus = setSql(strParAyuda)
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
    End If
    Set g.DataSource = Nothing
    Unload Me

End Sub

Public Sub ExecuteKeyascii(ByVal KeyAscii As Integer, strParAyuda As String, ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
Dim StrMsgError As String
    
    MousePointer = 0
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    SRptBus(0) = ""
    SqlAdic = strParAdic
    sqlBus = setSql(strParAyuda)
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        TextBox1.Text = SRptBus(0)
        TextBox2.Text = SRptBus(1)
    End If
    Set g.DataSource = Nothing
    Unload Me

End Sub

Public Sub ExecuteKeyasciiReturnText(ByVal KeyAscii As Integer, strParAyuda As String, ByRef strCod As String, ByRef strDes As String, strParAdic As String)
Dim StrMsgError As String
Dim intI As Integer
    
    pblnAceptar = False
    MousePointer = 0
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    SRptBus(0) = ""
    SqlAdic = strParAdic
    sqlBus = setSql(strParAyuda)
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
    End If
    Set g.DataSource = Nothing
    Unload Me

End Sub
