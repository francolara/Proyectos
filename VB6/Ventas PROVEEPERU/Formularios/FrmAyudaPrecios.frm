VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form FrmAyudaPrecios 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Precios"
   ClientHeight    =   5625
   ClientLeft      =   2865
   ClientTop       =   2535
   ClientWidth     =   9180
   Icon            =   "FrmAyudaPrecios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9180
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   75
      TabIndex        =   5
      Top             =   675
      Width           =   9015
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   4125
         Left            =   75
         OleObjectBlob   =   "FrmAyudaPrecios.frx":000C
         TabIndex        =   6
         Top             =   150
         Width           =   8835
      End
   End
   Begin VB.CommandButton Cmdbusq 
      DownPicture     =   "FrmAyudaPrecios.frx":2968
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
      Left            =   8700
      Picture         =   "FrmAyudaPrecios.frx":2D13
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   225
      Visible         =   0   'False
      Width           =   405
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
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   0
      Top             =   225
      Width           =   7605
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
      Left            =   7200
      TabIndex        =   3
      Top             =   5040
      Width           =   1905
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
      TabIndex        =   2
      Top             =   5040
      Width           =   3120
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
      TabIndex        =   1
      Top             =   315
      Width           =   915
   End
End
Attribute VB_Name = "FrmAyudaPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Private SqlAdic As String
Private sqlBus As String
Private sqlCond As String
Private SRptBus(3) As String
Private strAyuda As String
Private EsNuevo As Boolean

Private Sub CmdBusq_Click()
    fill
End Sub

Private Sub Form_Activate()
TxtBusq.SetFocus
End Sub

Private Sub Form_Deactivate()
SqlAdic = ""
If EsNuevo = False Then
    TxtBusq.Text = ""
End If
End Sub

Private Sub Form_Load()
 Me.Caption = "Ayuda de Precios"
 ConfGrid g, False, False, False, False
  EsNuevo = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
SqlAdic = ""
TxtBusq.Text = ""
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
g_OnKeyDown 13, 1
End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Select Case KeyCode
 Case 13:
      
   SRptBus(0) = g.Columns.ColumnByFieldName("idUM").Value
   SRptBus(1) = g.Columns.ColumnByFieldName("GlsUM").Value
   SRptBus(2) = g.Columns.ColumnByFieldName("factor").Value
   
   g.Dataset.Close
   g.Dataset.Active = False
   'Set g = Nothing
   'Unload Me
   Me.Hide
End Select
End Sub

Private Sub TxtBusq_Change()
  If EsNuevo = False Then
    CmdBusq_Click
  End If
End Sub


Private Sub TxtBusq_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDown Then g.SetFocus
  If EsNuevo = True Then TxtBusq.SelStart = Len(TxtBusq.Text) + 1
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


sqlCond = sqlBus + " like '%" & Trim(TxtBusq.Text) & "%' " & SqlAdic & " order by 1"

If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open sqlCond, Cn, adOpenStatic, adLockOptimistic
    
Set g.DataSource = rsdatos

''With g
''     .DefaultFields = False
''     .Dataset.ADODataset.ConnectionString = strcn '''Cn
''
''
''    .Dataset.ADODataset.CursorLocation = clUseClient
''    .Dataset.Active = False
''    .Dataset.ADODataset.CommandText = sqlCond
''    .Dataset.DisableControls
''    .Dataset.Active = True
''    .KeyField = "idUM"
''End With

LblReg.Caption = "(" + Format(g.Count, "0") + ")Registros"
End Sub


Public Sub Execute(strCodProd As String, strCodLista As String, ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
    Dim StrMsgError As String
    Dim intI As Integer
    pblnAceptar = False
    MousePointer = 0
    
    SRptBus(0) = ""
    
    SqlAdic = strParAdic
    
    sqlBus = setSql(strCodProd, strCodLista)
    
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        TextBox1.Text = SRptBus(0)
        TextBox2.Text = SRptBus(1)
    End If
End Sub

Private Function setSql(strCodProd As String, strCodLista As String) As String
    setSql = "SELECT p.idUM,u.abreUM as GlsUM,CAST(r.factor AS DECIMAL(12,2)) AS factor,CAST(p.VVUnit AS DECIMAL(12,2)) AS VVUnit,CAST(p.IGVUnit AS DECIMAL(12,2)) AS IGVUnit,CAST(p.PVUnit AS DECIMAL(12,2)) AS PVUnit " & _
             "FROM preciosventa p,unidadMedida u, presentaciones r " & _
             "WHERE p.idUM = u.idUM " & _
               "AND p.idProducto = '" & strCodProd & "' " & _
               "AND p.idEmpresa = '" & glsEmpresa & "' " & _
               "AND p.idUM = r.idUM " & _
               "AND p.idProducto = r.idProducto " & _
               "AND r.idEmpresa = '" & glsEmpresa & "' " & _
               "AND p.idLista = '" & strCodLista & "' AND u.GlsUM "
End Function

Public Sub ExecuteReturnText(strCodProd As String, strCodLista As String, ByRef strCod As String, ByRef strDes As String, strParAdic As String, ByRef dblFactor As Double)
    Dim StrMsgError As String
    Dim intI As Integer
    pblnAceptar = False
    MousePointer = 0
    
    SRptBus(0) = ""
    
    SqlAdic = strParAdic
    
    sqlBus = setSql(strCodProd, strCodLista)
    
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
        dblFactor = Val("" & SRptBus(2))
    End If
End Sub

Public Sub ExecuteKeyasciiReturnText(ByVal KeyAscii As Integer, strCodProd As String, strCodLista As String, ByRef strCod As String, ByRef strDes As String, strParAdic As String)
    Dim StrMsgError As String
    Dim intI As Integer
    pblnAceptar = False
    MousePointer = 0
    
    SRptBus(0) = ""
    
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    SqlAdic = strParAdic
    
    sqlBus = setSql(strCodProd, strCodLista)
    
    fill
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        strCod = SRptBus(0)
        strDes = SRptBus(1)
    End If
End Sub
