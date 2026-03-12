VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmAyudaPlanCuenta 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   8400
   ClientLeft      =   1260
   ClientTop       =   1530
   ClientWidth     =   10845
   Icon            =   "FrmAyudaPlanCuenta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   10845
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7245
      Left            =   75
      TabIndex        =   5
      Top             =   495
      Width           =   10710
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   6960
         Left            =   75
         OleObjectBlob   =   "FrmAyudaPlanCuenta.frx":000C
         TabIndex        =   1
         Top             =   150
         Width           =   10485
      End
   End
   Begin VB.TextBox TxtBusq 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   990
      MaxLength       =   255
      TabIndex        =   0
      Top             =   120
      Width           =   8370
   End
   Begin CATControls.CATTextBox txt_Ano 
      Height          =   285
      Left            =   9900
      TabIndex        =   6
      Top             =   135
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   503
      BackColor       =   16777215
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
      Container       =   "FrmAyudaPlanCuenta.frx":205C
      Estilo          =   3
      Vacio           =   -1  'True
      EnterTab        =   -1  'True
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Año"
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
      Index           =   4
      Left            =   9495
      TabIndex        =   7
      Top             =   180
      Width           =   300
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
      Left            =   7560
      TabIndex        =   4
      Top             =   7875
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
      TabIndex        =   3
      Top             =   7875
      Width           =   3120
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   135
      TabIndex        =   2
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "FrmAyudaPlanCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst            As New ADODB.Recordset
Private SqlAdic    As String
Private sqlBus     As String
Private sqlCond    As String
Private SRptBus(2) As String
Private strAyuda   As String
Private EsNuevo    As Boolean
Private conexion   As String
Private CadenaAnno As String

Private Sub Form_Activate()
On Error Resume Next

    TxtBusq.SetFocus

End Sub

Private Sub Form_Deactivate()

    SqlAdic = ""
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

    TxtBusq.Enabled = True
    Me.Caption = "Búsqueda"
    ConfGrid G, False, False, False, False
    EsNuevo = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SqlAdic = ""
    TxtBusq.Text = ""
    
End Sub

Private Sub g_OnDblClick()
    
    g_OnKeyDown 13, 1

End Sub

Private Sub g_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error Resume Next

    Select Case KeyCode
        Case 13:
            SRptBus(0) = Trim(G.Columns.ColumnByFieldName("cod").Value)
            SRptBus(1) = Trim(G.Columns.ColumnByFieldName("des").Value)
            G.Dataset.Close
            G.Dataset.Active = False
            Me.Hide
        Case 27
            Text1.SetFocus
    End Select
    
End Sub
 
 
Private Sub txt_Ano_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    Fill conexion
 End If
End Sub

Private Sub TxtBusq_Change()
    
    If EsNuevo = False Then
        Fill conexion
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
        Fill conexion
        If G.Count > 1 Then G.SetFocus
    End If
    EsNuevo = False

End Sub

Private Sub Fill(cone As String)

    If TxtBusq.Enabled = True Then
        'sqlCond = sqlBus & " like '" & Trim(TxtBusq.Text) & "%' or A.glsNombreCuenta like '" & Trim(TxtBusq.Text) & "%' " & CadenaAnno & " " & SqlAdic & ") order by 1"
        sqlCond = sqlBus & " like '" & Trim(TxtBusq.Text) & "%' or A.glsNombreCuenta like '%" & Trim(TxtBusq.Text) & "%'  " & SqlAdic & ") AND a.idAnno =  '" & txt_Ano.Text & "' order by 1"
    Else
        sqlCond = sqlBus
    End If
    
    With G
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cone
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sqlCond
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "cod"
    End With
    LblReg.Caption = "(" + Format(G.Count, "0") + ")Registros"

End Sub

Public Sub Execute(strParAyuda As String, ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
Dim StrMsgError As String
Dim intI As Integer

    pblnAceptar = False
    MousePointer = 0
    
    SRptBus(0) = ""
    SqlAdic = strParAdic
    sqlBus = setSql(strParAyuda)
    Fill conexion
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        TextBox1.Text = SRptBus(0)
        TextBox2.Text = SRptBus(1)
    End If
    
End Sub

Private Function setSql(strParAyuda As String) As String
    
    setSql = ""
    setSql = "SELECT A.idCtacontable as cod, A.glsNombreCuenta as des, A.gradocuenta as grado, b.CuentaDestino1 as CtaDestino " & _
            "FROM PlanCuentas A LEFT JOIN CuentaDestino B ON a.idanno = b.idanno and a.idCtacontable = b.idCtacontable and  a.idempresa = b.idempresa " & _
            "WHERE a.idEmpresa = '" & glsEmpresa & "'  " & _
            " AND  (A.idCtacontable "
    
End Function

Public Sub ExecuteReturnText(cone As String, strParAyuda As String, ByRef StrCod As String, ByRef StrDes As String, strParAdic As String, strAnno As String)
Dim StrMsgError As String
Dim intI As Integer
    
    pblnAceptar = False
    MousePointer = 0
    SRptBus(0) = ""
    SqlAdic = strParAdic
    
    txt_Ano.Text = strAnno
    
    If strAnno = "2010" Then
        CadenaAnno = " And a.IdAnno in ('2010') "
    Else
        CadenaAnno = " And a.IdAnno not in ('2010') "
    End If
  
    sqlBus = setSql(strParAyuda)
    conexion = cone
    Fill cone
    
    Me.Show vbModal
    sw_entro_ayuda = False
    If SRptBus(0) <> "" Then
        sw_entro_ayuda = True
        StrCod = SRptBus(0)
        StrDes = SRptBus(1)
    End If
    
End Sub

Public Sub ExecuteReturnTextBL(ByRef StrCod As String, varValores As String)
Dim StrMsgError As String
Dim strBLs() As String
Dim strSelBLs As String
Dim intI As Integer

    pblnAceptar = False
    MousePointer = 0
    strBLs = Split(varValores, ",")
    strSelBLs = ""
    For intI = 0 To UBound(strBLs)
        If intI > 0 Then strSelBLs = strSelBLs + " UNION ALL "
        strSelBLs = strSelBLs & "SELECT " & CStr(intI + 1) & " as cod,'" & strBLs(intI) & "' as des"
    Next
        
    SRptBus(0) = ""
    sqlBus = strSelBLs
    TxtBusq.Enabled = False
    Fill conexion
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        StrCod = SRptBus(1)
    End If
    TxtBusq.Enabled = True

End Sub

Public Sub ExecuteKeyacii(ByVal KeyAscii As Integer, strParAyuda As String, ByRef TextBox1 As Object, ByRef TextBox2 As Object, strParAdic As String)
Dim StrMsgError As String
    
    MousePointer = 0
    
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    SRptBus(0) = ""
    SqlAdic = strParAdic
    sqlBus = setSql(strParAyuda)
    Fill conexion
    
    Me.Show vbModal
    If SRptBus(0) <> "" Then
        TextBox1.Text = SRptBus(0)
        TextBox2.Text = SRptBus(1)
    End If
    
End Sub

Public Sub ExecuteKeyaciiReturnText(ByVal KeyAscii As Integer, strParAyuda As String, ByRef StrCod As String, ByRef StrDes As String, strParAdic As String)
Dim StrMsgError As String
Dim intI As Integer

    pblnAceptar = False
    MousePointer = 7
    EsNuevo = True
    
    TxtBusq.Text = Chr(KeyAscii)
    TxtBusq.SelStart = Len(TxtBusq.Text) + 1
    
    SRptBus(0) = ""
    SqlAdic = strParAdic
    sqlBus = setSql(strParAyuda)
    Fill conexion
    
    Me.Show vbModal
    sw_entro_ayuda = False
    If SRptBus(0) <> "" Then
        sw_entro_ayuda = True
        StrCod = SRptBus(0)
        StrDes = SRptBus(1)
    End If

End Sub
