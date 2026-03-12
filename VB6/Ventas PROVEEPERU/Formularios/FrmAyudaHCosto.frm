VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmAyudaHCosto 
   Caption         =   "Ayuda de Hojas de Costo"
   ClientHeight    =   5895
   ClientLeft      =   1575
   ClientTop       =   2160
   ClientWidth     =   13095
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   13095
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   " Leyenda "
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
      Height          =   705
      Left            =   6885
      TabIndex        =   4
      Top             =   -15
      Width           =   6150
      Begin VB.Label Label10 
         BackColor       =   &H000080FF&
         Height          =   150
         Left            =   4410
         TabIndex        =   12
         Top             =   345
         Width           =   285
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Facturado Parcial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4725
         TabIndex        =   11
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Facturado Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3015
         TabIndex        =   10
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label Label6 
         BackColor       =   &H008080FF&
         Height          =   150
         Left            =   2700
         TabIndex        =   9
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Terminado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1665
         TabIndex        =   8
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Height          =   150
         Left            =   1305
         TabIndex        =   7
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proceso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   450
         TabIndex        =   6
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Height          =   150
         Left            =   90
         TabIndex        =   5
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   6780
      Begin VB.TextBox txtbusqueda 
         Appearance      =   0  'Flat
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
         Left            =   1005
         TabIndex        =   0
         Top             =   240
         Width           =   3990
      End
      Begin CATControls.CATTextBox txt_Ano 
         Height          =   315
         Left            =   5760
         TabIndex        =   13
         Top             =   240
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
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
         Container       =   "FrmAyudaHCosto.frx":0000
         Estilo          =   3
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label8 
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
         Height          =   210
         Left            =   5280
         TabIndex        =   14
         Top             =   285
         Width           =   300
      End
      Begin VB.Label Label1 
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
         Height          =   210
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   735
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid G 
      Height          =   5115
      Left            =   75
      OleObjectBlob   =   "FrmAyudaHCosto.frx":001C
      TabIndex        =   1
      Top             =   720
      Width           =   12945
   End
End
Attribute VB_Name = "FrmAyudaHCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LColor, LColor1
Dim CIdArea                             As String
Dim CIndTipo                            As String
Dim CIdAlmacen                          As String
Dim CIdCentroCosto                      As String
Dim NIdStock                            As String
Dim NIdCon                              As String
Dim NIdFacHC                            As String
Dim CIdDocPres                          As String
Dim CFecha                              As String

Public Sub MostrarForm(strMsgError As String, PIdArea As String, PIndTipo As String, PIdAlmacen As String, PIdCentroCosto As String, PIdDocPres As String, PFecha As String)
On Error GoTo err


    CIdArea = PIdArea
    CIndTipo = PIndTipo
    CIdAlmacen = PIdAlmacen
    CIdDocPres = PIdDocPres
    CFecha = Format(PFecha, "yyyy-mm-dd")
    CIdCentroCosto = ""
    
    Me.Show 1
    PIdCentroCosto = CIdCentroCosto

    
    Unload Me
    
    Exit Sub

err:
    If strMsgError = "" Then strMsgError = err.Description
End Sub

Private Sub G_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
On Error GoTo err
Dim strMsgError                 As String

    If Trim(Node.Values(4)) = "P" Then
        Color = &HC0FFFF    ''vbYellow
    End If
    If Trim(Node.Values(4)) = "T" Then
        Color = LColor(8)  ''vbGreen
    End If
    If Trim(Node.Values(5)) = "1" Then
        Color = &H8080FF     ''vbRed
    End If
    If Trim(Node.Values(5)) = "2" Then
        Color = &H80FF&          ''vbNose
    End If
    FontColor = vbBlack
    
    Exit Sub

err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub G_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
On Error GoTo err
Dim strMsgError                 As String

    If Trim(Node.Values(4)) = "P" Then
        Color = &HC0FFFF    ''vbYellow
    End If
    If Trim(Node.Values(4)) = "T" Then
        Color = LColor(8)  ''vbGreen
    End If
    If Trim(Node.Values(5)) = "1" Then
        Color = &H8080FF     ''vbRed
    End If
    If Trim(Node.Values(5)) = "2" Then
        Color = &H80FF&          ''vbNose
    End If
    FontColor = vbBlack
    
    Exit Sub

err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub G_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
On Error GoTo err
Dim strMsgError                 As String

    If Trim(Node.Values(4)) = "P" Then
        Color = &HC0FFFF    ''vbYellow
    End If
    If Trim(Node.Values(4)) = "T" Then
        Color = LColor(8)  ''vbGreen
    End If
    If Trim(Node.Values(5)) = "1" Then
        Color = &H8080FF     ''vbRed
    End If
    If Trim(Node.Values(5)) = "2" Then
        Color = &H80FF&          ''vbNose
    End If
    FontColor = vbBlack
    
    Exit Sub

err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub G_OnDblClick()
On Error GoTo err
Dim strMsgError     As String

    CIdCentroCosto = "" & G.Columns.ColumnByFieldName("IdCentroCosto").Value
    Me.Hide
    
    Exit Sub

err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub G_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo err
Dim strMsgError     As String

    With G.Dataset
        If G.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                G.m.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
    End With
    
    Exit Sub
    
err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub Fill(strMsgError As String)
On Error GoTo err
Dim CSqlC                           As String
Dim RsC                             As New ADODB.Recordset
Dim nPC                             As String
Dim rsTemp1                         As New ADODB.Recordset
Dim rspa1                           As New ADODB.Recordset
    
    
    nPC = ComputerName
    nPC = Replace(nPC, "-", "")
    nPC = Trim(nPC)
    
    Select Case CIndTipo
        Case "S" 'Vale de Salida
            
            If NIdStock = "" Then
                NIdStock = Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_TS")
            End If
            If NIdCon = "" Then
                NIdCon = Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_C")
            End If
            If NIdFacHC = "" Then
                NIdFacHC = Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_FHC")
            End If
            
            Set rsTemp1 = DataProcedimiento("Spu_EliminaTemporales", strMsgError, NIdStock, NIdCon, "0a", "0a", "0a", NIdFacHC)
            If strMsgError <> "" Then GoTo err
            
            Set rsTemp1 = DataProcedimiento("Spu_CalculaStock", strMsgError, glsEmpresa, NIdStock, "", 0, "", "", "", CIdAlmacen, CFecha, "")
            If strMsgError <> "" Then GoTo err
             
            Set rsTemp1 = DataProcedimiento("Spu_CalculaConsumo", strMsgError, glsEmpresa, NIdCon, "", 0, "", "", "", "", CIdArea, "")
            If strMsgError <> "" Then GoTo err
            
            Set rsTemp1 = DataProcedimiento("Spu_CalculaFacturacionPorHC", strMsgError, glsEmpresa, NIdFacHC, "", 0, "", "", "", "", CIdArea)
            If strMsgError <> "" Then GoTo err
            
            Set rsTemp1 = DataProcedimiento("Spu_AyudaHCostosValeSalida", strMsgError, glsEmpresa, CIdArea, CIdAlmacen, NIdStock, NIdCon, NIdFacHC, "%" & txtbusqueda.Text & "%", txt_Ano.Text)
            If strMsgError <> "" Then GoTo err
         
            CSqlC = "Call Spu_AyudaHCostosValeSalida('" & glsEmpresa & "','" & CIdArea & "','" & CIdAlmacen & "'," & NIdStock & "," & NIdCon & "," & NIdFacHC & ",'" & "%" & txtbusqueda.Text & "%" & "','" & txt_Ano.Text & "')"

        Case "I" 'Vale de Ingreso
             
            If NIdCon = "" Then
                NIdCon = Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_C")
            End If
            
            If NIdFacHC = "" Then
                NIdFacHC = Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_FHC")
            End If
            
            Set rsTemp1 = DataProcedimiento("Spu_EliminaTemporales", strMsgError, "0a", NIdCon, "0a", "0a", "0a", NIdFacHC)
            If strMsgError <> "" Then GoTo err
             
            Set rsTemp1 = DataProcedimiento("Spu_CalculaConsumo", strMsgError, glsEmpresa, NIdCon, "", 0, "", "", "", "", CIdArea, "")
            If strMsgError <> "" Then GoTo err
            
            Set rsTemp1 = DataProcedimiento("Spu_CalculaFacturacionPorHC", strMsgError, glsEmpresa, NIdFacHC, "", 0, "", "", "", "", CIdArea)
            If strMsgError <> "" Then GoTo err
            
            Set rsTemp1 = DataProcedimiento("Spu_AyudaHCostosValeIngreso", strMsgError, glsEmpresa, CIdArea, NIdCon, NIdFacHC, "%" & txtbusqueda.Text & "%", txt_Ano.Text)
            If strMsgError <> "" Then GoTo err
          
            CSqlC = "Call Spu_AyudaHCostosValeIngreso('" & glsEmpresa & "','" & CIdArea & "'," & NIdCon & "," & NIdFacHC & ",'" & "%" & txtbusqueda.Text & "%" & "','" & txt_Ano.Text & "')"

        Case "O" 'Otros
        
            If NIdFacHC = "" Then
                NIdFacHC = Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_C")
            End If
            
            Set rsTemp1 = DataProcedimiento("Spu_EliminaTemporales", strMsgError, "0a", "0a", "0a", "0a", "0a", NIdFacHC)
            If strMsgError <> "" Then GoTo err
            
            Set rsTemp1 = DataProcedimiento("Spu_CalculaFacturacionPorHC", strMsgError, glsEmpresa, NIdFacHC, "", 0, "", "", "", "", CIdArea)
            If strMsgError <> "" Then GoTo err
            
            Set rsTemp1 = DataProcedimiento("Spu_AyudaHCostosValeOtros", strMsgError, glsEmpresa, CIdArea, NIdFacHC, CIdDocPres, "%" & txtbusqueda.Text & "%", txt_Ano.Text)
            If strMsgError <> "" Then GoTo err
                     
    End Select
 
    
    With G

        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = rsTemp1.Source
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "IdCentroCosto"

    End With

    
    Exit Sub

err:
    If strMsgError = "" Then strMsgError = err.Description
End Sub

Private Sub G_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error GoTo err
Dim strMsgError     As String

    Select Case KeyCode
        Case 13:
            G_OnDblClick
        Case 27:
            CIdCentroCosto = ""
            Me.Hide
    End Select
    
    Exit Sub
    
err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo err
Dim strMsgError     As String

    ConfGrid G, False, False, False, False

    LColor1 = Array(RGB(255, 200, 200), vbWhite, RGB(200, 255, 200), RGB(200, 200, 255))
    LColor = Array(vbRed, vbWhite, vbGreen, vbBlue, vbYellow, RGB(0, 255, 255), RGB(255, 0, 0), RGB(255, 200, 200), RGB(200, 255, 200), RGB(200, 200, 255), vbBlack)
    
    txt_Ano.Text = Year(getFechaSistema)
    
    Fill strMsgError
    If strMsgError <> "" Then GoTo err
    
    Exit Sub
    
err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err
Dim strMsgError     As String
Dim rsEli           As New ADODB.Recordset
        
 
    NIdStock = IIf(NIdStock = "", "0a", NIdStock)
    NIdCon = IIf(NIdCon = "", "0a", NIdCon)
    NIdFacHC = IIf(NIdFacHC = "", "0a", NIdFacHC)
    
    Set rsEli = DataProcedimiento("Spu_EliminaTemporales", strMsgError, NIdStock, NIdCon, "0a", "0a", "0a", NIdFacHC)
    If strMsgError <> "" Then GoTo err
    
    NIdStock = ""
    NIdCon = ""
    NIdFacHC = ""
    
    G.Dataset.Close
    
    Exit Sub
    
err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub txt_Ano_KeyPress(KeyAscii As Integer)
On Error GoTo err
Dim strMsgError     As String

    If KeyAscii = 13 Then
        Fill strMsgError
        If strMsgError <> "" Then GoTo err
    End If
    
    Exit Sub
    
err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
Dim strMsgError     As String

    If KeyCode = 40 Then
        G.Columns.FocusedIndex = 2
        G.SetFocus
    End If
    
    Exit Sub
    
err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
On Error GoTo err
Dim strMsgError     As String

    If KeyAscii = 13 Then
        Fill strMsgError
        If strMsgError <> "" Then GoTo err
    End If
    
    Exit Sub
    
err:
    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub
