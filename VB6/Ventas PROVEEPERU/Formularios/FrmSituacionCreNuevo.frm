VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmSituacionCreNuevo 
   Appearance      =   0  'Flat
   Caption         =   "Situación Crediticia"
   ClientHeight    =   9480
   ClientLeft      =   3945
   ClientTop       =   1185
   ClientWidth     =   10065
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   10065
   Begin VB.CommandButton CmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   420
      Left            =   4365
      TabIndex        =   11
      Top             =   8955
      Width           =   1770
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8790
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9915
      Begin VB.Frame FraRegistro 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   135
         TabIndex        =   14
         Top             =   2385
         Width           =   9645
         Begin TabDlg.SSTab SSTab1 
            Height          =   4425
            Left            =   90
            TabIndex        =   15
            Top             =   180
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   7805
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Documentos por Cobrar"
            TabPicture(0)   =   "FrmSituacionCreNuevo.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "GDocumentos"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Guías por Facturar"
            TabPicture(1)   =   "FrmSituacionCreNuevo.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "GGuiasNF"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Documentos Generados"
            TabPicture(2)   =   "FrmSituacionCreNuevo.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "GDocumentosGen"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin DXDBGRIDLibCtl.dxDBGrid GDocumentos 
               Height          =   3915
               Left            =   90
               OleObjectBlob   =   "FrmSituacionCreNuevo.frx":0054
               TabIndex        =   16
               Top             =   405
               Width           =   9270
            End
            Begin DXDBGRIDLibCtl.dxDBGrid GGuiasNF 
               Height          =   3915
               Left            =   -74910
               OleObjectBlob   =   "FrmSituacionCreNuevo.frx":29FC
               TabIndex        =   17
               Top             =   405
               Width           =   9270
            End
            Begin DXDBGRIDLibCtl.dxDBGrid GDocumentosGen 
               Height          =   3915
               Left            =   -74910
               OleObjectBlob   =   "FrmSituacionCreNuevo.frx":53A4
               TabIndex        =   18
               Top             =   405
               Width           =   9270
            End
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   2700
         TabIndex        =   9
         Top             =   7155
         Width           =   4605
         Begin VB.TextBox TxtSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0"
            Top             =   1080
            Width           =   1680
         End
         Begin CATControls.CATTextBox TxtLineaAprobada 
            Height          =   315
            Left            =   2070
            TabIndex        =   19
            Top             =   180
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            BackColor       =   12640511
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
            Locked          =   -1  'True
            Container       =   "FrmSituacionCreNuevo.frx":7D4C
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox TxtDeuda 
            Height          =   315
            Left            =   2070
            TabIndex        =   21
            Top             =   585
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            BackColor       =   12640511
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
            Locked          =   -1  'True
            Container       =   "FrmSituacionCreNuevo.frx":7D68
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label8 
            Caption         =   "Saldo Disponible"
            Height          =   195
            Left            =   135
            TabIndex        =   24
            Top             =   1110
            Width           =   1635
         End
         Begin VB.Label Label2 
            Caption         =   "Deuda Actual"
            Height          =   195
            Left            =   135
            TabIndex        =   22
            Top             =   630
            Width           =   1635
         End
         Begin VB.Label Label3 
            Caption         =   "Línea Aprobada"
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   225
            Width           =   1455
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            Caption         =   "__________________________"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   1890
            TabIndex        =   10
            Top             =   810
            Width           =   2145
         End
      End
      Begin VB.Frame FraLineaCredito 
         Appearance      =   0  'Flat
         Caption         =   " Situación Crediticia "
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   135
         TabIndex        =   1
         Top             =   180
         Width           =   9645
         Begin CATControls.CATTextBox txt_Cod_Cli_Linea 
            Height          =   315
            Left            =   1665
            TabIndex        =   2
            Top             =   315
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmSituacionCreNuevo.frx":7D84
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox tztGlsLinea 
            Height          =   315
            Left            =   2745
            TabIndex        =   3
            Top             =   315
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmSituacionCreNuevo.frx":7DA0
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtRuc_Linea 
            Height          =   315
            Left            =   1665
            TabIndex        =   4
            Top             =   675
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmSituacionCreNuevo.frx":7DBC
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtfomaPagoLinea 
            Height          =   315
            Left            =   1665
            TabIndex        =   5
            Top             =   1035
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmSituacionCreNuevo.frx":7DD8
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsEstado 
            Height          =   315
            Left            =   1665
            TabIndex        =   12
            Top             =   1395
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmSituacionCreNuevo.frx":7DF4
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtCodMoneda 
            Height          =   315
            Left            =   1665
            TabIndex        =   25
            Top             =   1755
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmSituacionCreNuevo.frx":7E10
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox TxtGlsMoneda 
            Height          =   315
            Left            =   2745
            TabIndex        =   26
            Top             =   1755
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "FrmSituacionCreNuevo.frx":7E2C
            Vacio           =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Moneda Línea"
            Height          =   210
            Left            =   315
            TabIndex        =   27
            Top             =   1800
            Width           =   1005
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   315
            TabIndex        =   13
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pago"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   315
            TabIndex        =   8
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "R.U.C."
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   315
            TabIndex        =   7
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   210
            Left            =   315
            TabIndex        =   6
            Top             =   360
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "FrmSituacionCreNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTmpDocumentos                      As String
Dim CTmpGuiasNF                         As String
Dim CTmpDocumentosGen                   As String
Dim CIdCliente                          As String
Dim CGlsCliente                         As String
Dim cruccliente                         As String

Private Sub cmdsalir_Click()
On Error GoTo Err
Dim StrMsgError                         As String

    Unload Me
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub MostrarFrom(StrMsgError As String, PIdCliente As String, PGlsCliente As String, PRucCliente As String)
On Error GoTo Err
Dim CSqlC                   As String
    
    CIdCliente = PIdCliente
    CGlsCliente = PGlsCliente
    cruccliente = PRucCliente
    
    FrmSituacionCreNuevo.Show 1
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError                         As String
    
    CTmpDocumentos = ""
    CTmpDocumentosGen = ""
    CTmpGuiasNF = ""
    
    ConfGrid GDocumentos, False, True, False, False
    ConfGrid GGuiasNF, False, True, False, False
    ConfGrid GDocumentosGen, False, True, False, False
    
    SSTab1.Tab = 0
    
    txt_Cod_Cli_Linea.Text = Trim("" & CIdCliente)
    tztGlsLinea.Text = Trim("" & CGlsCliente)
    txtRuc_Linea.Text = Trim("" & cruccliente)

    txtfomaPagoLinea.Text = Trim("" & traerCampo("formaspagos", "GlsFormaPago", "idformapago", Trim("" & traerCampo("clientes", "idFormaPago", "idCliente", Trim("" & CIdCliente), True)), True))
    
    TxtCodMoneda.Text = Trim("" & traerCampo("ControlLineaCredito", "IdMoneda", "IdCliente", Trim("" & CIdCliente), True))
        
    CargaLineaCredito StrMsgError, CIdCliente
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CargaLineaCredito(StrMsgError As String, PIdCliente As String)
On Error GoTo Err
Dim CSqlC                               As String
Dim RsC                                 As New ADODB.Recordset
Dim CPC                                 As String
    
    'Linea Actual
    CSqlC = "Select A.Linea_Actual,A.Linea_Usada,A.Saldo,A.IndSuspension,B.GlsDato " & _
            "From ControlLineaCredito A " & _
            "Inner Join Datos B " & _
                "On A.Estado = B.IdDato " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdCliente = '" & PIdCliente & "'"
    RsC.Open CSqlC, Cn, adOpenStatic, adLockReadOnly
    If Not RsC.EOF Then
        
        If Trim("" & RsC.Fields("IndSuspension")) = "1" Then
            
            TxtGlsEstado.Text = "Suspendido"
        
        Else
            
            TxtGlsEstado.Text = "" & RsC.Fields("GlsDato")
            
        End If
        
        TxtLineaAprobada.Text = Val("" & RsC.Fields("Linea_Actual"))
        TxtDeuda.Text = Val("" & RsC.Fields("Linea_Usada"))
        TxtSaldo.Text = Format(Val("" & RsC.Fields("Saldo")), "#,###,##0.00")
        
    Else
        
        TxtCodMoneda.Text = traerCampo("Parametros", "ValParametro", "GlsParametro", "MONEDA_LINEA_CREDITO", True)
        
        TxtLineaAprobada.Text = "0"
        TxtDeuda.Text = "0"
        TxtSaldo.Text = Format(Val("0"), "#,###,##0.00")
        
    End If
    
    RsC.Close: Set RsC = Nothing
    
    'Documentos
    If Len(Trim(CTmpDocumentos)) = 0 Then
        
        CPC = ComputerName
        CPC = Replace(CPC, "-", "")
        CPC = Trim(CPC)
        
        CTmpDocumentos = "TmpDocumentos" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & CPC)
        
    End If
    
    CSqlC = "Call Spu_VerificaMorosos('" & glsEmpresa & "','" & PIdCliente & "','1','" & CTmpDocumentos & "','','')"
    
    With GDocumentos
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = CSqlC
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "Nro_Comp"
    End With
    
    'Guias Por Facturar
    If Len(Trim(CTmpGuiasNF)) = 0 Then
        
        CPC = ComputerName
        CPC = Replace(CPC, "-", "")
        CPC = Trim(CPC)
        
        CTmpGuiasNF = "TmpGuiasNF" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & CPC)
        
    End If
    
    CSqlC = "Call Spu_VerificaMorosos('" & glsEmpresa & "','" & PIdCliente & "','2','','" & CTmpGuiasNF & "','')"
    
    With GGuiasNF
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = CSqlC
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "Nro_Comp"
    End With
    
    'Docuentos Generados
    If Len(Trim(CTmpDocumentosGen)) = 0 Then
        
        CPC = ComputerName
        CPC = Replace(CPC, "-", "")
        CPC = Trim(CPC)
        
        CTmpDocumentosGen = "TmpDocumentosGen" & Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & CPC)
        
    End If
    
    CSqlC = "Call Spu_VerificaMorosos('" & glsEmpresa & "','" & PIdCliente & "','3','','','" & CTmpDocumentosGen & "')"
    
    With GDocumentosGen
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = CSqlC
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "Nro_Comp"
    End With
    
    Me.Refresh
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err
Dim StrMsgError                         As String
    
    If Len(Trim(CTmpDocumentos)) > 0 Then
        
        Cn.Execute "Drop Table If Exists " & CTmpDocumentos
    
    End If
    
    If Len(Trim(CTmpGuiasNF)) > 0 Then
        
        Cn.Execute "Drop Table If Exists " & CTmpGuiasNF
    
    End If
    
    If Len(Trim(CTmpDocumentosGen)) > 0 Then
        
        Cn.Execute "Drop Table If Exists " & CTmpDocumentosGen
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodMoneda_Change()
On Error GoTo Err
Dim StrMsgError
    
    TxtGlsMoneda.Text = traerCampo("Monedas", "GlsMoneda", "IdMoneda", TxtCodMoneda.Text, False)
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtCodMoneda_Click()
On Error GoTo Err
Dim StrMsgError
    
    TxtGlsMoneda.Text = traerCampo("Monedas", "GlsMoneda", "IdMoneda", TxtCodMoneda.Text, False)
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub GDocumentos_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)

    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If

End Sub

Private Sub GDocumentos_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GDocumentos_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GGuiasNF_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)

    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If

End Sub

Private Sub GGuiasNF_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GGuiasNF_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GDocumentosGen_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)

    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If

End Sub

Private Sub GDocumentosGen_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub

Private Sub GDocumentosGen_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    
    If Column.FieldName = "ValTotal" Or Column.FieldName = "SaldoPorCobrar" Then
        Text = Format(Text, "###,###,#0.00")
    End If
    
End Sub
