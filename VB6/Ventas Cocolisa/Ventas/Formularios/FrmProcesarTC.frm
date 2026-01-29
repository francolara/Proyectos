VERSION 5.00
Begin VB.Form FrmProcesarTC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesar Tipos de Cambio"
   ClientHeight    =   2070
   ClientLeft      =   6345
   ClientTop       =   2325
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1575
      Width           =   1140
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2715
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   135
      TabIndex        =   4
      Top             =   135
      Width           =   5100
      Begin VB.ComboBox cbxAno 
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
         ItemData        =   "FrmProcesarTC.frx":0000
         Left            =   1710
         List            =   "FrmProcesarTC.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   2340
      End
      Begin VB.ComboBox cbxMes 
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
         ItemData        =   "FrmProcesarTC.frx":0050
         Left            =   1710
         List            =   "FrmProcesarTC.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2340
      End
      Begin VB.Label Label1 
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
         Left            =   900
         TabIndex        =   6
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   900
         TabIndex        =   5
         Top             =   765
         Width           =   300
      End
   End
End
Attribute VB_Name = "FrmProcesarTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError                             As String

    If MsgBox("Está Seguro(a) de Procesar los Tipos de Cambios ?", vbQuestion + vbYesNo, "Atención") = vbYes Then
        
        PROCESA_TC StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
    End If

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
Dim fecha As Date
Dim i As Integer

    fecha = Format(getFechaSistema, "dd/mm/yyyy")
    strAno = Format(Year(fecha), "0000")
    strMes = Format(Month(fecha), "00")
    
    cbxAno.Clear
    For i = 2008 To Val(strAno)
        cbxAno.AddItem i
    Next
    
    For i = 0 To cbxAno.ListCount - 1
        cbxAno.ListIndex = i
        If cbxAno.Text = strAno Then Exit For
    Next
    CbxMes.ListIndex = Val(strMes) - 1
    
End Sub

Private Sub PROCESA_TC(StrMsgError As String)
On Error GoTo Err
Dim CSqlC                               As String
'dim ctipo As String, cserie As String, cnumero As String, cupdate As String, CFecha As String
Dim strAno                              As String, strMes As String
Dim CParam(4)                           As String
Dim RsC                                 As New ADODB.Recordset
Dim COnline                             As String
Dim cperiodo                            As String
Dim corigen                             As String
Dim IndGenero                           As Boolean

    COnline = traerCampo("Parametros", "ValParametro", "GlsParametro", "CONTABILIDAD_ONLINE", True)
    
    If COnline = "1" Then
        RsC.Open "Select GlsParametro,ValParametro From Parametros Where IdEmpresa = '" & glsEmpresa & "'", Cn, adOpenStatic, adLockReadOnly
        Do While Not RsC.EOF
            Select Case UCase(RsC.Fields("GlsParametro"))
                Case "ORIGEN_CONTABLE": CParam(0) = "" & RsC.Fields("ValParametro")
                Case "TRANS_CONTA_CTA_40": CParam(1) = "" & RsC.Fields("ValParametro")
                Case "TRANS_CONTA_CTA_70": CParam(2) = "" & RsC.Fields("ValParametro")
                Case "AYUDA_CENTRO_COSTO_DETALLE": CParam(3) = "" & RsC.Fields("ValParametro")
            End Select
            RsC.MoveNext
        Loop
        RsC.Close: Set RsC = Nothing
    End If
    
    strAno = cbxAno.Text
    strMes = Format(CbxMes.ListIndex + 1, "00")
    
    CSqlC = "Update A Set A.TipoCambio = B.TcVenta,A.FechaPT = GETDATE(),A.IdUsuarioPT = '" & glsUser & "' FROM DocVentas A " & _
            "Inner Join TiposDeCambio B " & _
                "On A.FecEmision = B.Fecha " & _
            " " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And Year(A.FecEmision) = '" & strAno & "' And Month(A.FecEmision) = " & Val(strMes) & " " & _
            "And A.IdDocumento In('01','03','07','08','12','56') And A.TipoCambio <> B.TcVenta"
    
    Cn.Execute CSqlC
    
    CSqlC = "Update A Set A.TipoCambio = D.TcVenta,A.FechaPT = GETDATE(),A.IdUsuarioPT = '" & glsUser & "' FROM DocVentas A " & _
            "Inner Join DocReferencia B " & _
                "On A.IdEmpresa = B.IdEmpresa And A.IdDocumento = B.TipoDocOrigen And A.IdSerie = B.SerieDocOrigen And A.IdDocVentas = B.NumDocOrigen " & _
            "Inner Join DocVentas C " & _
                "On B.IdEmpresa = C.IdEmpresa And B.TipoDocReferencia = C.IdDocumento And B.SerieDocReferencia = C.IdSerie And B.NumDocReferencia = C.IdDocVentas " & _
            "Inner Join TiposDeCambio D " & _
                "On C.FecEmision = D.Fecha " & _
            " " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And Year(A.FecEmision) = '" & strAno & "' And Month(A.FecEmision) = " & Val(strMes) & " " & _
            "And A.IdDocumento In('07') And B.TipoDocReferencia In('01') And A.TipoCambio <> D.TcVenta"
    
    Cn.Execute CSqlC
    
    If COnline = "1" Then
    
        cperiodo = cbxAno.Text & Format(CbxMes.ListIndex + 1, "00")
        corigen = traerCampo("parametros", "valparametro", "glsparametro", "ORIGEN_CONTABLE", True)
        
        CSqlC = "Delete From AsientoContable " & _
                "Where IdEmpresa = '" & glsEmpresa & "' And IdPeriodo = '" & cperiodo & "' And IdOrigen = '" & corigen & "'"
        
        CnConta.Execute (CSqlC)
        
        CSqlC = "Delete From AsientoContableDetalle " & _
                "Where IdEmpresa = '" & glsEmpresa & "' And IdPeriodo = '" & cperiodo & "' And IdOrigen = '" & corigen & "'"
        
        CnConta.Execute (CSqlC)
                  
        CSqlC = "Update DocVentas " & _
                "Set IndTrasladoConta = 'N',IdComprobante = '' " & _
                "Where IdEmpresa = '" & glsEmpresa & "' And Year(FecEmision) = " & Val(strAno) & " And Month(FecEmision) = " & Val(strMes) & ""
        
        Cn.Execute (CSqlC)
        
        Dim FrmGen              As New frmAsientoContableGenerar
        
        IndGenero = False
        
        FrmGen.Genera_Internamente StrMsgError, cperiodo, IndGenero
        If StrMsgError <> "" Then GoTo Err
        
        If IndGenero Then
        
            Dim FrmTrans            As New frmAsientoContableTransferir
            
            FrmTrans.Genera_Internamente StrMsgError, cperiodo, IndGenero
            If StrMsgError <> "" Then GoTo Err
        
        End If
        
    End If
    
    MsgBox "Fin del Proceso.", vbInformation, App.Title
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
