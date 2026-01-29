VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmResumenDiarioBoletas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen Diario de Boletas"
   ClientHeight    =   5550
   ClientLeft      =   3240
   ClientTop       =   2055
   ClientWidth     =   14025
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   14025
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   60
      TabIndex        =   9
      Top             =   660
      Width           =   13905
      Begin DXDBGRIDLibCtl.dxDBGrid G 
         Height          =   3885
         Left            =   60
         OleObjectBlob   =   "FrmResumenDiarioBoletas.frx":0000
         TabIndex        =   10
         Top             =   150
         Width           =   13785
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   7305
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   13935
      Begin VB.CommandButton CmbAyudaTipoDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7095
         Picture         =   "FrmResumenDiarioBoletas.frx":29F6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Visible         =   0   'False
         Width           =   390
      End
      Begin CATControls.CATTextBox TxtIdDocumento 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   210
         Width           =   915
         _ExtentX        =   1614
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
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Locked          =   -1  'True
         MaxLength       =   8
         Container       =   "FrmResumenDiarioBoletas.frx":2D80
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsDocumento 
         Height          =   315
         Left            =   2010
         TabIndex        =   3
         Top             =   225
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         BackColor       =   12648447
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
         Container       =   "FrmResumenDiarioBoletas.frx":2D9C
         Vacio           =   -1  'True
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   315
         Left            =   12555
         TabIndex        =   5
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131792897
         CurrentDate     =   38667
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   11700
         TabIndex        =   6
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc."
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   4
         Top             =   270
         Width           =   675
      End
   End
End
Attribute VB_Name = "FrmResumenDiarioBoletas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAyudaTipoDoc_Click()
On Error GoTo Err
Dim StrMsgError                         As String
    
    mostrarAyuda "DOCUMENTOS", TxtIdDocumento, TxtGlsDocumento
       
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError                         As String
Dim NItem                               As Integer
Dim strRuta                             As String
Dim CCarpeta                            As String
Dim IntFile                             As Integer
Dim strLinea                            As String
Dim RsC                                 As New ADODB.Recordset
Dim CSqlC                               As String
Dim RetVal

    CCarpeta = leeParametro("CARPETA_XML_VE")
        
    'strRuta = CCarpeta & "\" & strRUC & "-" & pRecCab("idDocumento") & "-" & pRecCab("idSerie") & "-" & pRecCab("idDocVentas") & ".xml"
    strRuta = CCarpeta & "\" & traerCampo("Empresas", "Ruc", "IdEmpresa", glsEmpresa, False) & "-RC-" & Format(getFechaSistema, "yyyymmdd") & "-1.xml"
    IntFile = FreeFile
    Open strRuta For Output As #IntFile
        
    strLinea = ""
    strLinea = strLinea & "<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?> " & vbCrLf
    strLinea = strLinea & "<SummaryDocuments " & vbCrLf
    strLinea = strLinea & "    xmlns=""urn:sunat:names:specification:ubl:peru:schema:xsd:SummaryDocuments-1"" " & vbCrLf
    strLinea = strLinea & "    xmlns:cac=""urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"" " & vbCrLf
    strLinea = strLinea & "    xmlns:cbc=""urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"" " & vbCrLf
    strLinea = strLinea & "    xmlns:ds=""http://www.w3.org/2000/09/xmldsig#"" " & vbCrLf
    strLinea = strLinea & "    xmlns:ext=""urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2"" " & vbCrLf
    strLinea = strLinea & "    xmlns:sac=""urn:sunat:names:specification:ubl:peru:schema:xsd:SunatAggregateComponents-1"" " & vbCrLf
    strLinea = strLinea & "    xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""> " & vbCrLf
    strLinea = strLinea & "    <ext:UBLExtensions> " & vbCrLf
    strLinea = strLinea & "        <ext:UBLExtension> " & vbCrLf
    strLinea = strLinea & "            <ext:ExtensionContent> " & vbCrLf
    strLinea = strLinea & "            </ext:ExtensionContent> " & vbCrLf
    strLinea = strLinea & "        </ext:UBLExtension> " & vbCrLf
    strLinea = strLinea & "    </ext:UBLExtensions> " & vbCrLf
    strLinea = strLinea & "    <cbc:UBLVersionID>2.0</cbc:UBLVersionID> " & vbCrLf
    strLinea = strLinea & "    <cbc:CustomizationID>1.0</cbc:CustomizationID> " & vbCrLf
    strLinea = strLinea & "    <cbc:ID>RC-" & Format(getFechaSistema, "yyyymmdd") & "-1</cbc:ID> " & vbCrLf
    strLinea = strLinea & "    <cbc:ReferenceDate>" & Format(DtpFecha.Value, "yyyy-mm-dd") & "</cbc:ReferenceDate> " & vbCrLf
    strLinea = strLinea & "    <cbc:IssueDate>" & Format(getFechaSistema, "yyyy-mm-dd") & "</cbc:IssueDate> " & vbCrLf
    
    strLinea = strLinea & "    <cac:Signature> " & vbCrLf
    strLinea = strLinea & "        <cbc:ID>SignatureSP</cbc:ID> " & vbCrLf
    strLinea = strLinea & "        <cac:SignatoryParty> " & vbCrLf
    strLinea = strLinea & "            <cac:PartyIdentification> " & vbCrLf
    strLinea = strLinea & "                <cbc:ID>" & traerCampo("Empresas", "Ruc", "IdEmpresa", glsEmpresa, False) & "</cbc:ID> " & vbCrLf
    strLinea = strLinea & "            </cac:PartyIdentification> " & vbCrLf
    strLinea = strLinea & "            <cac:PartyName> " & vbCrLf
    strLinea = strLinea & "                <cbc:Name><![CDATA[" & traerCampo("Empresas", "GlsEmpresa", "IdEmpresa", glsEmpresa, False) & "]]></cbc:Name> " & vbCrLf
    strLinea = strLinea & "            </cac:PartyName> " & vbCrLf
    strLinea = strLinea & "        </cac:SignatoryParty> " & vbCrLf
    strLinea = strLinea & "        <cac:DigitalSignatureAttachment> " & vbCrLf
    strLinea = strLinea & "            <cac:ExternalReference> " & vbCrLf
    strLinea = strLinea & "                <cbc:URI>#SignatureSP</cbc:URI> " & vbCrLf
    strLinea = strLinea & "            </cac:ExternalReference> " & vbCrLf
    strLinea = strLinea & "        </cac:DigitalSignatureAttachment> " & vbCrLf
    strLinea = strLinea & "    </cac:Signature> " & vbCrLf
    strLinea = strLinea & "    <cac:AccountingSupplierParty> " & vbCrLf
    strLinea = strLinea & "        <cbc:CustomerAssignedAccountID>" & traerCampo("Empresas", "Ruc", "IdEmpresa", glsEmpresa, False) & "</cbc:CustomerAssignedAccountID> " & vbCrLf
    strLinea = strLinea & "        <cbc:AdditionalAccountID>6</cbc:AdditionalAccountID> " & vbCrLf
    strLinea = strLinea & "        <cac:Party> " & vbCrLf
    strLinea = strLinea & "            <cac:PartyLegalEntity> " & vbCrLf
    strLinea = strLinea & "                <cbc:RegistrationName><![CDATA[" & traerCampo("Empresas", "GlsEmpresa", "IdEmpresa", glsEmpresa, False) & "]]></cbc:RegistrationName> " & vbCrLf
    strLinea = strLinea & "            </cac:PartyLegalEntity> " & vbCrLf
    strLinea = strLinea & "        </cac:Party> " & vbCrLf
    strLinea = strLinea & "    </cac:AccountingSupplierParty> " & vbCrLf
    
    CSqlC = "Call Spu_VEResumenDiarioBoletas('" & glsEmpresa & "','" & TxtIdDocumento.Text & "','" & Format(DtpFecha.Value, "yyyy-mm-dd") & "')"
    AbrirRecordset StrMsgError, Cn, RsC, CSqlC: If StrMsgError <> "" Then GoTo Err
    If Not RsC.EOF Then
    Do While Not RsC.EOF
        
    NItem = NItem + 1
    strLinea = strLinea & "    <sac:SummaryDocumentsLine> " & vbCrLf
    strLinea = strLinea & "        <cbc:LineID>" & NItem & "</cbc:LineID>" & vbCrLf
    strLinea = strLinea & "        <cbc:DocumentTypeCode>" & RsC.Fields("IdDocumento") & "</cbc:DocumentTypeCode> " & vbCrLf
    strLinea = strLinea & "        <sac:DocumentSerialID>" & RsC.Fields("IdSerie") & "</sac:DocumentSerialID> " & vbCrLf
    strLinea = strLinea & "        <sac:StartDocumentNumberID>" & RsC.Fields("IdDocVentasIni") & "</sac:StartDocumentNumberID> " & vbCrLf
    strLinea = strLinea & "        <sac:EndDocumentNumberID>" & RsC.Fields("IdDocVentasFin") & "</sac:EndDocumentNumberID> " & vbCrLf
    strLinea = strLinea & "        <sac:TotalAmount currencyID=""PEN"">" & RsC.Fields("TotalPrecioVenta") & "</sac:TotalAmount> " & vbCrLf
    strLinea = strLinea & "        <sac:BillingPayment> " & vbCrLf
    strLinea = strLinea & "            <cbc:PaidAmount currencyID=""PEN"">" & RsC.Fields("TotalBaseImponible") & "</cbc:PaidAmount> " & vbCrLf
    strLinea = strLinea & "            <cbc:InstructionID>01</cbc:InstructionID> " & vbCrLf
    strLinea = strLinea & "        </sac:BillingPayment> " & vbCrLf
    strLinea = strLinea & "        <sac:BillingPayment> " & vbCrLf
    strLinea = strLinea & "            <cbc:PaidAmount currencyID=""PEN"">" & RsC.Fields("Exonerado") & "</cbc:PaidAmount> " & vbCrLf
    strLinea = strLinea & "            <cbc:InstructionID>02</cbc:InstructionID> " & vbCrLf
    strLinea = strLinea & "        </sac:BillingPayment> " & vbCrLf
    strLinea = strLinea & "        <sac:BillingPayment> " & vbCrLf
    strLinea = strLinea & "            <cbc:PaidAmount currencyID=""PEN"">" & RsC.Fields("Inafecto") & "</cbc:PaidAmount> " & vbCrLf
    strLinea = strLinea & "            <cbc:InstructionID>03</cbc:InstructionID> " & vbCrLf
    strLinea = strLinea & "        </sac:BillingPayment> " & vbCrLf
    strLinea = strLinea & "        <cac:AllowanceCharge> " & vbCrLf
    strLinea = strLinea & "            <cbc:ChargeIndicator>true</cbc:ChargeIndicator> " & vbCrLf
    strLinea = strLinea & "            <cbc:Amount currencyID=""PEN"">" & RsC.Fields("OtrosCargos") & "</cbc:Amount> " & vbCrLf
    strLinea = strLinea & "        </cac:AllowanceCharge> " & vbCrLf
    'If Val("" & RsC.Fields("Isc")) > 0 Then
    strLinea = strLinea & "        <cac:TaxTotal> " & vbCrLf
    strLinea = strLinea & "            <cbc:TaxAmount currencyID=""PEN"">" & RsC.Fields("Isc") & "</cbc:TaxAmount> " & vbCrLf
    strLinea = strLinea & "            <cac:TaxSubtotal> " & vbCrLf
    strLinea = strLinea & "                <cbc:TaxAmount currencyID=""PEN"">" & RsC.Fields("Isc") & "</cbc:TaxAmount> " & vbCrLf
    strLinea = strLinea & "                <cac:TaxCategory> " & vbCrLf
    strLinea = strLinea & "                    <cac:TaxScheme> " & vbCrLf
    strLinea = strLinea & "                        <cbc:ID>2000</cbc:ID> " & vbCrLf
    strLinea = strLinea & "                        <cbc:Name>ISC</cbc:Name> " & vbCrLf
    strLinea = strLinea & "                        <cbc:TaxTypeCode>EXC</cbc:TaxTypeCode> " & vbCrLf
    strLinea = strLinea & "                    </cac:TaxScheme> " & vbCrLf
    strLinea = strLinea & "                </cac:TaxCategory> " & vbCrLf
    strLinea = strLinea & "            </cac:TaxSubtotal> " & vbCrLf
    strLinea = strLinea & "        </cac:TaxTotal> " & vbCrLf
    'End If
    'If Val("" & RsC.Fields("TotalIgvVenta")) > 0 Then
    strLinea = strLinea & "        <cac:TaxTotal> " & vbCrLf
    strLinea = strLinea & "            <cbc:TaxAmount currencyID=""PEN"">" & RsC.Fields("TotalIgvVenta") & "</cbc:TaxAmount> " & vbCrLf
    strLinea = strLinea & "            <cac:TaxSubtotal> " & vbCrLf
    strLinea = strLinea & "                <cbc:TaxAmount currencyID=""PEN"">" & RsC.Fields("TotalIgvVenta") & "</cbc:TaxAmount> " & vbCrLf
    strLinea = strLinea & "                <cac:TaxCategory> " & vbCrLf
    strLinea = strLinea & "                    <cac:TaxScheme> " & vbCrLf
    strLinea = strLinea & "                        <cbc:ID>1000</cbc:ID> " & vbCrLf
    strLinea = strLinea & "                        <cbc:Name>IGV</cbc:Name> " & vbCrLf
    strLinea = strLinea & "                        <cbc:TaxTypeCode>VAT</cbc:TaxTypeCode> " & vbCrLf
    strLinea = strLinea & "                    </cac:TaxScheme> " & vbCrLf
    strLinea = strLinea & "                </cac:TaxCategory> " & vbCrLf
    strLinea = strLinea & "            </cac:TaxSubtotal> " & vbCrLf
    strLinea = strLinea & "        </cac:TaxTotal> " & vbCrLf
    'End If
    If Val("" & RsC.Fields("OtrosTributos")) > 0 Then
    strLinea = strLinea & "        <cac:TaxTotal> " & vbCrLf
    strLinea = strLinea & "            <cbc:TaxAmount currencyID=""PEN"">" & RsC.Fields("OtrosTributos") & "</cbc:TaxAmount> " & vbCrLf
    strLinea = strLinea & "            <cac:TaxSubtotal> " & vbCrLf
    strLinea = strLinea & "                <cbc:TaxAmount currencyID=""PEN"">" & RsC.Fields("OtrosTributos") & "</cbc:TaxAmount> " & vbCrLf
    strLinea = strLinea & "                <cac:TaxCategory> " & vbCrLf
    strLinea = strLinea & "                    <cac:TaxScheme> " & vbCrLf
    strLinea = strLinea & "                        <cbc:ID>9999</cbc:ID> " & vbCrLf
    strLinea = strLinea & "                        <cbc:Name>OTROS</cbc:Name> " & vbCrLf
    strLinea = strLinea & "                        <cbc:TaxTypeCode>OTH</cbc:TaxTypeCode> " & vbCrLf
    strLinea = strLinea & "                    </cac:TaxScheme> " & vbCrLf
    strLinea = strLinea & "                </cac:TaxCategory> " & vbCrLf
    strLinea = strLinea & "            </cac:TaxSubtotal> " & vbCrLf
    strLinea = strLinea & "        </cac:TaxTotal> " & vbCrLf
    End If
    strLinea = strLinea & "    </sac:SummaryDocumentsLine> " & vbCrLf
        
    RsC.MoveNext
    Loop
    End If
    RsC.Close: Set RsC = Nothing
    
    strLinea = strLinea & "</SummaryDocuments> "
    
    Print #IntFile, strLinea
    
    Close #IntFile
    MsgBox "Se creo el archivo XML, en unos momentos se enviara al sistema SOL", vbInformation
    
    RetVal = ShellExecute(Me.hwnd, "Open", leeParametro("RUTA_EJECUTABLE_ENVIAXMLSUNAT") & "\FEapp.exe", "", "", 1)
    
    Exit Sub
Err:
    If RsC.State = 1 Then RsC.Close: Set RsC = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub dtpFecha_Change()
On Error GoTo Err
Dim StrMsgError                         As String
    
    ListaDocumentos StrMsgError
    If StrMsgError <> "" Then GoTo Err
       
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                         As String
    
    If KeyCode = 13 Then
        cmdaceptar.SetFocus
    End If
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError                         As String

    ConfGrid G, False, True, False, False
    TxtIdDocumento.Text = "03"
    DtpFecha.Value = getFechaSistema
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub g_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
On Error GoTo Err
Dim StrMsgError                         As String
    
    G.Dataset.Post
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub TxtIdDocumento_Change()
On Error GoTo Err
Dim StrMsgError                         As String
    
    TxtGlsDocumento.Text = traerCampo("Documentos", "GlsDocumento", "IdDocumento", TxtIdDocumento.Text, False)
       
    ListaDocumentos StrMsgError
    If StrMsgError <> "" Then GoTo Err
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ListaDocumentos(StrMsgError As String)
On Error GoTo Err
Dim CSqlC                       As String
    
    CSqlC = "Call Spu_VEResumenDiarioBoletas('" & glsEmpresa & "','" & TxtIdDocumento.Text & "','" & Format(DtpFecha.Value, "yyyy-mm-dd") & "')"
    
    With G
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = CSqlC
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "Item"
    End With
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub TxtIdDocumento_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError                         As String
    
    If KeyAscii = 13 Then
        DtpFecha.SetFocus
    End If
        
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
