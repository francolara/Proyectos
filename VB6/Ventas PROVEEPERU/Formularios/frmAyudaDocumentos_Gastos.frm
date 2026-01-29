VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmAyudaDocumentos_Gastos 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de Documentos - Intereses "
   ClientHeight    =   7050
   ClientLeft      =   4020
   ClientTop       =   1740
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   75
      Top             =   1020
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
            Picture         =   "frmAyudaDocumentos_Gastos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":317E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":3518
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":396A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":3D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAyudaDocumentos_Gastos.frx":4716
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1164
      ButtonWidth     =   1561
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Procesar"
            Object.ToolTipText     =   "Procesar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actualizar"
            Object.ToolTipText     =   "Actualizar"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   75
      TabIndex        =   0
      Top             =   675
      Width           =   8850
      Begin VB.CommandButton cmbAyudaCliente 
         Height          =   375
         Left            =   8415
         Picture         =   "frmAyudaDocumentos_Gastos.frx":4DE8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   420
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Tag             =   "TidCtaCorriente"
         Top             =   450
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
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
         MaxLength       =   8
         Container       =   "frmAyudaDocumentos_Gastos.frx":5172
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2175
         TabIndex        =   3
         Top             =   450
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "frmAyudaDocumentos_Gastos.frx":518E
         Vacio           =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtp_Fecha 
         Height          =   315
         Left            =   6345
         TabIndex        =   4
         Tag             =   "FFecDeposito"
         Top             =   795
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         _Version        =   393216
         Format          =   42795009
         CurrentDate     =   38955
      End
      Begin CATControls.CATTextBox txtGls_RUC 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Tag             =   "TNroDeposito"
         Top             =   795
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
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
         MaxLength       =   11
         Container       =   "frmAyudaDocumentos_Gastos.frx":51AA
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_TipoCambio 
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Tag             =   "NTipoCambio"
         Top             =   1140
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Container       =   "frmAyudaDocumentos_Gastos.frx":51C6
         Text            =   "0.000"
         Decimales       =   3
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Direccion 
         Height          =   285
         Left            =   3150
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   705
         _ExtentX        =   1244
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
         MaxLength       =   8
         Container       =   "frmAyudaDocumentos_Gastos.frx":51E2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Tienda 
         Height          =   285
         Left            =   3960
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
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
         MaxLength       =   8
         Container       =   "frmAyudaDocumentos_Gastos.frx":51FE
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label lbl_TC 
         Appearance      =   0  'Flat
         Caption         =   "T/C:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label lbl_FechaEmision 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   3900
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   495
         Width           =   525
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "R.U.C.:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   870
         Width           =   525
      End
   End
   Begin VB.Frame fraDetalle 
      Appearance      =   0  'Flat
      Caption         =   " Detalle "
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   90
      TabIndex        =   9
      Top             =   2385
      Width           =   8835
      Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
         Height          =   4065
         Left            =   60
         OleObjectBlob   =   "frmAyudaDocumentos_Gastos.frx":521A
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   270
         Width           =   8595
      End
   End
End
Attribute VB_Name = "frmAyudaDocumentos_Gastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsg As New ADODB.Recordset
Dim indCargandoCliente As Boolean
Dim strTipoCobro As String
Dim indResultado As Boolean
Dim strCodMoneda As String
Dim indFormIntereses As Boolean
Dim strTipoEfectivo As String
Dim strFecha As String

Private Sub cmbAyudaCliente_Click()
    indCargandoCliente = True
    mostrarAyudaClientes txtCod_Cliente, txtGls_Cliente, txtGls_RUC, txtGls_Direccion, txtCod_Tienda
    indCargandoCliente = False
End Sub

Private Sub dtp_Fecha_Change()
Dim StrMsgError As String
Dim rst As New ADODB.Recordset

On Error GoTo Err

    csql = "Select tcFacturacion,tcCompra, tcVenta From tiposdecambio Where DATE_FORMAT(fecha,GET_FORMAT(DATE, 'EUR')) = DATE_FORMAT('" & Format(dtp_Fecha.Value, "yyyy-mm-dd") & "',GET_FORMAT(DATE, 'EUR'))"
    
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF Then
        Txt_TipoCambio.Text = Val("" & rst.Fields("tcFacturacion"))
    Else
        Txt_TipoCambio.Text = 0#
    End If

    If rst.State = 1 Then rst.Close
    Set rst = Nothing
Exit Sub
Err:
If rst.State = 1 Then rst.Close
Set rst = Nothing
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

    Me.left = 1905
    indResultado = False
    ConfGrid GDetalle, True, True, False, False
    
    ListaDocumentos StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
Exit Sub

Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
    
End Sub

Private Sub gDetalle_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    If GDetalle.Dataset.State = dsEdit Then
        GDetalle.Dataset.Post
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String

On Error GoTo Err

Select Case Button.Index
Case 1 'Nuevo
    txtCod_Cliente.Text = ""
    
    dtp_Fecha.Value = Format(getFechaSistema, "dd/mm/yyyy")
    dtp_Fecha_Change
    Set GDetalle.DataSource = Nothing
Case 2 'Procesar
    indResultado = True
    Unload Me
Case 3 'Actualizar
    ListaDocumentos StrMsgError
    If StrMsgError <> "" Then GoTo Err
Case 4 'Salir
    indResultado = False
    Unload Me
End Select

Exit Sub

Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Cliente_Change()
If indCargandoCliente Then Exit Sub

    Dim rst As New ADODB.Recordset
    If txtCod_Cliente.Text <> "" Then
        csql = "SELECT p.ruc,concat(p.direccion,' ',u.glsUbigeo,' ',d.glsUbigeo) as direccion,p.GlsPersona " & _
               "FROM personas p LEFT JOIN ubigeo u ON P.idDistrito = u.idDistrito LEFT JOIN ubigeo d ON left(u.idDistrito,2) = d.idDpto AND d.idProv = '00' AND d.idDist = '00' " & _
               "Where p.idPersona = '" & txtCod_Cliente.Text & "'"
               
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rst.EOF Then
            indCargandoCliente = True
            txtGls_RUC.Text = "" & rst.Fields("ruc")
            txtGls_Direccion.Text = "" & rst.Fields("direccion")
            txtGls_Cliente.Text = "" & rst.Fields("GlsPersona")
            indCargandoCliente = False
        End If
        rst.Close
    Else
        indCargandoCliente = True
        txtGls_RUC.Text = ""
        txtGls_Direccion.Text = ""
        txtGls_Cliente.Text = ""
        indCargandoCliente = False
    End If
    Set rst = Nothing
End Sub

Private Sub txtCod_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    mostrarAyudaClientesKeyascii KeyAscii, txtCod_Cliente, txtGls_Cliente, txtGls_RUC, txtGls_Direccion
    KeyAscii = 0
    dtp_Fecha.Value = strFecha
    If txtCod_Cliente.Text <> "" Then SendKeys "{tab}"
End If
End Sub

Private Sub txtGls_RUC_Change()
If indCargandoCliente Then Exit Sub

    Dim rst As New ADODB.Recordset
    If txtGls_RUC.Text <> "" Then
        csql = "SELECT p.idPersona,concat(p.direccion,' ',u.glsUbigeo,' ',d.glsUbigeo) as direccion,p.GlsPersona " & _
               "FROM personas p LEFT JOIN ubigeo u ON P.idDistrito = u.idDistrito LEFT JOIN ubigeo d ON left(u.idDistrito,2) = d.idDpto AND d.idProv = '00' AND d.idDist = '00' " & _
               "Where p.ruc = '" & txtGls_RUC.Text & "'"
               
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rst.EOF Then
            indCargandoCliente = True
            txtCod_Cliente.Text = "" & rst.Fields("idPersona")
            txtGls_Direccion.Text = "" & rst.Fields("direccion")
            txtGls_Cliente.Text = "" & rst.Fields("GlsPersona")
            indCargandoCliente = False
        End If
        rst.Close
    End If
    Set rst = Nothing
End Sub

Public Sub MostrarForm(ByVal strVarVendedor As String, ByRef rsRpta As ADODB.Recordset, ByRef DtpFecha As String, ByRef StrMsgError As String)

On Error GoTo Err

    strFecha = Format(DtpFecha, "dd/mm/yyyy")
    dtp_Fecha.Value = strFecha
    dtp_Fecha_Change
    Me.Show 1
    If indResultado Then
           Set rsRpta = rsg
    End If
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub ListaDocumentos(ByRef StrMsgError As String)
Dim rst As New ADODB.Recordset

Dim intElemen     As Integer
Dim dblsaldo      As Double
Dim dblTC         As Double

On Error GoTo Err

dblTC = Txt_TipoCambio.Value
    
    csql = "SELECT Fec_Comp,Nro_Comp,Fec_VCTO,idMoneda,ValTotal,ValSaldo,idCta_Dcto FROM cta_dcto " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' And ValTotal > 0 and (( idCliente = '" & txtCod_Cliente.Text & "' " & _
           "AND indDeb_Hab = 'D'  " & _
             " " & _
            ") " & _
            " OR (idCliente = '" & txtCod_Cliente.Text & "' " & _
            "AND indDeb_Hab = 'H' " & _
            "AND UCASE(LEFT(Nro_Comp,3)) = 'CRE')) " & _
           "ORDER BY Nro_Comp,Fec_Comp"
    
    
    rst.Open csql, Cn, adOpenKeyset, adLockOptimistic
    
    If rsg.State = 1 Then rsg.Close
    Set rsg = Nothing
        
        
    rsg.Fields.Append "Item", adInteger, , adFldRowID
    rsg.Fields.Append "Fecha", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "NroDocumento", adVarChar, 15, adFldIsNullable
    rsg.Fields.Append "FecVcto", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "ValTotalDoc", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "Girar", adVarChar, 1, adFldIsNullable
    rsg.Fields.Append "idCliente", adVarChar, 500, adFldIsNullable
    rsg.Fields.Append "RUC", adVarChar, 11, adFldIsNullable
    rsg.Fields.Append "idMoneda", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "TC", adVarChar, 14, adFldIsNullable
    
    
    rsg.Open
    
    If rst.RecordCount = 0 Then
        rsg.AddNew
        rsg.Fields("Item") = 1
        rsg.Fields("Fecha") = ""
        rsg.Fields("NroDocumento") = ""
        rsg.Fields("FecVcto") = ""
        rsg.Fields("ValTotalDoc") = 0
        rsg.Fields("Girar") = "N"
        rsg.Fields("idCliente") = ""
        rsg.Fields("RUC") = ""
        rsg.Fields("idMoneda") = ""
        rsg.Fields("TC") = Txt_TipoCambio.Text
    Else
        Do While Not rst.EOF
            intElemen = intElemen + 1
            dblsaldo = 0#
            
            rsg.AddNew
            
            rsg.Fields("item") = Format(intElemen, "0000")
            rsg.Fields("Fecha") = Format(rst.Fields("Fec_Comp"), "dd/mm/yyyy")
            rsg.Fields("NroDocumento") = rst.Fields("Nro_Comp") & ""
            rsg.Fields("FecVcto") = Format(rst.Fields("Fec_VCTO"), "DD/MM/YYYY")
            rsg.Fields("idCliente") = txtCod_Cliente.Text
            rsg.Fields("RUC") = txtGls_RUC.Text
            rsg.Fields("idMoneda") = IIf(Trim("" & rst.Fields("idMoneda")) = "PEN", "S/.", "US$.")
            rsg.Fields("ValTotalDoc") = Format(Val(rst.Fields("ValTotal")), "0.00")
            rsg.Fields("TC") = Txt_TipoCambio.Text
            rst.MoveNext
        Loop
    End If
    
    rst.Close
    Set rst = Nothing
    
   
    
    mostrarDatosGridSQL GDetalle, rsg, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If Not GDetalle.Dataset.EOF Then
        GDetalle.Dataset.First
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

