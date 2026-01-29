VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmListaDocExportarEntreEmpresas 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Documento"
   ClientHeight    =   8760
   ClientLeft      =   1245
   ClientTop       =   1500
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   750
      Top             =   4500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   14310
      _ExtentX        =   25241
      _ExtentY        =   1164
      ButtonWidth     =   1270
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   8160
      Left            =   0
      TabIndex        =   8
      Top             =   585
      Width           =   14220
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   13980
         Begin VB.CommandButton cmbAyudaMarca 
            Height          =   315
            Left            =   11430
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":3518
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   650
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaEmpresa 
            Height          =   315
            Left            =   5520
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":38A2
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   230
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaCliente 
            Height          =   315
            Left            =   5520
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":3C2C
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   650
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaDocumento 
            Height          =   315
            Left            =   11430
            Picture         =   "frmListaDocExportarEntreEmpresas.frx":3FB6
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   230
            Width           =   390
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   285
            Left            =   600
            TabIndex        =   10
            Top             =   1080
            Width           =   6915
            _ExtentX        =   12197
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
            Container       =   "frmListaDocExportarEntreEmpresas.frx":4340
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Documento 
            Height          =   315
            Left            =   6885
            TabIndex        =   2
            Top             =   225
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
            MaxLength       =   8
            Container       =   "frmListaDocExportarEntreEmpresas.frx":435C
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Documento 
            Height          =   315
            Left            =   7815
            TabIndex        =   14
            Top             =   225
            Width           =   3585
            _ExtentX        =   6324
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
            Container       =   "frmListaDocExportarEntreEmpresas.frx":4378
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Cliente 
            Height          =   315
            Left            =   930
            TabIndex        =   1
            Top             =   645
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
            MaxLength       =   8
            Container       =   "frmListaDocExportarEntreEmpresas.frx":4394
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Cliente 
            Height          =   315
            Left            =   1875
            TabIndex        =   17
            Top             =   645
            Width           =   3600
            _ExtentX        =   6350
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
            Container       =   "frmListaDocExportarEntreEmpresas.frx":43B0
            Vacio           =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp_Desde 
            Height          =   315
            Left            =   12645
            TabIndex        =   4
            Top             =   225
            Width           =   1200
            _ExtentX        =   2117
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
            Format          =   193593345
            CurrentDate     =   38955
         End
         Begin MSComCtl2.DTPicker dtp_Hasta 
            Height          =   315
            Left            =   12645
            TabIndex        =   5
            Top             =   645
            Width           =   1200
            _ExtentX        =   2117
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
            Format          =   193593345
            CurrentDate     =   38955
         End
         Begin CATControls.CATTextBox txtCod_Empresa 
            Height          =   315
            Left            =   930
            TabIndex        =   0
            Top             =   240
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
            MaxLength       =   8
            Container       =   "frmListaDocExportarEntreEmpresas.frx":43CC
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Empresa 
            Height          =   315
            Left            =   1875
            TabIndex        =   22
            Top             =   240
            Width           =   3600
            _ExtentX        =   6350
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
            Container       =   "frmListaDocExportarEntreEmpresas.frx":43E8
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Marca 
            Height          =   315
            Left            =   6885
            TabIndex        =   3
            Top             =   645
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
            MaxLength       =   8
            Container       =   "frmListaDocExportarEntreEmpresas.frx":4404
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Marca 
            Height          =   315
            Left            =   7815
            TabIndex        =   25
            Top             =   645
            Width           =   3585
            _ExtentX        =   6324
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
            Container       =   "frmListaDocExportarEntreEmpresas.frx":4420
            Vacio           =   -1  'True
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   6030
            TabIndex        =   26
            Top             =   675
            Width           =   450
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   225
            TabIndex        =   23
            Top             =   270
            Width           =   630
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   12090
            TabIndex        =   20
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   12090
            TabIndex        =   19
            Top             =   675
            Width           =   420
         End
         Begin VB.Label lbl_Cliente 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   225
            TabIndex        =   18
            Top             =   675
            Width           =   480
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   6000
            TabIndex        =   15
            Top             =   255
            Width           =   810
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Busqueda"
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
            Left            =   3000
            TabIndex        =   11
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3825
         Left            =   75
         OleObjectBlob   =   "frmListaDocExportarEntreEmpresas.frx":443C
         TabIndex        =   7
         Top             =   1485
         Width           =   14025
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   2535
         Left            =   75
         OleObjectBlob   =   "frmListaDocExportarEntreEmpresas.frx":DFA2
         TabIndex        =   6
         Top             =   5475
         Width           =   14025
      End
   End
End
Attribute VB_Name = "frmListaDocExportarEntreEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsg As New ADODB.Recordset
Private rsd As New ADODB.Recordset
Private rse As New ADODB.Recordset
Private strTDExportar As String
Dim indNuevoDoc As Boolean

Private Sub cmbAyudaCliente_Click()
Dim StrMsgError As String

    'Insertamos los Clientes y sus tinedas que no se encuentran  en la Empresa Logistica Apimas
    csql = "Insert Into Clientes(idCliente, idEmpresa, idVendedorCampo, idEmpTrans, idLista, idFormaPago, Val_LineaCredito, indAgenteRetencion, Val_LineaCreditoConsumida, idMonedaLineaCredito, Val_Dscto, idClienteInterno, idCanal, indventasterceros, idGrupoCliente, idProvCliente, GlsObservacion, indComision) " & _
           "Select idCliente, '" & glsEmpresa & "', idVendedorCampo, idEmpTrans, idLista, idFormaPago, Val_LineaCredito, indAgenteRetencion, Val_LineaCreditoConsumida, idMonedaLineaCredito, Val_Dscto, idClienteInterno, idCanal, indventasterceros, idGrupoCliente, idProvCliente, GlsObservacion, indComision " & _
           "From Clientes c Inner Join Personas p  on c.idCliente = p.idPersona " & _
           "Where idEmpresa  = '01' " & _
           "And IdCliente Not In(Select IdCliente From Clientes Where idEmpresa  = '" & glsEmpresa & "')"
    Cn.Execute (csql)
     
    csql = "Insert Into  TiendasCliente(idEmpresa, idPersona, item, GlsNombre, GlsDireccion, GlsTelefonos, GlsContacto, idDistrito, idPais, idtdacli) " & _
            "Select '" & glsEmpresa & "', p.idPersona, item, GlsNombre, GlsDireccion, GlsTelefonos, p.GlsContacto, p.idDistrito, p.idPais, idtdacli " & _
            "From TiendasCliente t Inner Join Personas p  on t.idPersona = p.idPersona " & _
            "Where  t.idEmpresa  = '01' " & _
            "And  iDtdaCli Not In (Select iDtdaCli From  TiendasCliente Where  idEmpresa  = '" & glsEmpresa & "')"
    Cn.Execute (csql)
            
    csql = "INSERT INTO EmpTrans(idEmpTrans, idEmpresa) " & _
            "SELECT idEmpTrans, '" & glsEmpresa & "' FROM  EmpTrans " & _
            "WHERE IdEmpresa  = '01' " & _
            "AND idEmpTrans NOT IN (SELECT idEmpTrans FROM EmpTrans WHERE  idEmpresa = '" & glsEmpresa & "') "
    Cn.Execute (csql)
    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub cmbAyudaDocumento_Click()
    
    mostrarAyuda "DOCUMENOSEXP", txtCod_Documento, txtGls_Documento, " AND c.idDocumento = '" & strTDExportar & "'"

End Sub

Private Sub cmbAyudaEmpresa_Click()
    
    mostrarAyuda "EMPRESA", txtCod_Empresa, txtGls_Empresa, " And idEmpresa Not In('" & glsEmpresa & "')"

End Sub

Private Sub CmbAyudaMarca_Click()
    
    mostrarAyuda "MARCA", txtCod_Marca, txtGls_Marca

End Sub

Private Sub dtp_Desde_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub dtp_Hasta_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    strRptNum = ""
    strRptSerie = ""
    ConfGrid gLista, True, False, False, False
    ConfGrid gListaDetalle, True, False, False, False
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub listaDocVentas(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst               As New ADODB.Recordset
Dim rsr               As New ADODB.Recordset
Dim strCond           As String
Dim strFiltroAprob    As String

    '--- FORMATO GRILLA
    Set gLista.DataSource = Nothing
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing

    If rsd.State = 1 Then rsd.Close: Set rsd = Nothing
    
    '--- Formato cabecera
    rsg.Fields.Append "Item", adChar, 20, adFldRowID
    rsg.Fields.Append "chkMarca", adChar, 1, adFldIsNullable
    rsg.Fields.Append "idDocVentas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    rsg.Fields.Append "idPerVendedor", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsVendedor", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "FecEmision", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "estDocVentas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsg.Fields.Append "TotalValorVenta", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVVenta", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPrecioVenta", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "numOrdenCompra", adVarChar, 200, adFldIsNullable
    rsg.Fields.Append "llegada", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "FecIniTraslado", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "idFormaPago", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "TC", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "idPerChofer", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "glsChofer", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "idPerEmpTrans", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GlsEmpTrans", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "idVehiculo", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GlsVehiculo", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "Placa", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "Marca", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "Color", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "Modelo", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "CodInsCrip", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "Brevete", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "rucEmpTrans", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "IdCentroCosto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdEmpresa", adVarChar, 2, adFldIsNullable
    rsg.Fields.Append "IdSucursal", adVarChar, 8, adFldIsNullable
    rsg.Open , , adOpenKeyset, adLockOptimistic
    
    '--- Formato Detalle
    rsd.Fields.Append "Item", adVarChar, 20, adFldRowID
    rsd.Fields.Append "chkMarca", adChar, 1, adFldIsNullable
    rsd.Fields.Append "idProducto", adChar, 8, adFldIsNullable
    rsd.Fields.Append "idCodFabricante", adVarChar, 30, adFldIsNullable
    rsd.Fields.Append "GlsProducto", adVarChar, 500, adFldIsNullable
    rsd.Fields.Append "idMarca", adChar, 8, adFldIsNullable
    rsd.Fields.Append "GlsMarca", adVarChar, 185, adFldIsNullable
    rsd.Fields.Append "idUM", adChar, 8, adFldIsNullable
    rsd.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    rsd.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    rsd.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "Cantidad2", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "TotalVVBruto", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "TotalPVBruto", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "PorDcto", adVarChar, 20, adFldIsNullable
    rsd.Fields.Append "DctoVV", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "DctoPV", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "idTipoProducto", adChar, 5, adFldIsNullable
    rsd.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsd.Fields.Append "idDocVentas", adChar, 8, adFldIsNullable
    rsd.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    rsd.Fields.Append "NumLote", adVarChar, 30, adFldIsNullable
    rsd.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    rsd.Fields.Append "VVUnitLista", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "PVUnitLista", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "VVUnitNeto", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "PVUnitNeto", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "CodigoRapido", adVarChar, 30, adFldIsNullable
    rsd.Fields.Append "idTallaPeso", adDouble, 14, adFldIsNullable
    rsd.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    rsd.Fields.Append "IdEmpresa", adVarChar, 2, adFldIsNullable
    rsd.Fields.Append "IdSucursal", adVarChar, 8, adFldIsNullabl
    rsd.Fields.Append "ItemPro", adInteger, , adFldRowID
    rsd.Open , , adOpenKeyset, adLockOptimistic
    
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND GlsVendedor LIKE '%" & strCond & "%'"
    End If
    csql = ""
    csql = "Select ConCat(d.IdDocumento,d.IdDocVentas,d.IdSerie) As Item,d.IdFormaPago,d.IdDocVentas,d.IdSerie,d.IdDocumento,d.IdPerVendedor,d.GlsVendedor,d.FecEmision," & _
               "d.EstDocVentas,d.IdMoneda,d.TotalValorVenta,d.TotalIGVVenta,d.TotalPrecioVenta,d.NumOrdenCompra,d.Llegada,d.FecIniTraslado,d.TipoCambio," & _
               "d.IdPerChofer,d.GlsChofer,d.IdPerEmpTrans,d.GlsEmpTrans,d.IdVehiculo,d.GlsVehiculo,d.Placa,d.Marca,d.Color,d.Modelo,d.CodInsCrip,d.Brevete,d.RucEmpTrans," & _
               "d.IdCentroCosto,d.idEmpresa,d.idSucursal,dd.idMarca  " & _
               "From Docventas d Inner Join DocventasDet dd   On d.idEmpresa = dd.idEmpresa   And d.idSucursal = dd.idSucursal " & _
               "And d.idDocumento = dd.idDocumento   And d.idSerie = dd.idSerie   And d.idDocventas = dd.idDocVentas " & _
               "Where d.IdEmpresa = '" & Trim(txtCod_Empresa.Text) & "' And d.IdDocumento = '" & txtCod_Documento.Text & "' " & _
               "And d.EstDocventas <> 'ANU' And d.IdPerCliente = '" & txtCod_Cliente.Text & "' " & _
               "And d.FecEmision BetWeen '" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "' And '" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "' " & _
               "And(Select Count(*) From docreferenciaempresasdet dre " & _
               "Where dre.idEmpresaReferencia = dd.idEmpresa " & _
               "And dre.numDocReferencia = dd.idDocventas " & _
               "And dre.idSucursalReferencia  =dd.idSucursal " & _
               "And dre.tipoDocReferencia = dd.idDocumento " & _
               "And dre.serieDocReferencia = dd.idSerie " & _
               "And dre.idProducto = dd.idProducto ) = 0 " & _
               "And idMarca like '%" & txtCod_Marca.Text & "%' " & _
               "Group By d.idEmpresa,d.idSucursal,d.idSerie ,d.idDocventas "

    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF And Not rst.BOF Then
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = rst.Fields("Item")
            rsg.Fields("chkMarca") = 0
            rsg.Fields("idDocVentas") = rst.Fields("idDocVentas")
            rsg.Fields("idSerie") = rst.Fields("idSerie")
            rsg.Fields("idPerVendedor") = rst.Fields("idPerVendedor")
            rsg.Fields("GlsVendedor") = rst.Fields("GlsVendedor")
            rsg.Fields("FecEmision") = Format(rst.Fields("FecEmision"), "dd/mm/yyyy")
            rsg.Fields("estDocVentas") = rst.Fields("estDocVentas")
            rsg.Fields("idMoneda") = rst.Fields("idMoneda")
            rsg.Fields("TotalValorVenta") = rst.Fields("TotalValorVenta")
            rsg.Fields("TotalIGVVenta") = rst.Fields("TotalIGVVenta")
            rsg.Fields("TotalPrecioVenta") = rst.Fields("TotalPrecioVenta")
            rsg.Fields("numOrdenCompra") = rst.Fields("numOrdenCompra")
            rsg.Fields("llegada") = rst.Fields("llegada")
            rsg.Fields("idmoneda") = rst.Fields("idmoneda")
            rsg.Fields("FecIniTraslado") = Format(rst.Fields("FecIniTraslado"), "dd/mm/yyyy")
            rsg.Fields("idFormaPago") = Trim("" & rst.Fields("idFormaPago"))
            rsg.Fields("TC") = Trim("" & rst.Fields("TipoCambio"))
            rsg.Fields("idPerChofer") = Trim("" & rst.Fields("idPerChofer"))
            rsg.Fields("glsChofer") = Trim("" & rst.Fields("glsChofer"))
            rsg.Fields("idPerEmpTrans") = Trim("" & rst.Fields("idPerEmpTrans"))
            rsg.Fields("GlsEmpTrans") = Trim("" & rst.Fields("GlsEmpTrans"))
            rsg.Fields("idVehiculo") = Trim("" & rst.Fields("idVehiculo"))
            rsg.Fields("GlsVehiculo") = Trim("" & rst.Fields("GlsVehiculo"))
            rsg.Fields("Placa") = Trim("" & rst.Fields("Placa"))
            rsg.Fields("Marca") = Trim("" & rst.Fields("Marca"))
            rsg.Fields("Color") = Trim("" & rst.Fields("Color"))
            rsg.Fields("Modelo") = Trim("" & rst.Fields("Modelo"))
            rsg.Fields("CodInsCrip") = Trim("" & rst.Fields("CodInsCrip"))
            rsg.Fields("Brevete") = Trim("" & rst.Fields("Brevete"))
            rsg.Fields("rucEmpTrans") = Trim("" & rst.Fields("rucEmpTrans"))
            rsg.Fields("IdCentroCosto") = Trim("" & rst.Fields("IdCentroCosto"))
            rsg.Fields("IdEmpresa") = Trim("" & rst.Fields("IdEmpresa"))
            rsg.Fields("IdSucursal") = Trim("" & rst.Fields("IdSucursal"))
            rst.MoveNext
        Loop
        mostrarDatosGridSQL gLista, rsg, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    Me.Refresh
    If rst.State = 1 Then rst.Close: Set rst = Nothing

    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub listaDetalle(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst                                     As New ADODB.Recordset
Dim strCond                                 As String
Dim indExisteDoc                            As Boolean
Dim strNumDoc                               As String
Dim strSerie                                As String
Dim NValoresPedido()                        As Double

    '--- FORMATO GRILLA
    strNumDoc = gLista.Columns.ColumnByFieldName("idDocVentas").Value
    strSerie = gLista.Columns.ColumnByFieldName("idSerie").Value
    
    '--- Validamos si ya adicionamos el detalle
    gListaDetalle.Dataset.Filter = ""
    gListaDetalle.Dataset.Filtered = True
    indExisteDoc = False
    Set gListaDetalle.DataSource = Nothing
    gListaDetalle.Dataset.DisableControls
    If rsd.RecordCount > 0 Then rsd.MoveFirst
    
    Do While Not rsd.EOF
        If rsd.Fields("idDocVentas") = strNumDoc And rsd.Fields("idSerie") = strSerie Then
            indExisteDoc = True
            Exit Do
        End If
        rsd.MoveNext
    Loop
    
    If indExisteDoc = False Then
        strCond = ""
        If Trim(txt_TextoBuscar.Text) <> "" Then
            strCond = Trim(txt_TextoBuscar.Text)
            strCond = " AND GlsCliente LIKE '%" & strCond & "%'"
        End If
        
        csql = "Select Item,IdProducto,IdCodFabricante,GlsProducto,IdMarca,GlsMarca,IdUM,GlsUM,Factor,Afecto," & _
           "Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVBruto,TotalPVBruto,PorDcto,DctoVV,DctoPV,TotalVVNeto," & _
           "TotalIGVNeto,TotalPVNeto,IdTipoProducto,IdMoneda,NumLote,FecVencProd,VVUnitLista,PVUnitLista,VVUnitNeto,PVUnitNeto,Cantidad2," & _
           "CodigoRapido,IdTallaPeso,IdDocVentas,IdSerie,IdDocumento,idEmpresa,idSucursal,ItemPro " & _
           "From DocVentasDet dd " & _
           "Where  IdEmpresa = '" & Trim(txtCod_Empresa.Text) & "'   And IdDocumento = '" & txtCod_Documento.Text & "' " & _
           "And IdDocVentas = '" & strNumDoc & "' And IdSerie = '" & strSerie & "'  " & _
           "And(Select Count(*) From docreferenciaempresasdet dre " & _
           "Where dre.idEmpresaReferencia = dd.idEmpresa " & _
           "And dre.numDocReferencia = dd.idDocventas " & _
           "And dre.idSucursalReferencia  =dd.idSucursal " & _
           "And dre.tipoDocReferencia = dd.idDocumento " & _
           "And dre.serieDocReferencia = dd.idSerie " & _
           "And dre.idProducto = dd.idProducto ) = 0 " & _
           "And idMarca like '%" & txtCod_Marca.Text & "%' "
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
     
        Do While Not rst.EOF
            rsd.AddNew
            rsd.Fields("Item") = strNumDoc & strSerie & CStr(rst.Fields("Item"))
            rsd.Fields("chkMarca") = 0
            rsd.Fields("idProducto") = rst.Fields("idProducto")
            rsd.Fields("idCodFabricante") = "" & rst.Fields("idCodFabricante")
            rsd.Fields("GlsProducto") = "" & rst.Fields("GlsProducto")
            rsd.Fields("idMarca") = "" & rst.Fields("idMarca")
            rsd.Fields("GlsMarca") = "" & rst.Fields("GlsMarca")
            rsd.Fields("idUM") = "" & rst.Fields("idUM")
            rsd.Fields("GlsUM") = "" & rst.Fields("GlsUM")
            rsd.Fields("Factor") = "" & rst.Fields("Factor")
            rsd.Fields("Afecto") = "" & rst.Fields("Afecto")
            rsd.Fields("Cantidad") = "" & rst.Fields("Cantidad")
            rsd.Fields("Cantidad2") = "" & rst.Fields("Cantidad2")
            rsd.Fields("VVUnit") = "" & rst.Fields("VVUnit")
            rsd.Fields("IGVUnit") = "" & rst.Fields("IGVUnit")
            rsd.Fields("PVUnit") = "" & rst.Fields("PVUnit")
            rsd.Fields("TotalVVBruto") = "" & rst.Fields("TotalVVBruto")
            rsd.Fields("TotalPVBruto") = "" & rst.Fields("TotalPVBruto")
            rsd.Fields("PorDcto") = "" & rst.Fields("PorDcto")
            rsd.Fields("DctoVV") = "" & rst.Fields("DctoVV")
            rsd.Fields("DctoPV") = "" & rst.Fields("DctoPV")
            rsd.Fields("TotalVVNeto") = "" & rst.Fields("TotalVVNeto")
            rsd.Fields("TotalIGVNeto") = "" & rst.Fields("TotalIGVNeto")
            rsd.Fields("TotalPVNeto") = "" & rst.Fields("TotalPVNeto")
            rsd.Fields("idTipoProducto") = "" & rst.Fields("idTipoProducto")
            rsd.Fields("idMoneda") = "" & rst.Fields("idMoneda")
            rsd.Fields("idDocVentas") = strNumDoc
            rsd.Fields("idSerie") = strSerie
            rsd.Fields("NumLote") = "" & rst.Fields("NumLote")
            rsd.Fields("FecVencProd") = "" & rst.Fields("FecVencProd")
            rsd.Fields("VVUnitLista") = "" & rst.Fields("VVUnitLista")
            rsd.Fields("PVUnitLista") = "" & rst.Fields("PVUnitLista")
            rsd.Fields("VVUnitNeto") = "" & rst.Fields("VVUnitNeto")
            rsd.Fields("PVUnitNeto") = "" & rst.Fields("PVUnitNeto")
            rsd.Fields("CodigoRapido") = "" & rst.Fields("CodigoRapido")
            rsd.Fields("idTallaPeso") = Val("" & rst.Fields("idTallaPeso"))
            rsd.Fields("idDocumento") = "" & rst.Fields("idDocumento")
            rsd.Fields("IdEmpresa") = "" & rst.Fields("IdEmpresa")
            rsd.Fields("IdSucursal") = "" & rst.Fields("IdSucursal")
            rsd.Fields("ItemPro") = Val("" & rst.Fields("ItemPro"))
            rst.MoveNext
        Loop
    End If
    
    If rsd.RecordCount > 0 Then
        mostrarDatosGridSQL gListaDetalle, rsd, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    gListaDetalle.Dataset.Filter = " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    gListaDetalle.Dataset.Filtered = True
    gListaDetalle.Dataset.EnableControls
    Me.Refresh
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub

Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    gListaDetalle.Dataset.EnableControls
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    If rsd.State = 1 Then rsd.Close: Set rsd = Nothing

End Sub

Private Sub gLista_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    
    If gLista.Dataset.State = dsEdit Then
        gLista.Dataset.Post
    End If
    If gListaDetalle.Count = 0 Then Exit Sub
    gListaDetalle.Dataset.First
    
    Do While Not gListaDetalle.Dataset.EOF
        gListaDetalle.Dataset.Edit
        gListaDetalle.Columns.ColumnByFieldName("chkMarca").Value = Column.Value
        gListaDetalle.Dataset.Post
        gListaDetalle.Dataset.Next
    Loop

End Sub

Private Sub gListaDetalle_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    
    If gListaDetalle.Dataset.State = dsEdit Then
        gListaDetalle.Dataset.Post
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError     As String
Dim strNumDoc       As String
Dim strSerie        As String
Dim PAR             As String

    Select Case Button.Index
        Case 1 'Nuevo
            If gLista.Count > 0 Then
                strNumDoc = gLista.Columns.ColumnByFieldName("idDocVentas").Value
                strSerie = gLista.Columns.ColumnByFieldName("idSerie").Value
                ValCot = traerCampo("docventas", "FecEmision", "iddocventas", strNumDoc, False, " idserie='" & strSerie & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'")
                PAR = traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VALIDEZ_COT_ALSISAC", False, "idEmpresa = '" & glsEmpresa & "'")
                If Trim(txtCod_Documento.Text) = "92" Then
                    If Val(DateDiff("d", CVDate(ValCot), CVDate(getFechaSistema))) >= PAR Then
                        MsgBox ("No se puede importar la Cotización, porque excede su fecha de vigencia.")
                        Exit Sub
                    Else
                        Me.Hide
                    End If
                Else
                    Me.Hide
                End If
            End If
        
        Case 3
            strRptNum = ""
            strRptSerie = ""
            Me.Hide
            Unload Me
    End Select
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String

    listaDetalle StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    If gLista.Count > 0 Then
        Me.Hide
    End If
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Public Sub MostrarForm(ByVal strTipoDocQueImporta As String, ByVal strCodCliente As String, ByRef rscd As ADODB.Recordset, ByRef rsdd As ADODB.Recordset, ByRef strTipoDocImportado As String, ByRef StrMsgError As String)
On Error GoTo Err

    indNuevoDoc = True
    Set gLista.DataSource = Nothing
    Set gListaDetalle.DataSource = Nothing
    strTDExportar = strTipoDocQueImporta
    txtCod_Cliente.Text = strCodCliente
    txtCod_Documento.Text = ""
    dtp_Desde.Value = getFechaSistema
    dtp_Hasta.Value = dtp_Desde.Value
    indNuevoDoc = False
   
    frmListaDocExportarEntreEmpresas.Show 1
    
    '--- Quitamos Filtros existentes
    gLista.Dataset.Filter = ""
    gLista.Dataset.Filtered = True
    
    gListaDetalle.Dataset.Filter = ""
    gListaDetalle.Dataset.Filtered = True
    
    Set gLista.DataSource = Nothing
    Set gListaDetalle.DataSource = Nothing
    
    If TypeName(rsg) = "Nothing" Then
        Exit Sub
    Else
        If rsg.State = 0 Then
            Exit Sub
        End If
    End If
    
    '--- Eliminamos los registros q no estan marcados
    If rsg.RecordCount > 0 Then
        rsg.MoveFirst
        Do While Not rsg.EOF
            If rsg.Fields("chkMarca") = "0" Then
                rsg.Delete adAffectCurrent
                rsg.Update
            End If
            rsg.MoveNext
        Loop
    End If
    
    If rsd.RecordCount > 0 Then
        rsd.MoveFirst
        Do While Not rsd.EOF
            If rsd.Fields("chkMarca") = "0" Then
                rsd.Delete adAffectCurrent
                rsd.Update
            End If
            rsd.MoveNext
        Loop
    End If
        
    '--- Devolvemos valores seleccionados
    strTipoDocImportado = txtCod_Documento.Text
       
    Set rscd = rsg.Clone(adLockReadOnly)
    Set rsdd = rsd.Clone(adLockReadOnly)
    
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    If rsd.State = 1 Then rsd.Close: Set rsd = Nothing

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    
End Sub

Private Sub txtCod_Cliente_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txtCod_Cliente_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CLIENTE", txtCod_Cliente, txtGls_Cliente
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Documento_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_Documento.Text = traerCampo("documentos", "GlsDocumento", "idDocumento", txtCod_Documento.Text, False)
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txtCod_Documento_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DOCUMENOSEXP", txtCod_Documento, txtGls_Documento, " AND c.idDocumento = '" & strTDExportar & "'"
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Empresa_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txtCod_Marca_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_Marca.Text = traerCampo("marcas", "glsMarca", "idMarca", txtCod_Marca.Text, True)
    
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub
