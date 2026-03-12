VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmCambiarComisionesDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Comision al Documento"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraComision 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   3450
      TabIndex        =   21
      Top             =   3300
      Visible         =   0   'False
      Width           =   5715
      Begin VB.CommandButton cmdOK 
         Caption         =   "Aceptar"
         Height          =   345
         Left            =   825
         TabIndex        =   23
         Top             =   1500
         Width           =   1740
      End
      Begin VB.CommandButton cmbCancelar 
         Caption         =   "Cancelar"
         Height          =   345
         Left            =   3300
         TabIndex        =   27
         Top             =   1500
         Width           =   1740
      End
      Begin CATControls.CATTextBox txtVal_NuevaComision 
         Height          =   285
         Left            =   2850
         TabIndex        =   22
         Tag             =   "NComisionVtas"
         Top             =   825
         Width           =   1815
         _ExtentX        =   3201
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
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   11
         Container       =   "frmCambiarComisionesDoc.frx":0000
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_ComisionActual 
         Height          =   285
         Left            =   2850
         TabIndex        =   25
         Tag             =   "NComisionVtas"
         Top             =   450
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BackColor       =   12640511
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
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   11
         Container       =   "frmCambiarComisionesDoc.frx":001C
         Text            =   "0.00"
         Decimales       =   2
         Estilo          =   4
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Comision Actual:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   1350
         TabIndex        =   26
         Top             =   450
         Width           =   1440
      End
      Begin VB.Label lbl_Comision 
         Appearance      =   0  'Flat
         Caption         =   "Nueva Comision:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   1350
         TabIndex        =   24
         Top             =   825
         Width           =   1440
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   675
      Top             =   1500
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
            Picture         =   "frmCambiarComisionesDoc.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":03D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":0824
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":0BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":0F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":12F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":168C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":1A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":1DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":215A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":24F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":31B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":3550
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":39A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":3D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarComisionesDoc.frx":474E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   1164
      ButtonWidth     =   1455
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refrescar"
            Object.ToolTipText     =   "Pagos"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      Caption         =   "Listado"
      ForeColor       =   &H00C00000&
      Height          =   8865
      Left            =   75
      TabIndex        =   1
      Top             =   675
      Width           =   12945
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   150
         TabIndex        =   12
         Top             =   225
         Width           =   12690
         Begin VB.CommandButton cmbAyudaDocumento 
            Height          =   315
            Left            =   12150
            Picture         =   "frmCambiarComisionesDoc.frx":4E20
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   225
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaVendedor 
            Height          =   315
            Left            =   5550
            Picture         =   "frmCambiarComisionesDoc.frx":51AA
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   225
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Vendedor 
            Height          =   285
            Left            =   990
            TabIndex        =   14
            Tag             =   "TidPerVendedor"
            Top             =   225
            Width           =   915
            _ExtentX        =   1614
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
            Container       =   "frmCambiarComisionesDoc.frx":5534
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Vendedor 
            Height          =   285
            Left            =   1950
            TabIndex        =   15
            Tag             =   "TGlsVendedor"
            Top             =   225
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   503
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
            Container       =   "frmCambiarComisionesDoc.frx":5550
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_TipoDocumento 
            Height          =   285
            Left            =   7590
            TabIndex        =   18
            Tag             =   "TidPerVendedor"
            Top             =   225
            Width           =   915
            _ExtentX        =   1614
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
            Container       =   "frmCambiarComisionesDoc.frx":556C
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_TipoDocumento 
            Height          =   285
            Left            =   8550
            TabIndex        =   19
            Tag             =   "TGlsVendedor"
            Top             =   225
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   503
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
            Container       =   "frmCambiarComisionesDoc.frx":5588
            Vacio           =   -1  'True
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            Caption         =   "Tipo Documento:"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   6300
            TabIndex        =   20
            Top             =   300
            Width           =   1365
         End
         Begin VB.Label lbl_Vendedor 
            Appearance      =   0  'Flat
            Caption         =   "Vendedor:"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   150
            TabIndex        =   16
            Top             =   300
            Width           =   765
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   12690
         Begin VB.ComboBox cbx_Mes 
            Height          =   315
            ItemData        =   "frmCambiarComisionesDoc.frx":55A4
            Left            =   8400
            List            =   "frmCambiarComisionesDoc.frx":55CC
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   225
            Width           =   1665
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   285
            Left            =   1500
            TabIndex        =   6
            Top             =   210
            Width           =   5340
            _ExtentX        =   9419
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
            Container       =   "frmCambiarComisionesDoc.frx":5635
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Ano 
            Height          =   285
            Left            =   11175
            TabIndex        =   7
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmCambiarComisionesDoc.frx":5651
            Estilo          =   3
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Busqueda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Width           =   765
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Mes:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7800
            TabIndex        =   10
            Top             =   300
            Width           =   345
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Año:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   10725
            TabIndex        =   9
            Top             =   300
            Width           =   330
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "(Filtra por Razon social del cliente o Numero del documento )"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   1950
            TabIndex        =   8
            Top             =   480
            Width           =   4275
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   4185
         Left            =   120
         OleObjectBlob   =   "frmCambiarComisionesDoc.frx":566D
         TabIndex        =   2
         Top             =   1650
         Width           =   12735
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   2685
         Left            =   75
         OleObjectBlob   =   "frmCambiarComisionesDoc.frx":9307
         TabIndex        =   3
         Top             =   6000
         Width           =   12735
      End
   End
End
Attribute VB_Name = "frmCambiarComisionesDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indCargando As Boolean

Private Sub cbx_Mes_Click()
Dim StrMsgError As String
On Error GoTo Err
listaDocVentas StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaDocumento_Click()
mostrarAyuda "DOCUMENTOS", txtCod_TipoDocumento, txtGls_TipoDocumento
End Sub

Private Sub cmbAyudaVendedor_Click()
    mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor
End Sub

Private Sub cmbCancelar_Click()
FraListado.Enabled = True
Toolbar1.Enabled = True
FraComision.Visible = False
End Sub

Private Sub cmdOK_Click()
Dim StrMsgError As String
On Error GoTo Err

'Actualizar Comision
If MsgBox("¿Seguro de Cambiar la Comision?", vbQuestion + vbYesNo, App.Title) = vbYes Then

    csql = "UPDATE docventas SET ComisionVtas = " & txtVal_NuevaComision.Value & _
           "WHERE idEmpresa = '" & glsEmpresa & "' " & _
             "AND idSucursal = '" & glsSucursal & "' " & _
             "AND idDocumento = '" & GLista.Columns.ColumnByFieldName("idDocumento").Value & "' " & _
             "AND idDocVentas = '" & GLista.Columns.ColumnByFieldName("idDocVentas").Value & "' " & _
             "AND idSerie = '" & GLista.Columns.ColumnByFieldName("idSerie").Value & "'"
    
    Cn.Execute csql
        
    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    FraListado.Enabled = True
    Toolbar1.Enabled = True
    FraComision.Visible = False

End If
    
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()

ConfGrid GLista, False, False, False, False
ConfGrid gListaDetalle, False, False, False, False

indCargando = True

cbx_Mes.ListIndex = Month(getFechaSistema) - 1
txt_Ano.Text = Format(Year(getFechaSistema), "0000")

indCargando = False

End Sub

Private Sub gLista_GotFocus()
listaDetalle
End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    listaDetalle
End Sub

Private Sub gLista_OnDblClick()
Dim strTD As String
Dim strNumDoc As String
Dim strSerie As String

    FraListado.Enabled = False
    Toolbar1.Enabled = False
    
    strTD = "" & GLista.Columns.ColumnByFieldName("idDocumento").Value
    strNumDoc = "" & GLista.Columns.ColumnByFieldName("idDocVentas").Value
    strSerie = "" & GLista.Columns.ColumnByFieldName("idSerie").Value
    
    txtVal_ComisionActual.Text = traerCampo("docventas", "ComisionVtas", "idDocumento", strTD, True, " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "' AND idSucursal = '" & glsSucursal & "'")
    txtVal_NuevaComision.Text = 0
    
    FraComision.Visible = True
    
    txtVal_NuevaComision.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim StrMsgError As String
On Error GoTo Err
Select Case Button.Index
    Case 1 'Refrescar
    
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err

    Case 2 'Salir
        Unload Me
End Select
Exit Sub
Err:
MsgBox StrMsgError, vbInformation, App.Title
End Sub


Private Sub txt_Ano_Change()
Dim StrMsgError As String
On Error GoTo Err
listaDocVentas StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_Change()
Dim StrMsgError As String
On Error GoTo Err
listaDocVentas StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_TipoDocumento_Change()
Dim StrMsgError As String
On Error GoTo Err

txtGls_TipoDocumento.Text = traerCampo("documentos", "GlsDocumento", "idDocumento", txtCod_TipoDocumento.Text, False)

listaDocVentas StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_TipoDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    mostrarAyudaKeyascii KeyAscii, "DOCUMENTO", txtCod_TipoDocumento, txtGls_TipoDocumento
    KeyAscii = 0
End If
End Sub

Private Sub txtCod_Vendedor_Change()
Dim StrMsgError As String
On Error GoTo Err

txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)

listaDocVentas StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Vendedor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    mostrarAyudaKeyascii KeyAscii, "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor
    KeyAscii = 0
End If
End Sub


Private Sub listaDocVentas(ByRef StrMsgError As String)
Dim strCond As String
Dim strTipoDoc As String
Dim strCodVendedor As String

On Error GoTo Err

    If indCargando Then Exit Sub
    
    strTipoDoc = txtCod_TipoDocumento.Text
    strCodVendedor = txtCod_Vendedor.Text

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND (GlsCliente LIKE '%" & strCond & "%' or idDocVentas LIKE '%" & strCond & "%') "
    End If
    
    csql = "SELECT concat(idDocumento,idDocVentas,idSerie) as Item ,idDocumento, idDocVentas,idSerie,idPerCliente,GlsCliente,RUCCliente,DATE_FORMAT(FecEmision,GET_FORMAT(DATE, 'EUR')) as FecEmision, ComisionVtas, estDocVentas,Format(TotalPrecioVenta,2) AS TotalPrecioVenta " & _
            "FROM docventas " & _
            "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'  AND year(FecEmision) = " & Val(txt_Ano.Text) & " AND Month(FecEmision) = " & cbx_Mes.ListIndex + 1
           
    If strCond <> "" Then csql = csql + strCond
    
    If strTipoDoc <> "" Then
        csql = csql + " AND (idDocumento = '" & strTipoDoc & "') "
    End If
    
    If strCodVendedor <> "" Then
        csql = csql + " AND (idPerVendedor = '" & strCodVendedor & "' OR idPerVendedorCampo = '" & strCodVendedor & "' ) "
    End If

    csql = csql + " ORDER BY idSerie,idDocVentas,FecEmision"
    
    With GLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With
    
    'DETALLE
    
    listaDetalle
    
Me.Refresh
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub


Private Sub listaDetalle()
    csql = "SELECT item, idProducto, GlsProducto, GlsMarca, GlsUM, Format(Cantidad,2) AS Cantidad, Format(PVUnit,2) AS PVUnit, FORMAT(PorDcto,2) AS PorDcto, Format(TotalPVNeto,2) AS TotalPVNeto " & _
           "FROM docventasdet " & _
           "WHERE idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "' AND idDocumento = '" & GLista.Columns.ColumnByFieldName("idDocumento").Value & "' AND idDocVentas = '" & GLista.Columns.ColumnByFieldName("idDocVentas").Value & "' AND idSerie = '" & GLista.Columns.ColumnByFieldName("idSerie").Value & "'"
    
    With gListaDetalle
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "item"
    End With
End Sub

