VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantFormasPago 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Formas de Pago"
   ClientHeight    =   5460
   ClientLeft      =   2910
   ClientTop       =   3195
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8910
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   4560
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
            Picture         =   "frmMantFormasPago.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFormasPago.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   4785
      Left            =   90
      TabIndex        =   14
      Top             =   630
      Width           =   8745
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   120
         TabIndex        =   15
         Top             =   150
         Width           =   8490
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   960
            TabIndex        =   0
            Top             =   225
            Width           =   7365
            _ExtentX        =   12991
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
            Container       =   "frmMantFormasPago.frx":3518
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   16
            Top             =   265
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3660
         Left            =   120
         OleObjectBlob   =   "frmMantFormasPago.frx":3534
         TabIndex        =   1
         Top             =   960
         Width           =   8490
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   90
      TabIndex        =   9
      Top             =   630
      Width           =   8730
      Begin VB.CheckBox Chk_cheque 
         Caption         =   "Cheque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   5
         Top             =   2070
         Width           =   915
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   310
         TabIndex        =   21
         Top             =   2460
         Width           =   8160
         Begin VB.OptionButton OptAmbos 
            Caption         =   "Ambos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6390
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton OptCompras 
            Caption         =   "Compras"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3945
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton OptVentas 
            Caption         =   "Ventas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1380
            TabIndex        =   6
            Top             =   360
            Width           =   1170
         End
         Begin CATControls.CATTextBox txttipo 
            Height          =   315
            Left            =   7605
            TabIndex        =   22
            Tag             =   "Tvisible_ventas"
            Top             =   180
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
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
            MaxLength       =   3
            Container       =   "frmMantFormasPago.frx":55F8
            EnterTab        =   -1  'True
         End
      End
      Begin VB.CommandButton cmbAyudaFormaPago 
         Height          =   315
         Left            =   8100
         Picture         =   "frmMantFormasPago.frx":5614
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1290
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_FormaPago 
         Height          =   315
         Left            =   7590
         TabIndex        =   10
         Tag             =   "TidFormaPago"
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
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
         Container       =   "frmMantFormasPago.frx":599E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_FormaPago 
         Height          =   315
         Left            =   1995
         TabIndex        =   2
         Tag             =   "TGlsFormaPago"
         Top             =   855
         Width           =   6510
         _ExtentX        =   11483
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
         MaxLength       =   255
         Container       =   "frmMantFormasPago.frx":59BA
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_TipoFormaPago 
         Height          =   315
         Left            =   1995
         TabIndex        =   3
         Tag             =   "TidTipoFormaPago"
         Top             =   1275
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
         Container       =   "frmMantFormasPago.frx":59D6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TipoFormaPago 
         Height          =   315
         Left            =   2940
         TabIndex        =   18
         Top             =   1275
         Width           =   5130
         _ExtentX        =   9049
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
         Container       =   "frmMantFormasPago.frx":59F2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_DiasVcto 
         Height          =   315
         Left            =   1995
         TabIndex        =   4
         Tag             =   "NdiasVcto"
         Top             =   1680
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
         Alignment       =   1
         FontName        =   "Arial"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         MaxLength       =   3
         Container       =   "frmMantFormasPago.frx":5A0E
         Estilo          =   3
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txt_cheque 
         Height          =   315
         Left            =   6660
         TabIndex        =   23
         Tag             =   "Tcheque"
         Top             =   2070
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
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
         MaxLength       =   3
         Container       =   "frmMantFormasPago.frx":5A2A
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Días Vcto."
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
         Left            =   360
         TabIndex        =   20
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo Forma de Pago"
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
         Left            =   360
         TabIndex        =   19
         Top             =   1290
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
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
         Left            =   360
         TabIndex        =   12
         Top             =   915
         Width           =   855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   6990
         TabIndex        =   11
         Top             =   255
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1164
      ButtonWidth     =   2064
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Nuevo      "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMantFormasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_cheque_Click()

    If Chk_cheque.Value = 1 Then
       txt_cheque.Text = "S"
    Else
       txt_cheque.Text = "N"
    End If

End Sub

Private Sub cmbAyudaFormaPago_Click()
    
    mostrarAyuda "TIPOFORMASPAGO", txtCod_TipoFormaPago, txtGls_TipoFormaPago

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    Me.left = 0
    Me.top = 0
    ConfGrid gLista, False, False, False, False
    listaFP StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    nuevo
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo As String
Dim strMsg      As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaHomonimia "formaspagos", "GlsFormaPago", "idFormaPago", txtGls_FormaPago.Text, txtCod_FormaPago.Text, True, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_FormaPago.Text = "" Then 'graba
        txtCod_FormaPago.Text = GeneraCorrelativoAnoMes("formaspagos", "idFormaPago")
        EjecutaSQLFormFP Me, 0, True, "formaspagos", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabo"
    
    Else 'modifica
        EjecutaSQLFormFP Me, 1, True, "formaspagos", StrMsgError, "idFormaPago"
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modifico"
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    listaFP StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()
    
    limpiaForm Me
    txt_DiasVcto.Text = 0

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarFP gLista.Columns.ColumnByName("idFormaPago").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub

Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub OptAmbos_Click()
    
    txttipo.Text = "A"

End Sub

Private Sub OptCompras_Click()
    
    txttipo.Text = "C"

End Sub

Private Sub OptVentas_Click()
    
    txttipo.Text = "V"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            nuevo
            OptVentas.Value = True
            txttipo.Text = "V"
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            fraGeneral.Enabled = True
        Case 4, 7 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
            
            listaFP StrMsgError
            If StrMsgError <> "" Then GoTo Err
    
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 6 'Imprimir
            gLista.m.ExportToXLS App.Path & "\Temporales\Mantenimiento_FormaPago.xls"
            ShellEx App.Path & "\Temporales\Mantenimiento_FormaPago.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        Case 8 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub habilitaBotones(indexBoton As Integer)
Dim indHabilitar As Boolean

    Select Case indexBoton
        Case 1, 2, 3 'Nuevo, Grabar, Modificar
            If indexBoton = 2 Then indHabilitar = True
            Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
            Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
            Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
            Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
            Toolbar1.Buttons(5).Visible = indHabilitar 'Eliminar
            Toolbar1.Buttons(6).Visible = indHabilitar 'Imprimir
            Toolbar1.Buttons(7).Visible = indHabilitar 'Lista
        Case 4, 7 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
    End Select

End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaFP StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaFP(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND f.GlsFormaPago LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT f.idFormaPago ,f.GlsFormaPago,t.GlsTipoFormaPago,diasVcto " & _
           "FROM formaspagos f,tipoformaspago t " & _
           "WHERE f.idTipoFormaPago = t.idTipoFormaPago AND f.idEmpresa = '" & glsEmpresa & "'"
    
    If strCond <> "" Then csql = csql & strCond
    csql = csql & " ORDER BY f.idFormaPago"
    
If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
    
Set gLista.DataSource = rsdatos

'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "idFormaPago"
'    End With

    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarFP(strCodFP As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT f.idFormaPago,f.GlsFormaPago,f.idTipoFormaPago,f.diasVcto,f.visible_ventas,f.cheque " & _
           "FROM formaspagos f " & _
           "WHERE f.idFormaPago = '" & strCodFP & "' AND f.idEmpresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    If Not rst.EOF Then
        If rst.Fields("visible_ventas").Value = "V" Then
            OptVentas.Value = True
        ElseIf rst.Fields("visible_ventas").Value = "C" Then
            OptCompras.Value = True
        ElseIf rst.Fields("visible_ventas").Value = "A" Then
            OptAmbos.Value = True
        ElseIf IsNull(rst.Fields("visible_ventas").Value) = True Or rst.Fields("visible_ventas").Value = "" Then
            OptAmbos.Value = False
            OptCompras.Value = False
            OptVentas.Value = False
        End If
        
        If rst.Fields("cheque").Value & "" = "S" Then
            Chk_cheque.Value = 1
        Else
            Chk_cheque.Value = 0
        End If
    End If
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_TipoFormaPago_Change()

    txtGls_TipoFormaPago.Text = traerCampo("tipoformaspago", "GlsTipoFormaPago", "idTipoFormaPago", txtCod_TipoFormaPago.Text, False)
    If txtCod_FormaPago.Text = "06090002" Then
        txt_DiasVcto.Vacio = False
    Else
        txt_DiasVcto.Vacio = True
        txt_DiasVcto.Text = 0
    End If
    
End Sub

Private Sub txtCod_TipoFormaPago_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "TIPOFORMASPAGO", txtCod_TipoFormaPago, txtGls_TipoFormaPago
        KeyAscii = 0
        If txtCod_TipoFormaPago.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset
Dim CArrLog(1)          As String

    If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    strCodigo = Trim(txtCod_FormaPago.Text)
    
    '--- Validando
    csql = "SELECT Item FROM pagosdocventas WHERE idFormadePago = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Pagos)"
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    CArrLog(0) = "FormasPagos"
                    
    Graba_Log_Nuevo StrMsgError, Cn, CArrLog, "IdEmpresa = '" & glsEmpresa & "' And IdFormaPago = '" & txtCod_FormaPago.Text & "'", "E", 1
    If StrMsgError <> "" Then GoTo Err
        
    '--- Eliminando el registro
    csql = "DELETE FROM formaspagos WHERE idFormaPago = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    '--- Nuevo
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub
    
Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub EjecutaSQLFormFP(F As Form, tipoOperacion As Integer, indEmpresa As Boolean, strTabla As String, ByRef StrMsgError As String, Optional strCampoCod As String, Optional g As dxDBGrid, Optional strTablaDet As String, Optional strCampoDet As String, Optional strDataCampo As String, Optional indFechaRegistro As Boolean = False)
On Error GoTo Err
Dim C As Object
Dim csql As String
Dim strCampo As String
Dim strTipoDato As String
Dim strCampos As String
Dim strValores As String
Dim strValCod As String
Dim strCampoEmpresa As String
Dim strValorEmpresa As String
Dim strCondEmpresa As String
Dim strCampoFecReg As String
Dim strValorFecReg As String
Dim GlsObsCli   As String
Dim indTrans As Boolean
Dim CArrLog(2)                      As String

    If indEmpresa Then
        strCampoEmpresa = ",idEmpresa"
        strValorEmpresa = ",'" & glsEmpresa & "'"
        strCondEmpresa = " AND idEmpresa = '" & glsEmpresa & "'"
    End If
    
    If indFechaRegistro Then
        strCampoFecReg = ",FecRegistro"
        strValorFecReg = ",Getdate()"
    End If

    indTrans = False
    csql = ""
    For Each C In F.Controls
        If TypeOf C Is CATTextBox Or TypeOf C Is DTPicker Or TypeOf C Is CheckBox Then
            If C.Tag <> "" Then
                strTipoDato = left(C.Tag, 1)
                strCampo = right(C.Tag, Len(C.Tag) - 1)
                Select Case tipoOperacion
                    Case 0 'inserta
                        strCampos = strCampos & strCampo & ","
                        
                        Select Case strTipoDato
                            Case "N"
                                strValores = strValores & C.Value & ","
                            Case "T"
                                strValores = strValores & "'" & Trim(C.Value) & "',"
                            Case "F"
                                strValores = strValores & "'" & Format(C.Value, "yyyy-mm-dd") & "',"
                        End Select
                    Case 1
                        If UCase(strCampoCod) <> UCase(strCampo) Then
                            Select Case strTipoDato
                                Case "N"
                                    strValores = C.Value
                                Case "T"
                                    strValores = "'" & C.Value & "'"
                                Case "F"
                                    strValores = "'" & Format(C.Value, "yyyy-mm-dd") & "'"
                            End Select
                            strCampos = strCampos & strCampo & "=" & strValores & ","
                        Else
                            strValCod = C.Value
                        End If
                End Select
            End If
        End If
    Next
    
    If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
    
    Select Case tipoOperacion
        Case 0
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            csql = "INSERT INTO " & strTabla & "(" & strCampos & strCampoEmpresa & strCampoFecReg & ") VALUES(" & strValores & strValorEmpresa & strValorFecReg & ")"
        Case 1
            csql = "UPDATE " & strTabla & " SET " & strCampos & " WHERE " & strCampoCod & " = '" & strValCod & "'" & strCondEmpresa
    End Select

    indTrans = True
    Cn.BeginTrans
    
    If tipoOperacion = 1 Then
        
        CArrLog(0) = "FormasPagos"
                    
        Graba_Log_Nuevo StrMsgError, Cn, CArrLog, "IdEmpresa = '" & glsEmpresa & "' And IdFormaPago = '" & txtCod_FormaPago.Text & "'", "A", 1
        If StrMsgError <> "" Then GoTo Err
        
    End If
    
    '--- Graba controles
    If strCampos <> "" Then
        Cn.Execute csql
    End If
    
    GlsObsCli = "" & glsobservacioncliente
    
    If UCase(strTabla) = "CLIENTES" Then
        csql = "Update Clientes Set GlsObservacion = '" & GlsObsCli & "' Where Idcliente ='" & strDataCampo & "' And idEmpresa= '" & glsEmpresa & "' "
               Cn.Execute (csql)
    ElseIf UCase(strTabla) = "PROVEEDORES" Then
        csql = "Update proveedores Set GlsObservacion = '" & GlsObsCli & "' Where Idproveedor ='" & strDataCampo & "' And idEmpresa= '" & glsEmpresa & "' "
               Cn.Execute (csql)
    End If
    
    '--- Grabando Grilla
    If TypeName(g) <> "Nothing" Then
        Cn.Execute "DELETE FROM " & strTablaDet & " WHERE " & strCampoDet & " = '" & strDataCampo & "'" & strCondEmpresa
        
        g.Dataset.First
        Do While Not g.Dataset.EOF
            strCampos = ""
            strValores = ""
            For i = 0 To g.Columns.Count - 1
                If UCase(left(g.Columns(i).ObjectName, 1)) = "W" Then
                    strTipoDato = Mid(g.Columns(i).ObjectName, 2, 1)
                    strCampo = Mid(g.Columns(i).ObjectName, 3)
                    
                    strCampos = strCampos & strCampo & ","
                    
                    Select Case strTipoDato
                        Case "N"
                            strValores = strValores & g.Columns(i).Value & ","
                        Case "T"
                            strValores = strValores & "'" & Trim(g.Columns(i).Value) & "',"
                        Case "F"
                            strValores = strValores & "'" & Format(g.Columns(i).Value, "yyyy-mm-dd") & "',"
                    End Select
                End If
            Next
            
            If Len(strCampos) > 1 Then strCampos = left(strCampos, Len(strCampos) - 1)
            If Len(strValores) > 1 Then strValores = left(strValores, Len(strValores) - 1)
            
            csql = "INSERT INTO " & strTablaDet & "(" & strCampos & "," & strCampoDet & strCampoEmpresa & ") VALUES(" & strValores & ",'" & strDataCampo & "'" & strValorEmpresa & ")"
            Cn.Execute csql
            
            g.Dataset.Next
        Loop
    End If
    
    Cn.CommitTrans
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    If indTrans Then Cn.RollbackTrans
End Sub

Private Sub Graba_Log_Nuevo(StrMsgError As String, PCn As ADODB.Connection, PTabla() As String, PWhere As String, PAccion As String, PCantArr As Long)
On Error GoTo Err
Dim cinsert                             As String
Dim cselect                             As String
Dim RsConsulta                          As New ADODB.Recordset
Dim i                                   As Long
Dim CCampos                             As String
Dim X                                   As Long
Dim NItem_Log                           As Long
    
    cselect = "Select ItemLog " & _
              "From " & PTabla(0) & "_Log " & _
              "Where " & PWhere & " " & _
              "Order By ItemLog Desc"
    RsConsulta.Open cselect, PCn, adOpenStatic, adLockReadOnly
    If Not RsConsulta.EOF Then
        NItem_Log = Val("" & RsConsulta.Fields("ItemLog")) + 1
    Else
        NItem_Log = 1
    End If
    RsConsulta.Close: Set RsConsulta = Nothing
    
    For X = 0 To PCantArr - 1
        cselect = "Select * From " & PTabla(X) & " Where " & PWhere
        With RsConsulta
            .Open cselect, PCn, adOpenStatic, adLockReadOnly
            If Not .EOF Then
                
                cinsert = ""
                cselect = ""
                
                For i = 0 To .Fields.Count - 1
                    cinsert = cinsert & .Fields(i).Name & ","
                    cselect = cselect & .Fields(i).Name & ","
                Next i
            
                cinsert = "Insert Into " & PTabla(X) & "_Log" & "(" & cinsert & IIf(X = 0, "AccionLog,FechaLog,HoraLog,IdUsuarioLog,GlsPCLog,GlsPCUsuarioLog,", "") & "ItemLog)"
                          
                cselect = "Select " & cselect & " " & _
                          IIf(X = 0, "'" & PAccion & "',Cast(getdate() As Date),Cast(getdate() As Time),'" & glsUser & "','" & ComputerName & "','" & fpUsuarioActual & "',", "") & _
                          "" & NItem_Log & " " & _
                          "From " & PTabla(X) & " Where " & PWhere
                PCn.Execute cinsert & cselect
            End If
            .Close
        End With
    Next X
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

