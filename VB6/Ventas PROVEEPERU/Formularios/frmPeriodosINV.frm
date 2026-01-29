VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmPeriodosINV 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Períodos de Inventario"
   ClientHeight    =   5025
   ClientLeft      =   4620
   ClientTop       =   2010
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   420
      Top             =   2520
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
            Picture         =   "frmPeriodosINV.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriodosINV.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   60
      TabIndex        =   8
      Top             =   630
      Width           =   7950
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   7725
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1020
            TabIndex        =   0
            Top             =   210
            Width           =   6585
            _ExtentX        =   11615
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
            Container       =   "frmPeriodosINV.frx":3518
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
            Left            =   165
            TabIndex        =   10
            Top             =   255
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3225
         Left            =   120
         OleObjectBlob   =   "frmPeriodosINV.frx":3534
         TabIndex        =   1
         Top             =   960
         Width           =   7740
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   60
      TabIndex        =   4
      Top             =   630
      Width           =   7950
      Begin CATControls.CATTextBox txtCod_PeriodoINV 
         Height          =   315
         Left            =   6435
         TabIndex        =   5
         Tag             =   "TidPeriodoInv"
         Top             =   240
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
         MaxLength       =   8
         Container       =   "frmPeriodosINV.frx":4C78
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_PeriodoINV 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Tag             =   "TglsPeriodoInv"
         Top             =   1035
         Width           =   6135
         _ExtentX        =   10821
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
         Container       =   "frmPeriodosINV.frx":4C94
         Estilo          =   1
         EnterTab        =   -1  'True
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
         Left            =   210
         TabIndex        =   7
         Top             =   1095
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
         Left            =   5790
         TabIndex        =   6
         Top             =   270
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   2170
      ButtonWidth     =   2540
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Nuevo         "
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
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmPeriodosINV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    
    ConfGrid gLista, False, False, False, False
    
    listaPeriodoInv StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 6
    nuevo
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim objVales As New clsVales
Dim strCodigo As String
Dim strMsg      As String
Dim indAcepto As Boolean, strMoneda As String, strFechaCorte As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaHomonimia "PeriodosINV", "GlsPeriodoInv", "idPeriodoInv", txtGls_PeriodoINV.Text, txtCod_PeriodoINV.Text, True, StrMsgError, " idSucursal = '" & glsSucursal & "'"
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_PeriodoINV.Text = "" Then 'graba
        indAcepto = False
        If MsgBox("¿Desea generar saldos iniciales para el nuevo periodo?", vbYesNo + vbInformation, App.Title) = vbYes Then
            frmPeriodosINV_Parametros.MostrarForm indAcepto, strMoneda, strFechaCorte
        End If
        txtCod_PeriodoINV.Text = GeneraCorrelativoAnoMes("PeriodosINV", "idPeriodoInv")
        
        objVales.EjecutaSQLFormPeriodoINV Me, 0, StrMsgError, IIf(indAcepto, 1, 0), strMoneda, strFechaCorte
        If StrMsgError <> "" Then GoTo Err
        
        strMsg = "Grabo"
    
    Else 'modifica
        objVales.EjecutaSQLFormPeriodoINV Me, 1, StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modifico"
    End If
    
    Set objVales = Nothing
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    listaPeriodoInv StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    Set objVales = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()
    
    limpiaForm Me

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarPeriodoInv gLista.Columns.ColumnByName("idPeriodoInv").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            nuevo
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            fraGeneral.Enabled = True
        Case 4, 6  'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        Case 5 'Imprimir
        Case 7 'Salir
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
            Toolbar1.Buttons(5).Visible = indHabilitar 'Imprimir
            Toolbar1.Buttons(6).Visible = indHabilitar 'Lista
        Case 4, 6 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
    End Select

End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaPeriodoInv StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaPeriodoInv(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND m.GlsPeriodoInv LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT m.idPeriodoInv ,m.GlsPeriodoInv " & _
           "FROM PeriodosInv m WHERE m.idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'"
           
    If strCond <> "" Then csql = csql & strCond
    csql = csql & " ORDER BY m.idPeriodoInv"
    
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
'        .KeyField = "idPeriodoInv"
'    End With
    
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarPeriodoInv(strCod As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT m.idPeriodoInv,m.GlsPeriodoInv " & _
           "FROM PeriodosInv m " & _
           "WHERE m.idPeriodoInv = '" & strCod & "' AND m.idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
