VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmMant_Canal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Canales"
   ClientHeight    =   5340
   ClientLeft      =   1950
   ClientTop       =   2445
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   4650
      Left            =   30
      TabIndex        =   8
      Top             =   630
      Width           =   8115
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   135
         TabIndex        =   9
         Top             =   150
         Width           =   7860
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1035
            TabIndex        =   0
            Top             =   210
            Width           =   6690
            _ExtentX        =   11800
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
            Container       =   "FrmMant_Canal.frx":0000
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
         Height          =   3540
         Left            =   120
         OleObjectBlob   =   "FrmMant_Canal.frx":001C
         TabIndex        =   1
         Top             =   960
         Width           =   7860
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   45
      TabIndex        =   4
      Top             =   630
      Width           =   8100
      Begin CATControls.CATTextBox txtCod_Canal 
         Height          =   315
         Left            =   6960
         TabIndex        =   5
         Tag             =   "TidCanal"
         Top             =   360
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
         Container       =   "FrmMant_Canal.frx":172E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Canal 
         Height          =   315
         Left            =   1185
         TabIndex        =   2
         Tag             =   "TglsCanal"
         Top             =   1455
         Width           =   6690
         _ExtentX        =   11800
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
         Container       =   "FrmMant_Canal.frx":174A
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
         Top             =   1515
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
         Left            =   6360
         TabIndex        =   6
         Top             =   390
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   2170
      ButtonWidth     =   2461
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Nuevo        "
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
            Object.ToolTipText     =   "Eliminar"
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
      Begin MSComctlLib.ImageList imgDocVentas 
         Left            =   7335
         Top             =   90
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
               Picture         =   "FrmMant_Canal.frx":1766
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":1B00
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":1F52
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":22EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":2686
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":2A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":2DBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":3154
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":34EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":3888
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":3C22
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMant_Canal.frx":48E4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmMant_Canal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    
    Me.top = 0
    Me.left = 0
    ConfGrid gLista, False, False, False, False
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarCanal gLista.Columns.ColumnByName("idCanal").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            limpiaForm Me
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            fraGeneral.Enabled = False
        Case 3 'Modificar
            fraGeneral.Enabled = True
        Case 4  'Cancelar
            fraGeneral.Enabled = False
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
            fraGeneral.Enabled = True
        Case 6 'Imprimir
            
        Case 7 'Lista
            fraListado.Visible = True
            fraGeneral.Visible = False
        Case 8 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub habilitaBotones(indexBoton As Integer)

    Select Case indexBoton
        Case 1, 5 'Nuevo, Eliminar
            Toolbar1.Buttons(1).Visible = False 'Nuevo
            Toolbar1.Buttons(2).Visible = True 'Grabar
            Toolbar1.Buttons(3).Visible = False 'Modificar
            Toolbar1.Buttons(4).Visible = False 'Cancelar
            Toolbar1.Buttons(5).Visible = False 'Eliminar
            Toolbar1.Buttons(6).Visible = False 'Imprimir
            Toolbar1.Buttons(7).Visible = True 'Lista
        Case 2, 4 'Grabar, Cancelar
            Toolbar1.Buttons(1).Visible = True 'Nuevo
            Toolbar1.Buttons(2).Visible = False 'Grabar
            Toolbar1.Buttons(3).Visible = True 'Modificar
            Toolbar1.Buttons(4).Visible = False 'Cancelar
            Toolbar1.Buttons(5).Visible = True 'Eliminar
            Toolbar1.Buttons(6).Visible = False 'Imprimir
            Toolbar1.Buttons(7).Visible = True 'Lista
        Case 3 'Modificar
            Toolbar1.Buttons(1).Visible = False 'Nuevo
            Toolbar1.Buttons(2).Visible = True 'Grabar
            Toolbar1.Buttons(3).Visible = False 'Modificar
            Toolbar1.Buttons(4).Visible = True 'Cancelar
            Toolbar1.Buttons(5).Visible = False 'Eliminar
            Toolbar1.Buttons(6).Visible = False 'Imprimir
            Toolbar1.Buttons(7).Visible = False 'Lista
        Case 7 'Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
    End Select

End Sub

Private Sub listaCanal(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND glsCanal LIKE '%" & strCond & "%'"
    End If
    csql = "SELECT idCanal, glsCanal FROM Canal WHERE idEmpresa = '" & glsEmpresa & "'"
    If strCond <> "" Then csql = csql & strCond
    csql = csql & " ORDER BY idCanal"
    
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
'        .KeyField = "idCanal"
'    End With
    
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaCanal StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
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
    
    validaHomonimia "Canal", "glsCanal", "idCanal", txtGls_Canal.Text, txtCod_Canal.Text, True, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_Canal.Text = "" Then 'graba
        txtCod_Canal.Text = GeneraCorrelativoAnoMes("Canal", "idCanal")
        EjecutaSQLForm Me, 0, True, "Canal", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabo"
    
    Else 'modifica
        EjecutaSQLForm Me, 1, True, "Canal", StrMsgError, "idCanal"
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modifico"
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    listaCanal StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    strCodigo = Trim(txtCod_Canal.Text)
    
    csql = "SELECT idCanal FROM Clientes WHERE idCanal = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Clientes)"
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando Registro
    csql = "DELETE FROM Canal WHERE idCanal = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    Cn.Execute csql
    
    Cn.CommitTrans
    
    '--- Nuevo
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    listaCanal StrMsgError
    If StrMsgError <> "" Then GoTo Err
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub
    
Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarCanal(strCodCanal As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT idCanal, GlsCanal FROM Canal " & _
           "WHERE idCanal = '" & strCodCanal & "' AND idEmpresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
