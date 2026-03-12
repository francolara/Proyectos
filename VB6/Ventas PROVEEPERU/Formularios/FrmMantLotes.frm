VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmMantLotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Tallas"
   ClientHeight    =   7485
   ClientLeft      =   4590
   ClientTop       =   1830
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   6720
      Left            =   45
      TabIndex        =   6
      Top             =   720
      Width           =   8115
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   120
         TabIndex        =   7
         Top             =   150
         Width           =   7860
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1035
            TabIndex        =   0
            Top             =   255
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
            Container       =   "FrmMantLotes.frx":0000
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Busqueda"
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
            TabIndex        =   8
            Top             =   300
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   5610
         Left            =   120
         OleObjectBlob   =   "FrmMantLotes.frx":001C
         TabIndex        =   1
         Top             =   960
         Width           =   7860
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   75
      TabIndex        =   9
      Top             =   1950
      Width           =   8100
      Begin VB.Frame Frame2 
         Caption         =   " Estado "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1440
         TabIndex        =   13
         Top             =   1485
         Width           =   5115
         Begin VB.OptionButton OpmDst 
            Caption         =   "Inactivo"
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
            Left            =   3690
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton OpmAct 
            Caption         =   "Activo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   450
            TabIndex        =   3
            Top             =   405
            Width           =   1095
         End
         Begin CATControls.CATTextBox txtestlote 
            Height          =   315
            Left            =   1680
            TabIndex        =   14
            Tag             =   "TEstado"
            Top             =   480
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
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
            MaxLength       =   8
            Container       =   "FrmMantLotes.frx":1724
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
      End
      Begin CATControls.CATTextBox txtCod_Lote 
         Height          =   315
         Left            =   6960
         TabIndex        =   10
         Tag             =   "TidLote"
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
         Container       =   "FrmMantLotes.frx":1740
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Lote 
         Height          =   315
         Left            =   1365
         TabIndex        =   2
         Tag             =   "TglsLote"
         Top             =   915
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
         Container       =   "FrmMantLotes.frx":175C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtsucursal 
         Height          =   315
         Left            =   3000
         TabIndex        =   15
         Tag             =   "Tidsucursal"
         Top             =   3240
         Visible         =   0   'False
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
         Container       =   "FrmMantLotes.frx":1778
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
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
         Left            =   6225
         TabIndex        =   12
         Top             =   390
         Width           =   495
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
         TabIndex        =   11
         Top             =   975
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
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
               Picture         =   "FrmMantLotes.frx":1794
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":1B2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":1F80
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":231A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":26B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":2A4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":2DE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":3182
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":351C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":38B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":3C50
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMantLotes.frx":4912
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmMantLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
    
    Me.top = 0
    Me.left = 0
    OpmAct.Value = True
    txtestlote.Text = "ACT"
    txtsucursal.Text = glsSucursal
    
    ConfGrid gLista, False, False, False, False
    listaLote StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
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

    mostrarLote gLista.Columns.ColumnByName("idlote").Value, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gLista_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
On Error GoTo Err
Dim StrMsgError                         As String
    
    If KeyCode = 116 Then
    
        listaLote StrMsgError
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub OpmAct_Click()
    
    txtestlote.Text = "ACT"

End Sub

Private Sub OpmDst_Click()
    
    txtestlote.Text = "INA"

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
            txtestlote.Text = "ACT"
            txtsucursal.Text = glsSucursal
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
            listaLote StrMsgError
            If StrMsgError <> "" Then GoTo Err
            
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

Private Sub listaLote(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND (idLote LIKE '%" & strCond & "%' Or GlsLote LIKE '%" & strCond & "%')"
    End If
    csql = "SELECT idLote,GlsLote FROM Lotes WHERE idEmpresa = '" & glsEmpresa & "' and idsucursal = '" & glsSucursal & "' "
           
    If strCond <> "" Then csql = csql & strCond
    csql = csql & " ORDER BY idLote"
    

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
'        .KeyField = "idLote"
'    End With
    
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaLote StrMsgError
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
    
    validaHomonimia "Lotes", "glsLote", "idlote", txtGls_Lote.Text, txtCod_Lote.Text, True, StrMsgError, " idsucursal = '" & glsSucursal & "' "
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_Lote.Text = "" Then 'graba
        txtCod_Lote.Text = GeneraCorrelativoAnoMes("lotes", "idlote")
        
        EjecutaSQLForm Me, 0, True, "lotes", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        strMsg = "Grabo"
        
    Else 'modifica
        EjecutaSQLForm Me, 1, True, "lotes", StrMsgError, "idlote"
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modifico"
    End If
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    fraGeneral.Enabled = False
    listaLote StrMsgError
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
    
    strCodigo = Trim(txtCod_Lote.Text)
    
    csql = "SELECT top 1 idlote FROM productosalmacenporlote WHERE idlote = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "' " & _
           "And IdSucursal = '" & glsSucursal & "' "
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso "
        GoTo Err
    End If
    
    rsValida.Close: Set rsValida = Nothing
    
    csql = "SELECT top 1 idlote FROM ValesDet WHERE idlote = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "' " & _
           "And IdSucursal = '" & glsSucursal & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso "
        GoTo Err
    End If
    
    rsValida.Close: Set rsValida = Nothing
    
    csql = "SELECT top 1 idlote FROM docventasdetlote WHERE idlote = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "' " & _
           "And IdSucursal = '" & glsSucursal & "' "
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso "
        GoTo Err
    End If
    
    rsValida.Close: Set rsValida = Nothing
    
    Cn.BeginTrans
    indTrans = True
    
    csql = "DELETE FROM lotes WHERE idlote = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "' and idsucursal = '" & glsSucursal & "' "
    Cn.Execute csql
    
    Cn.CommitTrans
    
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    
    listaLote StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub
    
Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarLote(strCodLote As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT idLote, GlsLote,Estado,idsucursal FROM Lotes " & _
           "WHERE idlote = '" & strCodLote & "' AND idEmpresa = '" & glsEmpresa & "' and idsucursal = '" & glsSucursal & "' "
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError                         As String
    
    If KeyCode = 116 Then
    
        listaLote StrMsgError
        If StrMsgError <> "" Then GoTo Err
    
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
