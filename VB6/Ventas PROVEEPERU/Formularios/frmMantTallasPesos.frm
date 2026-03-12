VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantTallasPesos 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Tallas"
   ClientHeight    =   4725
   ClientLeft      =   5865
   ClientTop       =   2805
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmMantTallasPesos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTallasPesos.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      Caption         =   "Listado"
      ForeColor       =   &H00000000&
      Height          =   4020
      Left            =   45
      TabIndex        =   1
      Top             =   600
      Width           =   7980
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   120
         TabIndex        =   2
         Top             =   165
         Width           =   7770
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   960
            TabIndex        =   3
            Top             =   255
            Width           =   6690
            _ExtentX        =   11800
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
            MaxLength       =   255
            Container       =   "frmMantTallasPesos.frx":3518
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Busqueda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Width           =   765
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   2955
         Left            =   120
         OleObjectBlob   =   "frmMantTallasPesos.frx":3534
         TabIndex        =   5
         Top             =   960
         Width           =   7770
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   2170
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
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   45
      TabIndex        =   6
      Top             =   600
      Width           =   7965
      Begin CATControls.CATTextBox txtCod_TallaPeso 
         Height          =   315
         Left            =   6735
         TabIndex        =   7
         Tag             =   "TidTallaPeso"
         Top             =   360
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
         Container       =   "frmMantTallasPesos.frx":4C56
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TallaPeso 
         Height          =   315
         Left            =   1365
         TabIndex        =   8
         Tag             =   "TGlsTallaPeso"
         Top             =   1080
         Width           =   6285
         _ExtentX        =   11086
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
         MaxLength       =   255
         Container       =   "frmMantTallasPesos.frx":4C72
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   345
         TabIndex        =   10
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
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
         Left            =   6000
         TabIndex        =   9
         Top             =   390
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmMantTallasPesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim StrMsgError As String
On Error GoTo Err

ConfGrid gLista, False, False, False, False

listaTallaPeso StrMsgError
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
Dim strCodigo As String
Dim strMsg      As String
On Error GoTo Err

validaFormSQL Me, StrMsgError
If StrMsgError <> "" Then GoTo Err

validaHomonimia "tallapeso", "GlsTallaPeso", "idTallaPeso", txtGls_TallaPeso.Text, txtCod_TallaPeso.Text, True, StrMsgError
If StrMsgError <> "" Then GoTo Err

If txtCod_TallaPeso.Text = "" Then 'graba
    txtCod_TallaPeso.Text = GeneraCorrelativoAnoMes("tallapeso", "idTallaPeso")
    
    EjecutaSQLForm Me, 0, True, "tallapeso", StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    strMsg = "Grabo"
Else 'modifica

    EjecutaSQLForm Me, 1, True, "tallapeso", StrMsgError, "idTallaPeso"
    If StrMsgError <> "" Then GoTo Err
    
    strMsg = "Modifico"
End If
MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title

fraGeneral.Enabled = False

listaTallaPeso StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()
    limpiaForm Me
End Sub

Private Sub gLista_OnDblClick()
Dim StrMsgError As String
On Error GoTo Err
mostrarTallaPeso gLista.Columns.ColumnByName("idTallaPeso").Value, StrMsgError
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
Dim StrMsgError As String
On Error GoTo Err
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
    Case 4, 7  'Cancelar
        fraListado.Visible = True
        fraGeneral.Visible = False
        fraGeneral.Enabled = False
    Case 5 'Eliminar
        eliminar StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Case 6 'Imprimir
    
'    Case 7 'Lista
'        fraListado.Visible = True
'        fraGeneral.Visible = False
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
        Toolbar1.Buttons(6).Visible = False
        Toolbar1.Buttons(7).Visible = False
End Select

End Sub

Private Sub txt_TextoBuscar_Change()
Dim StrMsgError As String
On Error GoTo Err
listaTallaPeso StrMsgError
If StrMsgError <> "" Then GoTo Err
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then gLista.SetFocus
End Sub

Private Sub listaTallaPeso(ByRef StrMsgError As String)
Dim strCond As String
On Error GoTo Err
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND m.GlsTallaPeso LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT m.idTallaPeso ,m.GlsTallaPeso " & _
           "FROM tallapeso m WHERE m.idEmpresa = '" & glsEmpresa & "'"
           
    If strCond <> "" Then csql = csql + strCond

    csql = csql + " ORDER BY m.idTallaPeso"
    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn '''Cn
        
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idTallaPeso"
End With
Me.Refresh
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarTallaPeso(strCod As String, ByRef StrMsgError As String)
Dim rst As New ADODB.Recordset
On Error GoTo Err
    csql = "SELECT m.idTallaPeso,m.GlsTallaPeso " & _
           "FROM tallapeso m " & _
           "WHERE m.idTallaPeso = '" & strCod & "' AND m.idEmpresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
Me.Refresh
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

On Error GoTo Err

If MsgBox("¿Seguro de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub

strCodigo = Trim(txtCod_TallaPeso.Text)

'*********************************************VALIDANDO*****************************************************
csql = "SELECT idTallaPeso FROM productos WHERE idTallaPeso = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
If Not rsValida.EOF Then
    StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Productos)"
    GoTo Err
End If
'***********************************************************************************************************

Cn.BeginTrans
indTrans = True

'*********************************************ELIMINANDO****************************************************
'Eliminando el registro
csql = "DELETE FROM tallapeso WHERE idTallaPeso = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
Cn.Execute csql

Cn.CommitTrans

'Nuevo
Toolbar1_ButtonClick Toolbar1.Buttons(1)

MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title

If rsValida.State = 1 Then rsValida.Close
Set rsValida = Nothing
Exit Sub
Err:
If rsValida.State = 1 Then rsValida.Close
Set rsValida = Nothing
If indTrans Then Cn.RollbackTrans
If StrMsgError = "" Then StrMsgError = Err.Description
End Sub





