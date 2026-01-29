VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantTiposNiveles 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Tipos Niveles"
   ClientHeight    =   4275
   ClientLeft      =   3345
   ClientTop       =   1485
   ClientWidth     =   7710
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   500
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
            Picture         =   "frmMantJerarquias.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantJerarquias.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   3525
      Left            =   45
      TabIndex        =   11
      Top             =   675
      Width           =   7620
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   120
         TabIndex        =   12
         Top             =   150
         Width           =   7365
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1050
            TabIndex        =   0
            Top             =   210
            Width           =   6150
            _ExtentX        =   10848
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
            Container       =   "frmMantJerarquias.frx":3518
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
            Left            =   210
            TabIndex        =   13
            Top             =   255
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   2550
         Left            =   120
         OleObjectBlob   =   "frmMantJerarquias.frx":3534
         TabIndex        =   1
         Top             =   870
         Width           =   7410
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   45
      TabIndex        =   5
      Top             =   675
      Width           =   7605
      Begin CATControls.CATTextBox txtCod_TipoNivel 
         Height          =   315
         Left            =   6600
         TabIndex        =   6
         Tag             =   "TidTipoNivel"
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
         Container       =   "frmMantJerarquias.frx":5101
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_TipoNivel 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Tag             =   "TglsTipoNivel"
         Top             =   1035
         Width           =   5970
         _ExtentX        =   10530
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
         Container       =   "frmMantJerarquias.frx":511D
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtVal_Peso 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Tag             =   "NPeso"
         Top             =   1425
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
         MaxLength       =   8
         Container       =   "frmMantJerarquias.frx":5139
         Estilo          =   3
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_BaseDatos 
         Height          =   315
         Left            =   5475
         TabIndex        =   4
         Top             =   2985
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
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
         Container       =   "frmMantJerarquias.frx":5155
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Base de Datos"
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
         Left            =   4275
         TabIndex        =   14
         Top             =   3045
         Visible         =   0   'False
         Width           =   1065
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
         Left            =   270
         TabIndex        =   9
         Top             =   1050
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
         Left            =   5955
         TabIndex        =   8
         Top             =   390
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nivel"
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
         Left            =   270
         TabIndex        =   7
         Top             =   1425
         Width           =   345
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   2170
      ButtonWidth     =   2540
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
            Caption         =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
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
Attribute VB_Name = "frmMantTiposNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLista, False, False, False, False
    listaTipoNivel StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    nuevo
    txtGls_BaseDatos.Text = "dbcomer"
    
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
    
    validaHomonimia "tiposniveles", "GlsTipoNivel", "idTipoNivel", txtGls_TipoNivel.Text, txtCod_TipoNivel.Text, True, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_TipoNivel.Text = "" Then '--- Graba
        txtCod_TipoNivel.Text = GeneraCorrelativoAnoMes("tiposniveles", "idTipoNivel")
        EjecutaSQLForm Me, 0, True, "tiposniveles", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabo"
    
    Else '--- Modifica
        EjecutaSQLForm Me, 1, True, "tiposniveles", StrMsgError, "idTipoNivel"
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modifico"
    End If
    
    generaVistaNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    fraGeneral.Enabled = False
    listaTipoNivel StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()
    
    limpiaForm Me

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarTipoNivel gLista.Columns.ColumnByName("idTipoNivel").Value, StrMsgError
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
            nuevo
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            fraGeneral.Enabled = True
        Case 4 'Eliminar
            
            If MsgBox("¿Seguro de eliminar el Registro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                EliminarRegistro StrMsgError
                If StrMsgError <> "" Then GoTo Err
            End If
        Case 5, 7 'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        Case 6 'Imprimir
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
            Toolbar1.Buttons(4).Visible = indHabilitar 'Eliminar
            Toolbar1.Buttons(5).Visible = Not indHabilitar 'Cancelar
            Toolbar1.Buttons(6).Visible = indHabilitar 'Imprimir
            Toolbar1.Buttons(7).Visible = indHabilitar 'Lista
        Case 5, 7 'Cancelar, Lista
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
On Error GoTo Err
Dim StrMsgError As String

    listaTipoNivel StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaTipoNivel(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " WHERE n.GlsTipoNivel LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT n.idTipoNivel ,n.GlsTipoNivel,n.Peso " & _
           "FROM tiposniveles n WHERE n.idEmpresa = '" & glsEmpresa & "'"
    If strCond <> "" Then csql = csql & strCond
    csql = csql & " ORDER BY n.Peso ASC"
    
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
'        .KeyField = "idTipoNivel"
'    End With
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarTipoNivel(strCodTipoNivel As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT n.idTipoNivel,n.GlsTipoNivel,n.Peso " & _
           "FROM tiposniveles n " & _
           "WHERE n.idEmpresa = '" & glsEmpresa & "' AND n.idTipoNivel = '" & strCodTipoNivel & "' ORDER BY Peso ASC"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub generaVistaNiveles(ByRef StrMsgError As String)
On Error GoTo Err
Dim rse As New ADODB.Recordset
Dim NumNivelesMax As Integer
Dim NumNiveles As Integer
Dim strCodEmpresa As String
Dim i As Integer
Dim strTablas As String
Dim strWhere As String
Dim strCampos As String
Dim strVista As String

    csql = "SELECT e.idEmpresa, Count(idTipoNivel) AS NumNiveles " & _
           "FROM empresas e, tiposniveles t " & _
           "WHERE E.idEmpresa = t.idEmpresa " & _
           "GROUP BY e.idEmpresa " & _
           "ORDER BY Count(idTipoNivel) DESC, e.idEmpresa ASC"
    rse.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    NumNivelesMax = -1
    Do While Not rse.EOF
        strCodEmpresa = "" & rse.Fields("idEmpresa")
        NumNiveles = "" & rse.Fields("NumNiveles")
        
        If NumNivelesMax = -1 Then NumNivelesMax = NumNiveles
        strTablas = ""
        strCampos = ""
        strWhere = ""
        
        For i = 1 To NumNiveles
            If i = 1 Then
                strTablas = "niveles n" & CStr(i)
                strWhere = ""
            Else
                strCampos = strCampos & ","
                strTablas = strTablas & ",niveles n" & CStr(i)
                If i > 2 Then strWhere = strWhere & " AND "
                strWhere = strWhere & "n" & CStr(i - 1) & ".idNivelPred = n" & CStr(i) & ".idNivel AND n" & CStr(i - 1) & ".idEmpresa = n" & CStr(i) & ".idEmpresa"
            End If
            strCampos = strCampos & "n" & CStr(i) & ".idNivel AS idNivel0" & CStr(i) & ", n" & CStr(i) & ".GlsNivel AS GlsNivel0" & CStr(i) & ""
        Next
        
        For i = 1 To (NumNivelesMax - NumNiveles)
            strCampos = strCampos & ",'',''"
        Next
        
        If strWhere <> "" Then
            strWhere = " WHERE n1.idEmpresa = '" & strCodEmpresa & "' AND " & strWhere
        Else
            strWhere = " WHERE n1.idEmpresa = '" & strCodEmpresa & "'"
        End If
            
        If strVista <> "" Then strVista = strVista & " UNION ALL "
        
        strVista = strVista & "SELECT n1.idEmpresa," & strCampos & _
                    " FROM " & strTablas & " " & strWhere
        rse.MoveNext
    Loop
    
    strVista = "CREATE OR ALTER VIEW vw_niveles AS " & strVista & ";"
    Cn.Execute strVista
    
    If rse.State = 1 Then rse.Close: Set rse = Nothing
    
    Exit Sub
    
Err:
    If rse.State = 1 Then rse.Close: Set rse = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub EliminarRegistro(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo As String
Dim strMsg      As String

    If txtCod_TipoNivel.Text <> "" Then 'Elimina
        csql = "DELETE FROM tiposniveles WHERE idEmpresa = '" & glsEmpresa & "' AND idTipoNivel = '" & txtCod_TipoNivel.Text & "'"
        Cn.Execute csql
        strMsg = "Elimino"
    End If
    
    generaVistaNiveles StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    nuevo
    habilitaBotones 7
    fraListado.Visible = True
    fraGeneral.Visible = False
    fraGeneral.Enabled = False
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    listaTipoNivel StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
