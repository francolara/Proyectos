VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantDirRecojo 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direcciones de Recojo"
   ClientHeight    =   7635
   ClientLeft      =   3975
   ClientTop       =   3000
   ClientWidth     =   8430
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
   ScaleHeight     =   7635
   ScaleWidth      =   8430
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
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
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   9000
      Top             =   4170
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
            Picture         =   "frmMantDirRecojo.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantDirRecojo.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6945
      Left            =   30
      TabIndex        =   7
      Top             =   600
      Width           =   8385
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   120
         TabIndex        =   8
         Top             =   165
         Width           =   8175
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   1005
            TabIndex        =   0
            Top             =   210
            Width           =   7050
            _ExtentX        =   12435
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
            Container       =   "frmMantDirRecojo.frx":3518
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Búsqueda"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   9
            Top             =   255
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   5835
         Left            =   120
         OleObjectBlob   =   "frmMantDirRecojo.frx":3534
         TabIndex        =   1
         Top             =   960
         Width           =   8175
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6960
      Left            =   30
      TabIndex        =   3
      Top             =   600
      Width           =   8370
      Begin VB.CommandButton cmbAyudaPais 
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
         Left            =   7860
         Picture         =   "frmMantDirRecojo.frx":4C56
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1650
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaDepa 
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
         Left            =   7860
         Picture         =   "frmMantDirRecojo.frx":4FE0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2040
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaProv 
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
         Left            =   7860
         Picture         =   "frmMantDirRecojo.frx":536A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2415
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaDistrito 
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
         Left            =   7860
         Picture         =   "frmMantDirRecojo.frx":56F4
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2790
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_DirRecojo 
         Height          =   285
         Left            =   7230
         TabIndex        =   4
         Tag             =   "TidDirRecojo"
         Top             =   660
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
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
         Container       =   "frmMantDirRecojo.frx":5A7E
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Marca 
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Tag             =   "TglsDirRecojo"
         Top             =   3180
         Width           =   6165
         _ExtentX        =   10874
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
         MaxLength       =   200
         Container       =   "frmMantDirRecojo.frx":5A9A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Pais 
         Height          =   315
         Left            =   1650
         TabIndex        =   15
         Tag             =   "TidPais"
         Top             =   1650
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
         Container       =   "frmMantDirRecojo.frx":5AB6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Pais 
         Height          =   315
         Left            =   2610
         TabIndex        =   16
         Top             =   1650
         Width           =   5205
         _ExtentX        =   9181
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
         Container       =   "frmMantDirRecojo.frx":5AD2
      End
      Begin CATControls.CATTextBox txtCod_Depa 
         Height          =   315
         Left            =   1650
         TabIndex        =   17
         Top             =   2025
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
         Container       =   "frmMantDirRecojo.frx":5AEE
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Depa 
         Height          =   315
         Left            =   2610
         TabIndex        =   18
         Top             =   2010
         Width           =   5205
         _ExtentX        =   9181
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
         Container       =   "frmMantDirRecojo.frx":5B0A
      End
      Begin CATControls.CATTextBox txtCod_Prov 
         Height          =   315
         Left            =   1650
         TabIndex        =   19
         Top             =   2400
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
         Container       =   "frmMantDirRecojo.frx":5B26
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Prov 
         Height          =   315
         Left            =   2610
         TabIndex        =   20
         Top             =   2400
         Width           =   5205
         _ExtentX        =   9181
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
         Container       =   "frmMantDirRecojo.frx":5B42
      End
      Begin CATControls.CATTextBox txtCod_Distrito 
         Height          =   315
         Left            =   1650
         TabIndex        =   21
         Tag             =   "TidDistrito"
         Top             =   2775
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
         Container       =   "frmMantDirRecojo.frx":5B5E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Distrito 
         Height          =   315
         Left            =   2610
         TabIndex        =   22
         Top             =   2775
         Width           =   5205
         _ExtentX        =   9181
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
         Container       =   "frmMantDirRecojo.frx":5B7A
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   270
         TabIndex        =   26
         Top             =   2370
         Width           =   660
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   270
         TabIndex        =   25
         Top             =   1995
         Width           =   1005
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "País"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   270
         TabIndex        =   24
         Top             =   1620
         Width           =   300
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   270
         TabIndex        =   23
         Top             =   2820
         Width           =   495
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6615
         TabIndex        =   6
         Top             =   690
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   270
         TabIndex        =   5
         Top             =   3240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMantDirRecojo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaDepa_Click()
    mostrarAyuda "DEPARTAMENTO", txtCod_Depa, txtGls_Depa, " AND idPais = '" & txtCod_Pais.Text & "'"
    If txtCod_Depa.Text <> "" Then SendKeys "{tab}"
End Sub

Private Sub cmbAyudaDistrito_Click()
    mostrarAyuda "DISTRITO", txtCod_Distrito, txtGls_Distrito, "AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Depa.Text & "' and idProv = '" + txtCod_Prov.Text + "'"
    If txtCod_Distrito.Text <> "" Then SendKeys "{tab}"
End Sub

Private Sub cmbAyudaPais_Click()
    mostrarAyuda "PAIS", txtCod_Pais, txtGls_Pais
    If txtCod_Pais.Text <> "" Then SendKeys "{tab}"
End Sub

Private Sub cmbAyudaProv_Click()
    mostrarAyuda "PROVINCIA", txtCod_Prov, txtGls_Prov, "AND idPais = '" & txtCod_Pais.Text & "' AND idDpto = '" & txtCod_Depa.Text + "'"
    If txtCod_Prov.Text <> "" Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLista, False, False, False, False
    
    listaMarca StrMsgError
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
Dim strCodigo   As String
Dim strMsg      As String, CSqlC As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    validaHomonimia "DirRecojos", "glsDirRecojo", "idDirRecojo", txtGls_Marca.Text, txtCod_DirRecojo.Text, True, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If txtCod_DirRecojo.Text = "" Then
        txtCod_DirRecojo.Text = GeneraCorrelativoAnoMes("DirRecojos", "idDirRecojo")
        EjecutaSQLForm Me, 0, True, "DirRecojos", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabó"
        
    Else
        EjecutaSQLForm Me, 1, True, "DirRecojos", StrMsgError, "idDirRecojo"
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Modificó"
    End If
    
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title

    fraGeneral.Enabled = False
    
    listaMarca StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub nuevo()
    
    limpiaForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    With gLista
        .Dataset.Close
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = ""
        .Dataset.ADODataset.CommandText = ""
        .Dataset.Active = False
    End With

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarMarca gLista.Columns.ColumnByName("idDirRecojo").Value, StrMsgError
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
        Case 4, 7  'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
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
On Error GoTo Err
Dim StrMsgError As String

    listaMarca StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub listaMarca(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND m.glsDirRecojo LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT m.idDirRecojo ,m.glsDirRecojo " & _
           "FROM DirRecojos m WHERE m.idEmpresa = '" & glsEmpresa & "'"
    If strCond <> "" Then csql = csql + strCond
    csql = csql & " ORDER BY m.idDirRecojo"
    
    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "idDirRecojo"
    End With
    Me.Refresh
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub mostrarMarca(strCodMarca As String, ByRef StrMsgError As String)
Dim StrDire     As String
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT m.idDirRecojo,m.glsDirRecojo,m.idPais,m.idDistrito " & _
           "FROM dirrecojos m " & _
           "WHERE m.idDirRecojo = '" & strCodMarca & "' AND m.idEmpresa = '" & glsEmpresa & "'"
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    StrDire = Trim("" & txtCod_Distrito.Text)
    
    If Len(Trim("" & StrDire)) > 0 Then
        
        txtCod_Depa.Text = left(Trim("" & StrDire), 2)
        txtCod_Prov.Text = Mid(Trim("" & StrDire), 3, 2)
        txtCod_Distrito.Text = StrDire
    End If
    
    
    Me.Refresh
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("Está seguro(a) de eliminar el registro?" & vbCrLf & "Se eliminaran todas sus dependencias.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    strCodigo = Trim(txtCod_Marca.Text)
    
    csql = "SELECT idMarca FROM productos WHERE idMarca = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
    rsValida.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsValida.EOF Then
        StrMsgError = "No se puede eliminar el registro, el registro se encuentra en uso (Productos)."
        GoTo Err
    End If
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando el registro
    csql = "DELETE FROM marcas WHERE idMarca = '" & strCodigo & "' AND idEmpresa = '" & glsEmpresa & "'"
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

Private Sub txtCod_Depa_Change()
    txtGls_Depa.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDpto", txtCod_Depa.Text, False, " idProv = '00' And idPais = '" & txtCod_Pais.Text & "' ")
    txtCod_Prov.Text = ""
    txtGls_Prov.Text = ""
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""
End Sub

Private Sub txtCod_Depa_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DEPARTAMENTO", txtCod_Depa, txtGls_Depa
        KeyAscii = 0
        If txtCod_Depa.Text <> "" Then SendKeys "{tab}"
    End If
End Sub

Private Sub txtCod_Distrito_Change()
    txtGls_Distrito.Text = traerCampo("Ubigeo", "GlsUbigeo", "idDistrito", txtCod_Distrito.Text, False, "idPais = '" & txtCod_Pais.Text & "'")
    If indCopiaDireccion Then
        indEditando = True
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & "/" & txtGls_Distrito.Text & "/" & txtGls_Prov.Text & "/" & txtGls_Depa.Text
        indEditando = False
    End If
End Sub

Private Sub txtCod_Distrito_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DISTRITO", txtCod_Distrito, txtGls_Distrito, "AND idDpto = '" & txtCod_Depa.Text & "' and idProv = '" + txtCod_Prov.Text + "'"
        KeyAscii = 0
        If txtCod_Distrito.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtCod_Pais_Change()
On Error GoTo Err
Dim StrMsgError                 As String

    txtGls_Pais.Text = traerCampo("Datos", "GlsDato", "idDato", txtCod_Pais.Text, False)
    txtCod_Depa.Text = ""
    txtGls_Depa.Text = ""
    txtCod_Prov.Text = ""
    txtGls_Prov.Text = ""
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""
    
    If txtGls_Pais.Text = "" Or traerCampo("Datos", "IdSunat", "IdDato", txtCod_Pais.Text, False) = "9589" Then
        txtCod_Depa.Vacio = False
        txtGls_Depa.Vacio = False
        txtCod_Prov.Vacio = False
        txtGls_Prov.Vacio = False
        txtCod_Distrito.Vacio = False
        txtGls_Distrito.Vacio = False
    Else
        txtCod_Depa.Vacio = True
        txtGls_Depa.Vacio = True
        txtCod_Prov.Vacio = True
        txtGls_Prov.Vacio = True
        txtCod_Distrito.Vacio = True
        txtGls_Distrito.Vacio = True
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Pais_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PAIS", txtCod_Pais, txtGls_Pais
        KeyAscii = 0
        If txtCod_Pais.Text <> "" Then SendKeys "{tab}"
    End If
End Sub


Private Sub txtCod_Prov_Change()
    txtGls_Prov.Text = traerCampo("Ubigeo", "GlsUbigeo", "idProv", txtCod_Prov.Text, False, " idDpto = '" & txtCod_Depa.Text & "' and idProv <> '00' and idDist = '00' And idPais = '" & txtCod_Pais.Text & "' ")
    txtCod_Distrito.Text = ""
    txtGls_Distrito.Text = ""
    If indCopiaDireccion Then
        indEditando = True
        txtGls_DireccionFiscal.Text = txtGls_Direccion.Text & "/" & txtGls_Distrito.Text & "/" & txtGls_Prov.Text & "/" & txtGls_Depa.Text
        indEditando = False
    End If
End Sub

Private Sub txtCod_Prov_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "PROVINCIA", txtCod_Prov, txtGls_Prov, "AND idDpto = '" & txtCod_Depa.Text + "'"
        KeyAscii = 0
        If txtCod_Prov.Text <> "" Then SendKeys "{tab}"
    End If
End Sub


