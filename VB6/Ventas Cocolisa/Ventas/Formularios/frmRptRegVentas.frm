VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmRptRegVentas 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Ventas"
   ClientHeight    =   9165
   ClientLeft      =   1710
   ClientTop       =   1545
   ClientWidth     =   14175
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   1680
      Top             =   5490
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
            Picture         =   "frmRptRegVentas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":317E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":3518
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":396A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":3D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRptRegVentas.frx":4716
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
      ForeColor       =   &H00C00000&
      Height          =   8475
      Left            =   60
      TabIndex        =   8
      Top             =   630
      Width           =   14040
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6810
         Left            =   135
         TabIndex        =   23
         Top             =   1530
         Width           =   13785
         Begin VB.Frame FraObs 
            Appearance      =   0  'Flat
            Caption         =   " Observación "
            ForeColor       =   &H80000008&
            Height          =   2445
            Left            =   3915
            TabIndex        =   25
            Top             =   1170
            Visible         =   0   'False
            Width           =   5685
            Begin VB.TextBox TxtGls_Observacion 
               Appearance      =   0  'Flat
               Height          =   1365
               Left            =   180
               MultiLine       =   -1  'True
               TabIndex        =   27
               Top             =   630
               Width           =   5325
            End
            Begin VB.CommandButton CmdAceptar 
               Appearance      =   0  'Flat
               Caption         =   "Aceptar"
               Height          =   330
               Left            =   1935
               TabIndex        =   26
               Top             =   2025
               Width           =   1680
            End
            Begin VB.Label lblDocumento 
               Alignment       =   2  'Center
               Caption         =   "Documento"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   180
               TabIndex        =   28
               Top             =   270
               Width           =   5325
            End
         End
         Begin DXDBGRIDLibCtl.dxDBGrid g 
            Height          =   6495
            Left            =   90
            OleObjectBlob   =   "frmRptRegVentas.frx":4DE8
            TabIndex        =   24
            Top             =   180
            Width           =   13605
         End
      End
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
         Height          =   1380
         Left            =   135
         TabIndex        =   9
         Top             =   135
         Width           =   13785
         Begin VB.CheckBox ChkA4 
            Caption         =   "A4"
            Height          =   330
            Left            =   6615
            TabIndex        =   29
            Top             =   990
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton CmdAyudaUnidProduc 
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
            Left            =   5770
            Picture         =   "frmRptRegVentas.frx":B0D6
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   960
            Width           =   390
         End
         Begin VB.CheckBox ChkOficial 
            Caption         =   "Documentos"
            Height          =   330
            Left            =   12240
            TabIndex        =   6
            Top             =   585
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CheckBox ChkOrden 
            Caption         =   "Agrupado por tipo de Documento"
            Height          =   375
            Left            =   10755
            TabIndex        =   3
            Top             =   180
            Width           =   2715
         End
         Begin VB.CommandButton cmbAyudaTipoDoc 
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
            Left            =   5770
            Picture         =   "frmRptRegVentas.frx":B460
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaSucursal 
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
            Left            =   5770
            Picture         =   "frmRptRegVentas.frx":B7EA
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   615
            Width           =   390
         End
         Begin CATControls.CATTextBox txt_Serie 
            Height          =   315
            Left            =   7245
            TabIndex        =   2
            Top             =   255
            Width           =   690
            _ExtentX        =   1217
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
            MaxLength       =   4
            Container       =   "frmRptRegVentas.frx":BB74
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Sucursal 
            Height          =   315
            Left            =   1080
            TabIndex        =   1
            Tag             =   "TidMoneda"
            Top             =   615
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
            Container       =   "frmRptRegVentas.frx":BB90
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Sucursal 
            Height          =   315
            Left            =   2010
            TabIndex        =   11
            Top             =   615
            Width           =   3735
            _ExtentX        =   6588
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
            Container       =   "frmRptRegVentas.frx":BBAC
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_TipoDoc 
            Height          =   315
            Left            =   1080
            TabIndex        =   0
            Tag             =   "TidMoneda"
            Top             =   255
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
            Container       =   "frmRptRegVentas.frx":BBC8
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_TipoDoc 
            Height          =   315
            Left            =   2010
            TabIndex        =   14
            Top             =   255
            Width           =   3735
            _ExtentX        =   6588
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
            Container       =   "frmRptRegVentas.frx":BBE4
            Vacio           =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   7245
            TabIndex        =   4
            Top             =   615
            Width           =   1245
            _ExtentX        =   2196
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
            Format          =   142082049
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   9225
            TabIndex        =   5
            Top             =   615
            Width           =   1245
            _ExtentX        =   2196
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
            Format          =   142082049
            CurrentDate     =   38667
         End
         Begin CATControls.CATTextBox txtCod_UnidProd 
            Height          =   285
            Left            =   1080
            TabIndex        =   20
            Tag             =   "Tidupp"
            Top             =   975
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
            Container       =   "frmRptRegVentas.frx":BC00
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_UnidProd 
            Height          =   285
            Left            =   2010
            TabIndex        =   21
            Top             =   975
            Width           =   3735
            _ExtentX        =   6588
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
            Container       =   "frmRptRegVentas.frx":BC1C
            Vacio           =   -1  'True
         End
         Begin VB.Label lbl_upp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Área"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   210
            TabIndex        =   22
            Top             =   990
            Width           =   360
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6600
            TabIndex        =   18
            Top             =   660
            Width           =   465
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   8655
            TabIndex        =   17
            Top             =   660
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Serie"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   6600
            TabIndex        =   16
            Top             =   270
            Width           =   375
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Tipo Doc."
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   210
            TabIndex        =   15
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Sucursal"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   210
            TabIndex        =   12
            Top             =   660
            Width           =   645
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   2170
      ButtonWidth     =   2514
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actualizar"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reg.Ventas"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consolidado"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anexo de Ventas"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Por Sucursal"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "R.V. Formato 2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "R.V Formato 3"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "R.V Formato 4"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmRptRegVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaSucursal_Click()
On Error GoTo Err
Dim StrMsgError As String

    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
    If txtCod_Sucursal.Text <> "" Then SendKeys "{tab}"
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmbAyudaTipoDoc_Click()
On Error GoTo Err
Dim StrMsgError     As String

    mostrarAyuda "DOCUMENTOS", txtCod_TipoDoc, txtGls_TipoDoc, IIf(ChkOficial.Value = 1, "And IndOficial='1'", "")
    If txtCod_TipoDoc.Text <> "" Then SendKeys "{tab}"
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError     As String
Dim Cadmysql        As String
    
    FraObs.Visible = False
        
    Cadmysql = "Update Docventas Set ObsRegVentas = '" & txtGls_Observacion.Text & "' " & _
               "WHERE iddocventas = '" & G.Columns.ColumnByFieldName("idDocventas").Value & "' " & _
               "AND idserie = '" & G.Columns.ColumnByFieldName("idserie").Value & "' " & _
               "AND iddocumento = '" & G.Columns.ColumnByFieldName("idDocumento").Value & "' " & _
               "AND idempresa = '" & glsEmpresa & "' "
    Cn.Execute (Cadmysql)
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaUnidProduc_Click()
On Error GoTo Err
Dim StrMsgError     As String

    mostrarAyuda "UNIDADPRODUC", txtCod_UnidProd, txtGls_UnidProd
    If txtCod_UnidProd.Text <> "" Then SendKeys "{tab}"
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()

    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = False
    
    If traerCampo("Parametros", "Valparametro", "glsparametro", "IMPRESION_A4", True) = 1 Then
        
        ChkA4.Value = 1
        
    Else
        
        ChkA4.Value = 0
        
    End If
        
    ChkOrden.Value = 1
    ConfGrid G, False, False, False, False
    
    txtGls_TipoDoc.Text = "TODOS LOS DOCUMENTOS"
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    
    If leeParametro("DESCRIPCION_AREA_O_UPP") = "1" Then
        lbl_upp.Caption = "Área"
        txtGls_UnidProd.Text = "TODAS LAS AREAS"
    Else
        lbl_upp.Caption = "UPP"
        txtGls_UnidProd.Text = "TODAS LAS UNIDADES DE PRODUCCION"
    End If
    
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    Me.top = 0
    Me.left = 0
    ChkOficial.Visible = IIf(GlsVisualiza_Filtro_Documento = "S", True, False)
    
    If Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "VISUALIZA_OBS_REGVENTAS", True)) = "S" Then
        G.Columns.ColumnByFieldName("ObsRegVentas").Visible = True
        Toolbar1.Buttons(7).Visible = True
        Toolbar1.Buttons(8).Visible = True
        Toolbar1.Buttons(9).Visible = True
    Else
        G.Columns.ColumnByFieldName("ObsRegVentas").Visible = False
        Toolbar1.Buttons(7).Visible = False
        Toolbar1.Buttons(8).Visible = False
        Toolbar1.Buttons(9).Visible = False
    End If

End Sub

Private Sub g_OnClick()
On Error GoTo Err
Dim StrMsgError As String

    Select Case G.Columns.FocusedColumn.Index
        Case G.Columns.ColumnByFieldName("ObsRegVentas").Index
            FraObs.Visible = True
            txtGls_Observacion.Text = Trim("" & traerCampo("Docventas", "ObsRegVentas", "iddocventas", G.Columns.ColumnByFieldName("iddocventas").Value, True, " idserie = '" & G.Columns.ColumnByFieldName("idSerie").Value & "' and iddocumento = '" & G.Columns.ColumnByFieldName("idDocumento").Value & "' "))
            lbldocumento.Caption = Trim("" & traerCampo("Documentos", "GlsDocumento", "Iddocumento", G.Columns.ColumnByFieldName("idDocumento").Value, False)) & "  " & G.Columns.ColumnByFieldName("idSerie").Value & "-" & G.Columns.ColumnByFieldName("iddocventas").Value
            txtGls_Observacion.SetFocus
    End Select

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim rsReporte           As ADODB.Recordset
Dim StrMsgError         As String
Dim strSucursal         As String
Dim strTipoDoc          As String
Dim strSerie            As String
Dim strFecIni           As String
Dim strFecFin           As String
Dim strformato2         As String
    
    strTipoDoc = Trim(txtCod_TipoDoc.Text)
    strSerie = Trim(txt_serie.Text)
    strSucursal = Trim(txtCod_Sucursal.Text)
    
    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
        
    Select Case Button.Index
        Case 1 '--- Actualizar
            listaRegVentas StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 2 '--- Reg. Ventas
            'If traerCampo("Parametros", "Valparametro", "glsparametro", "IMPRESION_A4", True) = 1 Then
            If ChkA4.Value = 1 Then
                If ChkOrden.Value = 1 Then
                    If traerCampo("Parametros", "Valparametro", "glsparametro", "REG_VENTAS_A4_FORMATO2", True) = 1 Then
                        mostrarReporte "rptRegVentasA4Formato2.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    
                    ElseIf traerCampo("Parametros", "Valparametro", "glsparametro", "REG_VENTAS_A4_FORMATO2", True) = 2 Then 'Formato Oficial para Inmac
                        mostrarReporte "rptRegistroVentasOficialA4Formato2.rpt", "parEmpresa|parSucursal|parTipoDcto|parSerie|parMoneda|ParFechaIni|ParFechaFin|parOficial", glsEmpresa & "|" & strSucursal & "|" & strTipoDoc & "|" & strSerie & "|" & xmoneda & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1"), "Registro de Ventas", StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    
                    ElseIf traerCampo("Parametros", "Valparametro", "glsparametro", "REG_VENTAS_A4_FORMATO2", True) = 3 Then 'Formato Oficial para Rally
                        mostrarReporte "rptRegVentasRallyA4.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
                    
                    ElseIf traerCampo("Parametros", "Valparametro", "glsparametro", "REG_VENTAS_A4_FORMATO2", True) = 4 Then 'Formato Despensa
                        mostrarReporte "rptRegVentasIVAPA4.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                        
                        
                    Else
                        mostrarReporte "rptRegVentasA4.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
                        If StrMsgError <> "" Then GoTo Err
                    End If
                Else
                    mostrarReporte "rptRegVentas_SingrupoA4.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                End If
            Else
                If ChkOrden.Value = 1 Then
                    mostrarReporte "rptRegVentas.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                    
                Else
                    mostrarReporte "rptRegVentas_Singrupo.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
                    If StrMsgError <> "" Then GoTo Err
                End If
            End If
            
        Case 3 '--- Consolidado
            mostrarReporte "rptRegVentasConsolidado.rpt", "varEmpresa|varSucursal|varTipoDoc|varSerie|varMoneda|varFechaIni|varFechaFin|varOficial|varIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "0") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
            If StrMsgError <> "" Then GoTo Err
    
        Case 4 '--- Anexo De Ventas
            mostrarReporte "rptanexoVentas.rpt", "varEmpresa|varSucursal|varTipoDoc|varSerie|varMoneda|varFechaIni|varFechaFin|varOficial|varIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "0") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 5 '--- Por Sucursal
            mostrarReporte "rptRegVentasSucursales.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
            If StrMsgError <> "" Then GoTo Err
        
        Case 6 '--- Reg. Ventas Formato 2
            strformato2 = Val(Trim("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "FORMATO_FORMATO2_VENTAS", True)))
            
            If strformato2 = 0 Then
                mostrarReporte "rptRegVentasA4Formato2.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
                If StrMsgError <> "" Then GoTo Err
            Else
               mostrarReporte "rptRegVentasFormato4_EmpRelacionada.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
               If StrMsgError <> "" Then GoTo Err

               mostrarReporte "rptRegVentasFormato4_Area.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
               If StrMsgError <> "" Then GoTo Err

               mostrarReporte "rptRegVentasFormato4.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
               If StrMsgError <> "" Then GoTo Err
            End If
        
        Case 7 '--- Reg. Ventas Formato 3
               ExportarExcel StrMsgError
               If StrMsgError <> "" Then GoTo Err
        Case 8 '--- Reg. Ventas Formato 4
               mostrarReporte "rptRegVentasFormato3.rpt", "parEmpresa|parSucursal|parTipoDoc|parSerie|parMoneda|parFechaIni|ParFechaFin|ParOficial|ParIdArea", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & strTipoDoc & "|" & strSerie & "|" & "PEN" & "|" & strFecIni & "|" & strFecFin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & txtCod_UnidProd.Text, "Registro de Ventas", StrMsgError
               If StrMsgError <> "" Then GoTo Err
        Case 9 '--- Excel
            G.m.ExportToXLS App.Path & "\Temporales\RegVentas.xls"
            ShellEx App.Path & "\Temporales\RegVentas.xls", essSW_MAXIMIZE, , , "open", Me.hwnd
        
        Case 10 '--- Salir
            Unload Me
    End Select
    
    Exit Sub
    
Err:
    If TypeName(rsReporte) = "Recordset" Then
        If rsReporte.State = 1 Then rsReporte.Close: Set rsReporte = Nothing
    End If
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Sucursal_Change()
On Error GoTo Err
Dim StrMsgError     As String

    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Sucursal_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError     As String

    If KeyCode = 8 Then
        txtCod_Sucursal.Text = ""
    ElseIf KeyCode = 113 Then
        mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        If txtCod_Sucursal.Text <> "" Then SendKeys "{tab}"
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_TipoDoc_Change()
On Error GoTo Err
Dim StrMsgError     As String

    If txtCod_TipoDoc.Text <> "" Then
        txtGls_TipoDoc.Text = traerCampo("documentos", "GlsDocumento", "idDocumento", txtCod_TipoDoc.Text, False)
    Else
        txtGls_TipoDoc.Text = "TODOS LOS DOCUMENTOS"
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_TipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError     As String

    If KeyCode = 8 Then
        txtCod_TipoDoc.Text = ""
    ElseIf KeyCode = 113 Then
        mostrarAyuda "DOCUMENTOS", txtCod_TipoDoc, txtGls_TipoDoc, IIf(ChkOficial.Value = 1, "And IndOficial='1'", "")
        If txtCod_TipoDoc.Text <> "" Then SendKeys "{tab}"
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_TipoDoc_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "DOCUMENTOS", txtCod_TipoDoc, txtGls_TipoDoc
        KeyAscii = 0
    End If

End Sub

Private Sub listaRegVentas(ByRef StrMsgError As String)
On Error GoTo Err
Dim strSucursal     As String
Dim strTipoDoc      As String
Dim strSerie        As String
Dim strFecIni       As String
Dim strFecFin       As String
Dim rsdatos         As New ADODB.Recordset

    strTipoDoc = Trim(txtCod_TipoDoc.Text)
    strSerie = Trim(txt_serie.Text)
    strSucursal = Trim(txtCod_Sucursal.Text)
    
    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    csql = "exec Spu_ListaRegVentas '" & glsEmpresa & "','" & strSucursal & "','" & strTipoDoc & "','" & strSerie & "','PEN','" & strFecIni & "','" & strFecFin & "','" & IIf(ChkOficial.Visible, ChkOficial.Value, "0") & "','" & txtCod_UnidProd.Text & "'"
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
    Set G.DataSource = rsdatos

'    With G
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = "Call Spu_ListaRegVentas('" & glsEmpresa & "','" & strSucursal & "','" & strTipoDoc & "','" & strSerie & "','PEN','" & strFecIni & "','" & strFecFin & "','" & IIf(ChkOficial.Visible, ChkOficial.Value, "0") & "','" & txtCod_UnidProd.Text & "')"
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "item"
'    End With

    Me.Refresh
    
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarReporteregv(ByVal rsReporte As ADODB.Recordset, ByVal GlsReporte As String, ByVal GlsTitulo As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim gStrRutaRpts    As String

    Screen.MousePointer = 11
    gStrRutaRpts = App.Path + "\Reportes\"
    
    Set reporte = aplicacion.OpenReport(gStrRutaRpts & GlsReporte)
                    
    If Not rsReporte.EOF And Not rsReporte.BOF Then
        reporte.Database.SetDataSource rsReporte, 3
        vistaPrevia.CRViewer91.ReportSource = reporte
        vistaPrevia.Caption = GlsTitulo
        vistaPrevia.CRViewer91.ViewReport
        vistaPrevia.CRViewer91.DisplayGroupTree = False
        Screen.MousePointer = 0
        vistaPrevia.WindowState = 2
        vistaPrevia.Show
    Else
        Screen.MousePointer = 0
        MsgBox "No existen Registros  Seleccionados", vbInformation, App.Title
    End If
    Screen.MousePointer = 0
        
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
    
    Exit Sub

Err:
    Screen.MousePointer = 0
    If rsReporte.State = 1 Then rsReporte.Close
    Set rsReporte = Nothing
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_UnidProd_Change()
On Error GoTo Err
Dim StrMsgError     As String

    If txtCod_UnidProd.Text = "" Then
        If leeParametro("DESCRIPCION_AREA_O_UPP") = "1" Then
            txtGls_UnidProd.Text = "TODAS LAS AREAS"
        Else
            txtGls_UnidProd.Text = "TODAS LAS UNIDADES DE PRODUCCION"
        End If
    Else
        txtGls_UnidProd.Text = traerCampo("UnidadProduccion", "DescUnidad", "CodUnidProd", txtCod_UnidProd.Text, True)
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_UnidProd_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
Dim StrMsgError     As String

    If KeyCode = 113 Then
        mostrarAyuda "UNIDADPRODUC", txtCod_UnidProd, txtGls_UnidProd
        If txtCod_UnidProd.Text <> "" Then SendKeys "{tab}"
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub ExportarExcel(ByRef StrMsgError As String)
On Error GoTo Err
Dim xl                  As Excel.Application
Dim wb                  As Workbook
Dim shRxn               As Worksheet
Dim rng                 As Range
Dim carchivo            As String
Dim cletra              As String
Dim crango              As String
Dim nfila               As Integer
Dim ncol                As Integer
Dim ncont               As Integer
Dim ncont_und           As Integer
Dim ntotventa           As Double
Dim ntotcant            As Double
Dim StrGlsEmpresa       As String
Dim StrGlsRuc           As String
Dim xidnivel            As String
Dim xcolum              As Integer
Dim xfila               As Integer
Dim i                   As Integer
Dim Existe              As Boolean
Dim RsCostoMensual      As New ADODB.Recordset
Dim IntMes              As String
Dim intAnnio            As String
Dim RsNiveles           As New ADODB.Recordset
Dim Cadmysql            As String
Dim ContHoja            As Integer
Dim GlsHoja             As String
Dim o                   As Integer
Dim rsReporte           As New ADODB.Recordset
Dim strSucursal         As String
Dim strTipoDoc          As String
Dim strSerie            As String
Dim strFecIni           As String
Dim strFecFin           As String
Dim strformato2         As String
Dim strFecEmi           As String

    strTipoDoc = Trim(txtCod_TipoDoc.Text)
    strSerie = Trim(txt_serie.Text)
    strSucursal = Trim(txtCod_Sucursal.Text)
    
    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")

    Me.MousePointer = 11
    carchivo = gStrRutaRpts & "VentasporFecha.xlt"
    
    On Error GoTo ExcelNoAbierto
    Set xl = GetObject(, "Excel.Application")
    GoTo YaEstabaAbierto
ExcelNoAbierto:
    Set xl = New Excel.Application
YaEstabaAbierto:
    On Error GoTo 0
    Set wb = xl.Workbooks.Open(carchivo)
    
    xl.Visible = False
    
    Existe = False
    ContHoja = 0
    GlsHoja = ""
    
    'Set rsReporte = DataProcedimiento("spu_ListaRegVentas", StrMsgError, glsEmpresa, strSucursal, strTipoDoc, strSerie, "PEN", strFecIni, strFecFin, IIf(ChkOficial.Visible, ChkOficial.Value, "1"), txtCod_UnidProd.Text)
    'If StrMsgError <> "" Then GoTo ERR
    
    csql = "Call spu_ListaRegVentas ('" & glsEmpresa & "', '" & strSucursal & "', '" & strTipoDoc & "','" & strSerie & "' , 'PEN','" & strFecIni & "' , '" & strFecFin & "','" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "', '" & txtCod_UnidProd.Text & "')"
    rsReporte.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    nfila = 8
    ncol = 1
    If Not rsReporte.EOF Then
        xl.Sheets(1).Select
        xl.Cells(1, 1).Value = Trim("" & traerCampo("empresas", "glsEmpresa", "idEmpresa", glsEmpresa, True))
        xl.Cells(2, 1).Value = Trim("" & traerCampo("empresas", "ruc", "idEmpresa", glsEmpresa, True))
        xl.Cells(5, 9).Value = "Del " & Format(dtpfInicio.Value, "yyyy-mm-dd") & " Al " & Format(dtpFFinal.Value, "yyyy-mm-dd")
        
        Do While Not rsReporte.EOF
        
            'xl.Cells(nfila, 1).Value = Format(Trim("" & rsReporte.Fields("FecEmision")), "dd/mm/yyyy")
            xl.Cells(nfila, 1).Select
            strFecEmi = Format(Trim("" & rsReporte.Fields("FecEmision")), "yyyy-mm-dd")
            xl.Selection.NumberFormat = "dd/mm/yyyy;@"
            With xl.Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlLTR
                .MergeCells = False
            End With
            xl.Cells(nfila, 1).Value = strFecEmi
            xl.Cells(nfila, 2).Value = Trim("" & rsReporte.Fields("DescUnidad"))
            xl.Cells(nfila, 3).Value = rsReporte.Fields("idDocumento")
            xl.Cells(nfila, 4).Value = rsReporte.Fields("idSerie")
            xl.Cells(nfila, 5).Value = rsReporte.Fields("iddocventas")
            xl.Cells(nfila, 6).Value = rsReporte.Fields("RucCliente")
            xl.Cells(nfila, 7).Value = rsReporte.Fields("GlsCliente")
            xl.Cells(nfila, 8).Value = rsReporte.Fields("glsUpCliente")
            xl.Cells(nfila, 9).Value = rsReporte.Fields("GlsFormaspago")
            xl.Cells(nfila, 10).Value = Format(rsReporte.Fields("TipoCambio"), "0.000")
            xl.Cells(nfila, 11).Value = Format(rsReporte.Fields("BaseImponible"), "0.00")
            xl.Cells(nfila, 12).Value = Format(rsReporte.Fields("dscto"), "0.00")
            xl.Cells(nfila, 13).Value = Format(rsReporte.Fields("Exonerado"), "0.00")
            xl.Cells(nfila, 14).Value = Format(rsReporte.Fields("TotalIgvVenta"), "0.00")
            xl.Cells(nfila, 15).Value = Format(rsReporte.Fields("totalPrecioVenta"), "0.00")
            xl.Cells(nfila, 17).Value = Trim("" & rsReporte.Fields("idCentroCosto"))
            xl.Cells(nfila, 17).Select
            xl.Columns("Q:Q").EntireColumn.AutoFit
            xl.Selection.NumberFormat = "@"
            With xl.Selection
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlLTR
                .MergeCells = False
            End With
             
            nfila = nfila + 1
            rsReporte.MoveNext
        Loop
    End If
    
    xl.Cells(nfila, 11).Formula = "=SUM(" & "K" & 8 & ":" & "k" & nfila - 1 & ")"
    xl.Cells(nfila, 14).Formula = "=SUM(" & "N" & 8 & ":" & "N" & nfila - 1 & ")"
    xl.Cells(nfila, 15).Formula = "=SUM(" & "O" & 8 & ":" & "O" & nfila - 1 & ")"
    xl.Cells(nfila, 11).Select
    xl.Selection.Font.Bold = True
    xl.Cells(nfila, 14).Select
    xl.Selection.Font.Bold = True
    xl.Cells(nfila, 15).Select
    xl.Selection.Font.Bold = True
                
    'Set rsReporte = DataProcedimiento("spu_ListaRegVentas_Area", StrMsgError, glsEmpresa, strSucursal, strTipoDoc, strSerie, "PEN", strFecIni, strFecFin, IIf(ChkOficial.Visible, ChkOficial.Value, "1"), txtCod_UnidProd.Text)
    'If StrMsgError <> "" Then GoTo ERR
    
    If rsReporte.State = 1 Then rsReporte.Close: Set rsReporte = Nothing
    
    csql = "Call spu_ListaRegVentas_Area ('" & glsEmpresa & "', '" & strSucursal & "', '" & strTipoDoc & "','" & strSerie & "' , 'PEN','" & strFecIni & "' , '" & strFecFin & "','" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "', '" & txtCod_UnidProd.Text & "')"
    rsReporte.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    nfila = 8
    ncol = 1
    If Not rsReporte.EOF Then
        xl.Sheets(2).Select
        xl.Cells(1, 1).Value = Trim("" & traerCampo("empresas", "glsEmpresa", "idEmpresa", glsEmpresa, True))
        xl.Cells(2, 1).Value = Trim("" & traerCampo("empresas", "ruc", "idEmpresa", glsEmpresa, True))
        xl.Cells(5, 6).Value = "Del " & Format(dtpfInicio.Value, "yyyy-mm-dd") & " Al " & Format(dtpFFinal.Value, "yyyy-mm-dd")
        Do While Not rsReporte.EOF
            xl.Cells(nfila, 4).Value = rsReporte.Fields("DescUnidad")
            xl.Cells(nfila, 5).Value = Format(rsReporte.Fields("BaseImponible"), "0.00")
            xl.Cells(nfila, 6).Value = Format(rsReporte.Fields("totalPrecioVenta"), "0.00")
            nfila = nfila + 1
            rsReporte.MoveNext
        Loop
    End If
        
    xl.Cells(nfila, 4).Value = "TOTAL "
    xl.Cells(nfila, 5).Formula = "=SUM(" & "E" & 8 & ":" & "E" & nfila - 1 & ")"
    xl.Cells(nfila, 6).Formula = "=SUM(" & "F" & 8 & ":" & "F" & nfila - 1 & ")"
    xl.Cells(nfila, 4).Select
    xl.Selection.Font.Bold = True
    xl.Cells(nfila, 5).Select
    xl.Selection.Font.Bold = True
    xl.Cells(nfila, 6).Select
    xl.Selection.Font.Bold = True
    
    For i = 8 To nfila - 1
        xl.Cells(i, 7).Formula = Format("=SUM(" & "(F" & i & "/" & "F" & nfila & ")* 100)", "0.00")
    Next i
    
    xl.Cells(nfila, 7).Formula = "=SUM(" & "G" & 8 & ":" & "G" & nfila - 1 & ")"
    xl.Cells(nfila, 7).Select
    xl.Selection.Font.Bold = True
    
    'Set rsReporte = DataProcedimiento("spu_ListaRegVentas_Relacionada", StrMsgError, glsEmpresa, strSucursal, strTipoDoc, strSerie, "PEN", strFecIni, strFecFin, IIf(ChkOficial.Visible, ChkOficial.Value, "1"), txtCod_UnidProd.Text)
    'If StrMsgError <> "" Then GoTo ERR
    
    If rsReporte.State = 1 Then rsReporte.Close: Set rsReporte = Nothing
    
    csql = "Call spu_ListaRegVentas_Relacionada ('" & glsEmpresa & "', '" & strSucursal & "', '" & strTipoDoc & "','" & strSerie & "' , 'PEN','" & strFecIni & "' , '" & strFecFin & "','" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "', '" & txtCod_UnidProd.Text & "')"
    rsReporte.Open csql, Cn, adOpenStatic, adLockReadOnly
    
    nfila = 8
    ncol = 1
    If Not rsReporte.EOF Then
        xl.Sheets(3).Select
        xl.Cells(1, 1).Value = Trim("" & traerCampo("empresas", "glsEmpresa", "idEmpresa", glsEmpresa, True))
        xl.Cells(2, 1).Value = Trim("" & traerCampo("empresas", "ruc", "idEmpresa", glsEmpresa, True))
        xl.Cells(5, 6).Value = "Del " & Format(dtpfInicio.Value, "yyyy-mm-dd") & " Al " & Format(dtpFFinal.Value, "yyyy-mm-dd")
        
        Do While Not rsReporte.EOF
            xl.Cells(nfila, 4).Value = rsReporte.Fields("GlsEmpRelacionada")
            xl.Cells(nfila, 5).Value = Format(rsReporte.Fields("BaseImponible"), "0.00")
            xl.Cells(nfila, 6).Value = Format(rsReporte.Fields("totalPrecioVenta"), "0.00")
            nfila = nfila + 1
            rsReporte.MoveNext
        Loop
    End If
    
    xl.Cells(nfila, 4).Value = "TOTAL "
    xl.Cells(nfila, 5).Formula = "=SUM(" & "E" & 8 & ":" & "E" & nfila - 1 & ")"
    xl.Cells(nfila, 6).Formula = "=SUM(" & "F" & 8 & ":" & "F" & nfila - 1 & ")"
    xl.Cells(nfila, 4).Select
    xl.Selection.Font.Bold = True
    xl.Cells(nfila, 5).Select
    xl.Selection.Font.Bold = True
    xl.Cells(nfila, 6).Select
    xl.Selection.Font.Bold = True
    
    If rsReporte.State = 1 Then rsReporte.Close: Set rsReporte = Nothing
    
    Me.MousePointer = 1
    xl.Visible = True
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
