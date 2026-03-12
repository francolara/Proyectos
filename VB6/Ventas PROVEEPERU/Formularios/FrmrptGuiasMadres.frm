VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmrptGuiasMadres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte - Guías Madres"
   ClientHeight    =   5175
   ClientLeft      =   5595
   ClientTop       =   3210
   ClientWidth     =   6900
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
   ScaleHeight     =   5175
   ScaleWidth      =   6900
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4635
      Width           =   1200
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4635
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4485
      Left            =   45
      TabIndex        =   10
      Top             =   45
      Width           =   6765
      Begin VB.CommandButton CmdAyudaProducto 
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
         Left            =   6165
         Picture         =   "FrmrptGuiasMadres.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2610
         Width           =   390
      End
      Begin VB.CheckBox ChkPendientes 
         Caption         =   "Sólo Pendientes de Importar"
         Height          =   270
         Left            =   225
         TabIndex        =   7
         Top             =   4080
         Width           =   2580
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   0
         Left            =   180
         TabIndex        =   22
         Top             =   3060
         Width           =   6375
         Begin VB.OptionButton OptResumido 
            Caption         =   "Resumido"
            Height          =   240
            Left            =   4050
            TabIndex        =   6
            Top             =   360
            Width           =   1275
         End
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado"
            Height          =   240
            Left            =   1125
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1320
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   885
         Left            =   180
         TabIndex        =   20
         Top             =   180
         Width           =   6375
         Begin VB.ComboBox CboTipoReporte 
            Height          =   330
            ItemData        =   "FrmrptGuiasMadres.frx":038A
            Left            =   1710
            List            =   "FrmrptGuiasMadres.frx":038C
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   315
            Width           =   4200
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Reporte"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   405
            TabIndex        =   21
            Top             =   360
            Width           =   1140
         End
      End
      Begin VB.CommandButton CmdAyudaUnidProduc2 
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
         Left            =   6165
         Picture         =   "FrmrptGuiasMadres.frx":038E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   390
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
         Left            =   6165
         Picture         =   "FrmrptGuiasMadres.frx":0718
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2160
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   1215
         Width           =   6375
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   1
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   107085825
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
            TabIndex        =   2
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   107085825
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   13
            Top             =   375
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3960
            TabIndex        =   12
            Top             =   375
            Width           =   420
         End
      End
      Begin CATControls.CATTextBox txtCod_UnidProd 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Tag             =   "Tidupp"
         Top             =   2160
         Visible         =   0   'False
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
         Container       =   "FrmrptGuiasMadres.frx":0AA2
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_UnidProd 
         Height          =   315
         Left            =   2025
         TabIndex        =   15
         Top             =   2160
         Visible         =   0   'False
         Width           =   4125
         _ExtentX        =   7276
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
         Container       =   "FrmrptGuiasMadres.frx":0ABE
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_UnidProd2 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Tag             =   "Tidupp"
         Top             =   2160
         Visible         =   0   'False
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
         Container       =   "FrmrptGuiasMadres.frx":0ADA
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_UnidProd2 
         Height          =   315
         Left            =   2025
         TabIndex        =   18
         Top             =   2160
         Visible         =   0   'False
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         BackColor       =   12648447
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
         Container       =   "FrmrptGuiasMadres.frx":0AF6
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox TxtCodProducto 
         Height          =   315
         Left            =   1080
         TabIndex        =   24
         Tag             =   "Tidupp"
         Top             =   2610
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
         Container       =   "FrmrptGuiasMadres.frx":0B12
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox TxtGlsProducto 
         Height          =   315
         Left            =   2025
         TabIndex        =   25
         Top             =   2610
         Width           =   4125
         _ExtentX        =   7276
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
         Container       =   "FrmrptGuiasMadres.frx":0B2E
         Vacio           =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   225
         TabIndex        =   26
         Top             =   2655
         Width           =   645
      End
      Begin VB.Label lblCamal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Camal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   225
         TabIndex        =   19
         Top             =   2205
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblGranja 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Granja"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   225
         TabIndex        =   16
         Top             =   2205
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmrptGuiasMadres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CboTipoReporte_Click()
    
    If right(CboTipoReporte.Text, 3) = "001" Then
        lblGranja.Visible = True
        txtCod_UnidProd.Visible = True
        txtGls_UnidProd.Visible = True
        CmdAyudaUnidProduc.Visible = True
        
        lblCamal.Visible = False
        txtCod_UnidProd2.Visible = False
        txtGls_UnidProd2.Visible = False
        CmdAyudaUnidProduc2.Visible = False
    Else
        lblCamal.Visible = True
        txtCod_UnidProd2.Visible = True
        txtGls_UnidProd2.Visible = True
        CmdAyudaUnidProduc2.Visible = True
        
        lblGranja.Visible = False
        txtCod_UnidProd.Visible = False
        txtGls_UnidProd.Visible = False
        CmdAyudaUnidProduc.Visible = False
    End If
    
End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim fIni As String
Dim Ffin As String
Dim CodMoneda As String

    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    If right(CboTipoReporte.Text, 3) = "001" Then
        If OptDetallado = True Then
            mostrarReporte "rptGuiasMadresporGranjaDetallado.Rpt", "ParEmpresa|ParFecInicio|ParFecFinal|ParGranja|ParCamal|ParPendiente|ParProducto", glsEmpresa & "|" & fIni & "|" & Ffin & "|" & Trim(txtCod_UnidProd.Text) & "|" & Trim(txtCod_UnidProd2.Text) & "|" & ChkPendientes.Value & "|" & TxtCodProducto.Text, "Detallado por Granja", StrMsgError
        Else
            mostrarReporte "rptGuiasMadresporGranjaResumido.Rpt", "ParEmpresa|ParFecInicio|ParFecFinal|ParGranja|ParCamal|ParPendiente|ParProducto", glsEmpresa & "|" & fIni & "|" & Ffin & "|" & Trim(txtCod_UnidProd.Text) & "|" & Trim(txtCod_UnidProd2.Text) & "|" & ChkPendientes.Value & "|" & TxtCodProducto.Text, "Resumido por Granja", StrMsgError
        End If
    ElseIf right(CboTipoReporte.Text, 3) = "002" Then
        If OptDetallado = True Then
            mostrarReporte "rptGuiasMadresporCamalDetallado.Rpt", "ParEmpresa|ParFecInicio|ParFecFinal|ParGranja|ParCamal|ParPendiente|ParProducto", glsEmpresa & "|" & fIni & "|" & Ffin & "|" & Trim(txtCod_UnidProd.Text) & "|" & Trim(txtCod_UnidProd2.Text) & "|" & ChkPendientes.Value & "|" & TxtCodProducto.Text, "Detallado por Camal", StrMsgError
        Else
            mostrarReporte "rptGuiasMadresporCamalResumido.Rpt", "ParEmpresa|ParFecInicio|ParFecFinal|ParGranja|ParCamal|ParPendiente|ParProducto", glsEmpresa & "|" & fIni & "|" & Ffin & "|" & Trim(txtCod_UnidProd.Text) & "|" & Trim(txtCod_UnidProd2.Text) & "|" & ChkPendientes.Value & "|" & TxtCodProducto.Text, "Resumido por Camal", StrMsgError
        End If
    End If
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaProducto_Click()
On Error GoTo Err
Dim StrMsgError                     As String
    
    mostrarAyuda "PRODUCTOS", TxtCodProducto, TxtGlsProducto
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaUnidProduc_Click()
    
    mostrarAyuda "UNIDADPRODUC", txtCod_UnidProd, txtGls_UnidProd
    If txtCod_UnidProd.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub CmdAyudaUnidProduc2_Click()
    
    mostrarAyuda "UNIDADPRODUC", txtCod_UnidProd2, txtGls_UnidProd2
    If txtCod_UnidProd2.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    Me.top = 0
    Me.left = 0
    txtGls_UnidProd.Text = "TODAS LAS GRANJAS"
    txtGls_UnidProd2.Text = "TODOS LOS CAMALES"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    
    CboTipoReporte.AddItem "Detallado y Resumido por Granja" & Space(150) & "001"
    CboTipoReporte.AddItem "Detallado y Resumido por Camal" & Space(150) & "002"
    CboTipoReporte.ListIndex = 0
    
    lblGranja.Visible = True
    txtCod_UnidProd.Visible = True
    txtGls_UnidProd.Visible = True
    CmdAyudaUnidProduc.Visible = True
    OptDetallado.Value = True
    
End Sub

Private Sub txtCod_UnidProd_Change()
    
    If txtCod_UnidProd.Text = "" Then
        txtGls_UnidProd.Text = "TODAS LAS GRANJAS"
    Else
        txtGls_UnidProd.Text = traerCampo("unidadproduccion", "Descunidad", "CodUnidProd", txtCod_UnidProd.Text, True)
    End If

End Sub

Private Sub txtCod_UnidProd2_Change()
    
    If txtCod_UnidProd2.Text = "" Then
        txtGls_UnidProd2.Text = "TODOS LOS CAMALES"
    Else
        txtGls_UnidProd2.Text = traerCampo("unidadproduccion", "Descunidad", "CodUnidProd", txtCod_UnidProd2.Text, True)
    End If

End Sub

Private Sub TxtCodProducto_Change()
On Error GoTo Err
Dim StrMsgError                     As String
    
    If Len(Trim("" & TxtCodProducto.Text)) > 0 Then
        
        TxtGlsProducto.Text = traerCampo("Productos", "GlsProducto", "IdProducto", TxtCodProducto.Text, True)
        
    Else
        
        TxtGlsProducto.Text = "TODOS LOS PRODUCTOS"
        
    End If
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
